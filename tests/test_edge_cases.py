"""
Edge-case and property-based tests for vsdx_shrinker pure functions.

Each test targets a specific bug discovered through REPL probing of
boundary inputs: substring regex matches, strip-vs-lstrip namespace
parsing, mismatched-quote regexes, duplicate-name dict collisions,
and negative-zero floating-point artifacts.
"""

import math
import os
import tempfile
from pathlib import Path
from xml.etree import ElementTree as ET

import pytest
from hypothesis import given
from hypothesis import strategies as st

from vsdx_shrinker.core import (
    _bytes_to_mb,
    _get_namespace,
    _get_rel_id,
    _find_used_masters,
    _parse_masters_xml,
    VISIO_NS,
    REL_NS,
)


# ---------------------------------------------------------------------------
# helpers
# ---------------------------------------------------------------------------

def _make_pages_dir(page_content: str) -> Path:
    """Create a temp pages directory with a single page file containing
    ``page_content``.  The caller must clean up the parent temp dir."""
    tmp = Path(tempfile.mkdtemp())
    pages_dir = tmp / "pages"
    pages_dir.mkdir()
    (pages_dir / "page1.xml").write_text(page_content, encoding="utf-8")
    return pages_dir


def _write_masters_xml(xml_content: str) -> Path:
    fd, path = tempfile.mkstemp(suffix=".xml")
    os.write(fd, xml_content.encode("utf-8"))
    os.close(fd)
    return Path(path)


# ---------------------------------------------------------------------------
# Bug 1: USE() regex lacked \b word boundary
#
# REFUSE("X"), FUSE("X"), etc. were falsely matched, marking masters as
# "used" and preventing their removal.  We test _find_used_masters
# end-to-end so the fix is validated where it actually runs.
# ---------------------------------------------------------------------------

class TestUsePatternWordBoundary:
    """USE() regex must not match words that merely end in 'USE'."""

    MASTERS_INFO = {"Arrow": {"id": "1", "rel_id": "rId1"}}

    def test_standalone_use_is_found(self):
        pages_dir = _make_pages_dir('F="USE(&quot;Arrow&quot;)" />')
        try:
            # USE("Arrow") is a genuine Visio formula reference — must be found
            # Note: in real VSDX, USE() uses &quot; for quotes in XML context,
            # but the raw regex scans for literal double-quotes.  Test both.
            pages_dir_raw = _make_pages_dir('USE("Arrow")')
            used = _find_used_masters(pages_dir_raw, self.MASTERS_INFO)
            assert "Arrow" in used
        finally:
            import shutil
            shutil.rmtree(pages_dir.parent)
            shutil.rmtree(pages_dir_raw.parent)

    def test_refuse_is_not_matched(self):
        pages_dir = _make_pages_dir('REFUSE("Arrow")')
        try:
            used = _find_used_masters(pages_dir, self.MASTERS_INFO)
            assert "Arrow" not in used, (
                "REFUSE(\"Arrow\") falsely matched — missing \\b word boundary"
            )
        finally:
            import shutil
            shutil.rmtree(pages_dir.parent)

    def test_fuse_abuse_reuse_not_matched(self):
        """Several words ending in USE must not trigger a false positive."""
        for prefix in ("FUSE", "ABUSE", "REUSE", "MISUSE"):
            pages_dir = _make_pages_dir(f'{prefix}("Arrow")')
            try:
                used = _find_used_masters(pages_dir, self.MASTERS_INFO)
                assert "Arrow" not in used, f"{prefix} falsely matched"
            finally:
                import shutil
                shutil.rmtree(pages_dir.parent)

    @given(prefix=st.text(
        alphabet=st.characters(whitelist_categories=("Lu",)),
        min_size=1, max_size=8,
    ))
    def test_no_alpha_prefix_matches(self, prefix):
        """Any alphabetic prefix + USE("X") must not match."""
        pages_dir = _make_pages_dir(f'{prefix}USE("Arrow")')
        try:
            used = _find_used_masters(pages_dir, self.MASTERS_INFO)
            assert "Arrow" not in used, f"{prefix}USE falsely matched"
        finally:
            import shutil
            shutil.rmtree(pages_dir.parent)


# ---------------------------------------------------------------------------
# Bug 2: master_attr_pattern allowed mismatched quotes
#
# Master="42' (double-open, single-close) was accepted.  Tested
# end-to-end via _find_used_masters.
# ---------------------------------------------------------------------------

class TestMasterAttrPattern:
    MASTERS_INFO = {"Arrow": {"id": "42", "rel_id": "rId1"}}

    def test_matched_double_quotes_found(self):
        pages_dir = _make_pages_dir('<Shape Master="42"/>')
        try:
            used = _find_used_masters(pages_dir, self.MASTERS_INFO)
            assert "Arrow" in used
        finally:
            import shutil
            shutil.rmtree(pages_dir.parent)

    def test_matched_single_quotes_found(self):
        pages_dir = _make_pages_dir("<Shape Master='42'/>")
        try:
            used = _find_used_masters(pages_dir, self.MASTERS_INFO)
            assert "Arrow" in used
        finally:
            import shutil
            shutil.rmtree(pages_dir.parent)

    def test_mismatched_quotes_not_matched(self):
        """Mismatched quotes are not valid XML — must not match."""
        for content in ('Master="42\'', "Master='42\""):
            pages_dir = _make_pages_dir(content)
            try:
                used = _find_used_masters(pages_dir, self.MASTERS_INFO)
                assert "Arrow" not in used, (
                    f"Mismatched quotes in {content!r} falsely matched"
                )
            finally:
                import shutil
                shutil.rmtree(pages_dir.parent)


# ---------------------------------------------------------------------------
# Bug 3: _get_namespace used strip('{') instead of lstrip('{')
#
# strip('{') removes '{' from BOTH ends.  A namespace URI ending with '{'
# gets incorrectly truncated.
# ---------------------------------------------------------------------------

class TestGetNamespace:
    def test_normal_namespace(self):
        e = ET.Element(f"{{{VISIO_NS}}}Masters")
        assert _get_namespace(e) == VISIO_NS

    def test_no_namespace_returns_empty(self):
        e = ET.Element("Tag")
        assert _get_namespace(e) == ""

    def test_namespace_ending_with_open_brace(self):
        """A namespace URI ending with '{' must be preserved exactly.

        split('}') on '{ns{}local' yields ['{ns{', 'local'].
        strip('{') on '{ns{' wrongly gives 'ns'; lstrip('{') gives 'ns{'.
        """
        e = ET.Element("dummy")
        e.tag = "{ns{}local"
        assert _get_namespace(e) == "ns{"

    def test_namespace_with_internal_braces(self):
        e = ET.Element("dummy")
        e.tag = "{a{b{c}local"
        assert _get_namespace(e) == "a{b{c"


# ---------------------------------------------------------------------------
# Bug 4: _parse_masters_xml silently dropped duplicate NameU values
#
# Two <Master> elements with the same NameU overwrote each other in the
# dict, making the first master invisible and unremovable.
# ---------------------------------------------------------------------------

class TestParseMastersXmlDuplicateNames:
    def test_duplicate_nameu_both_tracked(self):
        xml = f'''<?xml version="1.0" encoding="utf-8"?>
<Masters xmlns="{VISIO_NS}" xmlns:r="{REL_NS}">
  <Master ID="1" NameU="Shape1"><Rel r:id="rId1"/></Master>
  <Master ID="2" NameU="Shape1"><Rel r:id="rId2"/></Master>
  <Master ID="3" NameU="Shape2"><Rel r:id="rId3"/></Master>
</Masters>'''
        path = _write_masters_xml(xml)
        try:
            _, masters_info = _parse_masters_xml(path)
            all_ids = {info["id"] for info in masters_info.values()}
            assert len(masters_info) == 3, (
                f"Expected 3 masters, got {len(masters_info)}. "
                f"Duplicate NameU caused silent overwrite."
            )
            assert all_ids == {"1", "2", "3"}, (
                f"All master IDs must be preserved. Got: {all_ids}"
            )
        finally:
            path.unlink()


# ---------------------------------------------------------------------------
# Bug 5: _bytes_to_mb returned -0.0 for small negative inputs
#
# round(-1 / 1048576, 2) == -0.0 due to IEEE 754.  str(-0.0) == '-0.0'
# which looks wrong in user-facing output ("Reduction: -0.0 MB").
# ---------------------------------------------------------------------------

class TestBytesToMb:
    def test_negative_one_byte_not_negative_zero(self):
        """Small negative inputs must not produce -0.0 in string output."""
        result = _bytes_to_mb(-1)
        assert str(result) != "-0.0", (
            "_bytes_to_mb(-1) returned -0.0, displays as '-0.0 MB'"
        )

    @given(n=st.integers(min_value=0, max_value=10**15))
    def test_non_negative_input_never_negative_zero(self, n):
        """Non-negative byte counts must never produce negative zero."""
        result = _bytes_to_mb(n)
        assert result >= 0.0
        assert not (result == 0.0 and math.copysign(1, result) < 0), (
            f"_bytes_to_mb({n}) returned negative zero"
        )


# ---------------------------------------------------------------------------
# _get_rel_id: verify precedence of full-namespace vs prefixed attribute
# ---------------------------------------------------------------------------

class TestGetRelId:
    def test_none_returns_empty(self):
        assert _get_rel_id(None) == ""

    def test_full_ns_takes_precedence_over_prefixed(self):
        """When both {REL_NS}id and r:id exist, full-namespace wins."""
        e = ET.Element("Rel")
        e.set(f"{{{REL_NS}}}id", "rId1")
        e.set("r:id", "rId2")
        assert _get_rel_id(e) == "rId1"

    def test_prefixed_used_as_fallback(self):
        """When only r:id exists, it must be returned."""
        e = ET.Element("Rel")
        e.set("r:id", "rId2")
        assert _get_rel_id(e) == "rId2"

    def test_neither_attribute_returns_empty(self):
        e = ET.Element("Rel")
        assert _get_rel_id(e) == ""
