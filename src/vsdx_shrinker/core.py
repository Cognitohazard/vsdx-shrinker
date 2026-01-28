"""
Core functionality for shrinking Visio (.vsdx) files.

VSDX files are ZIP archives containing XML files. This module identifies and removes
unused master shapes that bloat file size, typically caused by copy-pasting from
PDFs or other vector sources.
"""

import re
import shutil
import tempfile
import zipfile
from dataclasses import dataclass
from pathlib import Path
from typing import Set, Dict, List, Optional
from xml.etree import ElementTree as ET


# XML namespaces used in VSDX files
VISIO_NS = 'http://schemas.microsoft.com/office/visio/2012/main'
REL_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
PKG_REL_NS = 'http://schemas.openxmlformats.org/package/2006/relationships'

NAMESPACES = {'v': VISIO_NS, 'r': REL_NS}

for prefix, uri in NAMESPACES.items():
    ET.register_namespace(prefix, uri)


class VsdxFormatError(Exception):
    """Raised when VSDX file format doesn't match expected structure."""
    pass


@dataclass
class VsdxPaths:
    """Paths to key directories and files within an extracted VSDX."""
    root: Path
    pages_dir: Path
    masters_dir: Path
    masters_xml: Path
    rels_path: Path


def _bytes_to_mb(size_bytes: int) -> float:
    """Convert bytes to megabytes, rounded to 2 decimal places."""
    return round(size_bytes / (1024 * 1024), 2)


def _validate_vsdx_path(path: str) -> Path:
    """Validate and convert a VSDX path string to Path object."""
    p = Path(path)
    if not p.exists():
        raise FileNotFoundError(f"File not found: {p}")
    if p.suffix.lower() != '.vsdx':
        raise ValueError(f"Not a .vsdx file: {p}")
    return p


def _get_namespace(element: ET.Element) -> str:
    """Extract namespace URI from an element's tag."""
    if '}' in element.tag:
        return element.tag.split('}')[0].strip('{')
    return ''


def _make_vsdx_paths(root: Path) -> VsdxPaths:
    """Create VsdxPaths from an extraction root directory."""
    masters_dir = root / 'visio' / 'masters'
    return VsdxPaths(
        root=root,
        pages_dir=root / 'visio' / 'pages',
        masters_dir=masters_dir,
        masters_xml=masters_dir / 'masters.xml',
        rels_path=masters_dir / '_rels' / 'masters.xml.rels',
    )


def _get_rel_id(rel_element: Optional[ET.Element]) -> str:
    """Extract relationship ID from a Rel element."""
    if rel_element is None:
        return ''
    return rel_element.get(f'{{{REL_NS}}}id', '') or rel_element.get('r:id', '')


def _read_xml_file(path: Path) -> str:
    """Read XML file content with encoding fallback."""
    try:
        return path.read_text(encoding='utf-8')
    except UnicodeDecodeError:
        return path.read_text(encoding='utf-8-sig')  # BOM variant


def _parse_xml_file(path: Path) -> ET.Element:
    """Parse XML file with error handling."""
    try:
        tree = ET.parse(path)
        return tree.getroot()
    except ET.ParseError as e:
        raise VsdxFormatError(f"Invalid XML in {path.name}: {e}")


def _get_page_files(pages_dir: Path) -> List[Path]:
    """Get page files, checking pages.xml relationships with fallback to glob."""
    page_files: List[Path] = []

    # Try to get pages from pages.xml relationships
    pages_rels = pages_dir / '_rels' / 'pages.xml.rels'
    if pages_rels.exists():
        try:
            root = _parse_xml_file(pages_rels)
            for rel in root.findall(f'.//{{{PKG_REL_NS}}}Relationship'):
                target = rel.get('Target', '')
                if target:
                    page_path = pages_dir / target
                    if page_path.exists():
                        page_files.append(page_path)
        except VsdxFormatError:
            pass  # Fall back to glob

    # Fallback to glob pattern if no pages found via rels
    if not page_files:
        page_files = list(pages_dir.glob("page*.xml"))

    return page_files


def _validate_vsdx_structure(paths: VsdxPaths) -> None:
    """Validate VSDX structure before processing. Raises VsdxFormatError if invalid."""
    if not paths.masters_xml.exists():
        return

    errors: List[str] = []

    # 1. Required files
    if not paths.rels_path.exists():
        errors.append(f"Missing relationships file: {paths.rels_path.name}")
    if not paths.pages_dir.exists():
        errors.append("Missing pages directory")

    # 2. Visio namespace validation
    root = _parse_xml_file(paths.masters_xml)
    ns = _get_namespace(root)
    if ns and ns != VISIO_NS:
        errors.append(f"Unexpected namespace: {ns}\n    Expected: {VISIO_NS}")

    # 3. Master element structure validation
    masters = root.findall('.//v:Master', NAMESPACES)
    if masters:
        sample = masters[0]
        required_attrs = {'ID', 'NameU'}
        missing = required_attrs - set(sample.attrib.keys())
        if missing:
            errors.append(f"Master elements missing required attributes: {missing}")

        # Check Rel child element exists and uses expected namespace
        rel = sample.find('.//v:Rel', NAMESPACES)
        if rel is None:
            errors.append("Master elements missing Rel child element")
        else:
            rel_id_full_ns = f'{{{REL_NS}}}id'
            has_full_ns = rel_id_full_ns in rel.attrib
            has_prefixed = 'r:id' in rel.attrib
            if not has_full_ns and not has_prefixed:
                id_attrs = [k for k in rel.attrib.keys() if k.endswith('}id') or k == 'id']
                if id_attrs:
                    errors.append(f"Rel element uses unexpected namespace for id: {id_attrs[0]}\n    Expected: {REL_NS}")
                else:
                    errors.append("Rel element missing r:id attribute")

    # 4. Relationships file namespace validation
    rels_ids: Set[str] = set()
    if paths.rels_path.exists():
        rels_root = _parse_xml_file(paths.rels_path)
        rels_ns = _get_namespace(rels_root)
        if rels_ns and rels_ns != PKG_REL_NS:
            errors.append(f"Unexpected relationships namespace: {rels_ns}\n    Expected: {PKG_REL_NS}")

        for rel in rels_root.findall(f'.//{{{PKG_REL_NS}}}Relationship'):
            rel_id = rel.get('Id', '')
            if rel_id:
                rels_ids.add(rel_id)

    # 5. Relationship integrity: every master's Rel r:id must exist in .rels
    if masters and rels_ids:
        for master in masters:
            rel = master.find('.//v:Rel', NAMESPACES)
            if rel is not None:
                rel_id = _get_rel_id(rel)
                if rel_id and rel_id not in rels_ids:
                    master_name = master.get('NameU', master.get('ID', 'unknown'))
                    errors.append(f"Master '{master_name}' references non-existent relationship: {rel_id}")

    if errors:
        raise VsdxFormatError(
            "VSDX format validation failed:\n  - " +
            "\n  - ".join(errors) +
            "\n\nThis file may use a newer or different format version."
        )


def _find_used_masters(pages_dir: Path, masters_info: Dict[str, Dict]) -> Set[str]:
    """Find masters referenced by USE() patterns OR Shape Master attributes."""
    used_names: Set[str] = set()

    # Build ID -> Name lookup for Master="ID" references
    id_to_name = {info['id']: name for name, info in masters_info.items()}

    # Patterns for both reference types
    use_pattern = re.compile(r'USE\("([^"]+)"\)')
    master_attr_pattern = re.compile(r'\bMaster=["\'](\d+)["\']')

    for page_file in _get_page_files(pages_dir):
        content = _read_xml_file(page_file)

        # Method 1: USE("name") patterns (formula inheritance)
        used_names.update(use_pattern.findall(content))

        # Method 2: Master="ID" attributes on shapes (instance relationship)
        for master_id in master_attr_pattern.findall(content):
            if master_id in id_to_name:
                used_names.add(id_to_name[master_id])

    return used_names


def _parse_masters_xml(masters_xml_path: Path) -> tuple[ET.Element, Dict[str, Dict]]:
    """Parse masters.xml and return (root element, {name: {id, rel_id, element}})."""
    tree = ET.parse(masters_xml_path)
    root = tree.getroot()

    masters_info: Dict[str, Dict] = {}
    for master in root.findall('.//v:Master', NAMESPACES):
        name = master.get('NameU', '')
        if name:
            rel = master.find('.//v:Rel', NAMESPACES)
            masters_info[name] = {
                'id': master.get('ID', ''),
                'rel_id': _get_rel_id(rel),
                'element': master,
            }

    return root, masters_info


def _parse_rels_xml(rels_path: Path) -> tuple[ET.Element, Dict[str, str]]:
    """Parse masters.xml.rels and return (root element, {rId: target_filename})."""
    tree = ET.parse(rels_path)
    root = tree.getroot()

    rels_info: Dict[str, str] = {}
    for rel in root.findall(f'.//{{{PKG_REL_NS}}}Relationship'):
        rel_id, target = rel.get('Id', ''), rel.get('Target', '')
        if rel_id and target:
            rels_info[rel_id] = target

    return root, rels_info


def _create_vsdx(source_dir: Path, output_path: Path) -> None:
    """Create a VSDX file from an extracted directory."""
    with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED) as zf:
        for file_path in source_dir.rglob('*'):
            if file_path.is_file():
                zf.write(file_path, file_path.relative_to(source_dir))


def _empty_result(original_size: int = 0, output_path: str = '') -> Dict:
    """Return a result dict for files with no masters to clean."""
    size_mb = _bytes_to_mb(original_size)
    return {
        'original_size_mb': size_mb,
        'new_size_mb': size_mb,
        'reduction_mb': 0.0,
        'reduction_percent': 0.0,
        'masters_removed': 0,
        'output_path': output_path,
    }


def _calculate_unused_size(
    unused_names: Set[str],
    masters_info: Dict[str, Dict],
    rels_info: Dict[str, str],
    masters_dir: Path,
) -> int:
    """Calculate total size of unused master files."""
    total = 0
    for name in unused_names:
        rel_id = masters_info.get(name, {}).get('rel_id', '')
        target = rels_info.get(rel_id, '')
        if target:
            target_file = masters_dir / target
            if target_file.exists():
                total += target_file.stat().st_size
    return total


def analyze_vsdx(vsdx_path: str) -> Dict:
    """
    Analyze a VSDX file and report on master shape usage.

    Args:
        vsdx_path: Path to the .vsdx file

    Returns:
        Dictionary with: total_masters, used_masters, unused_masters,
        used_names, unused_names, potential_savings_mb

    Raises:
        VsdxFormatError: If the file format doesn't match expected structure
    """
    path = _validate_vsdx_path(vsdx_path)

    with tempfile.TemporaryDirectory() as temp_dir:
        extract_dir = Path(temp_dir)
        with zipfile.ZipFile(path, 'r') as zf:
            zf.extractall(extract_dir)

        paths = _make_vsdx_paths(extract_dir)

        if not paths.masters_xml.exists():
            return {
                'total_masters': 0, 'used_masters': 0, 'unused_masters': 0,
                'used_names': [], 'unused_names': [], 'potential_savings_mb': 0.0,
            }

        _validate_vsdx_structure(paths)

        _, masters_info = _parse_masters_xml(paths.masters_xml)
        _, rels_info = _parse_rels_xml(paths.rels_path)

        used_names = _find_used_masters(paths.pages_dir, masters_info)
        all_names = set(masters_info.keys())
        unused_names = all_names - used_names

        unused_size = _calculate_unused_size(
            unused_names, masters_info, rels_info, paths.masters_dir
        )

        return {
            'total_masters': len(all_names),
            'used_masters': len(used_names & all_names),
            'unused_masters': len(unused_names),
            'used_names': sorted(used_names & all_names),
            'unused_names': sorted(unused_names),
            'potential_savings_mb': _bytes_to_mb(unused_size),
        }


def shrink_vsdx(
    input_path: str,
    output_path: Optional[str] = None,
    backup: bool = True,
) -> Dict:
    """
    Remove unused master shapes from a VSDX file.

    Args:
        input_path: Path to the input .vsdx file
        output_path: Path for output file. If None, overwrites input (with backup)
        backup: If True and output_path is None, create a .bak backup

    Returns:
        Dictionary with: original_size_mb, new_size_mb, reduction_mb,
        reduction_percent, masters_removed, output_path

    Raises:
        VsdxFormatError: If the file format doesn't match expected structure
    """
    path = _validate_vsdx_path(input_path)
    original_size = path.stat().st_size

    # Determine final output path
    if output_path is None:
        if backup:
            shutil.copy2(path, path.with_suffix('.vsdx.bak'))
        final_output = path
    else:
        final_output = Path(output_path)

    with tempfile.TemporaryDirectory() as temp_dir:
        temp_root = Path(temp_dir)
        extract_dir = temp_root / 'extracted'

        with zipfile.ZipFile(path, 'r') as zf:
            zf.extractall(extract_dir)

        paths = _make_vsdx_paths(extract_dir)

        if not paths.masters_xml.exists():
            if final_output != path:
                shutil.copy2(path, final_output)
            return _empty_result(original_size, str(final_output))

        # Validate structure before processing
        _validate_vsdx_structure(paths)

        # Parse masters and relationships
        masters_root, masters_info = _parse_masters_xml(paths.masters_xml)
        rels_root, rels_info = _parse_rels_xml(paths.rels_path)

        # Identify used vs unused masters (both USE() and Master="ID" references)
        used_names = _find_used_masters(paths.pages_dir, masters_info)

        all_names = set(masters_info.keys())
        names_to_remove = all_names - used_names

        # Determine what to keep
        keep_rel_ids: Set[str] = set()
        keep_files: Set[str] = set()
        for name in used_names:
            if name in masters_info:
                rel_id = masters_info[name]['rel_id']
                keep_rel_ids.add(rel_id)
                if rel_id in rels_info:
                    keep_files.add(rels_info[rel_id])

        # Remove unused masters from XML
        masters_removed = 0
        for name in names_to_remove:
            if name in masters_info:
                masters_root.remove(masters_info[name]['element'])
                masters_removed += 1

        ET.ElementTree(masters_root).write(
            paths.masters_xml, encoding='utf-8', xml_declaration=True
        )

        # Remove unused relationships
        for rel in list(rels_root):
            rel_id = rel.get('Id', '')
            if rel_id and rel_id not in keep_rel_ids:
                rels_root.remove(rel)

        ET.ElementTree(rels_root).write(
            paths.rels_path, encoding='utf-8', xml_declaration=True
        )

        # Delete unused master files
        for master_file in paths.masters_dir.glob('master*.xml'):
            if master_file.name != 'masters.xml' and master_file.name not in keep_files:
                master_file.unlink()

        # Create new VSDX
        temp_output = temp_root / 'output.vsdx'
        _create_vsdx(extract_dir, temp_output)
        shutil.move(str(temp_output), str(final_output))

    new_size = final_output.stat().st_size
    reduction = original_size - new_size

    return {
        'original_size_mb': _bytes_to_mb(original_size),
        'new_size_mb': _bytes_to_mb(new_size),
        'reduction_mb': _bytes_to_mb(reduction),
        'reduction_percent': round(reduction / original_size * 100, 1) if original_size else 0.0,
        'masters_removed': masters_removed,
        'output_path': str(final_output),
    }
