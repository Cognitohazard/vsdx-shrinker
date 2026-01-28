"""
vsdx-shrinker: Remove unused master shapes from Visio files to reduce file size.
"""

from .core import shrink_vsdx, analyze_vsdx, VsdxFormatError

__version__ = "0.1.0"
__all__ = ["shrink_vsdx", "analyze_vsdx", "VsdxFormatError"]
