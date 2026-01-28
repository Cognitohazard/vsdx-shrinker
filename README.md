# vsdx-shrinker

Remove unused master shapes from Visio (.vsdx) files to dramatically reduce file size.

## The Problem

Visio files can become bloated when you copy-paste content from PDFs, vector graphics, or other sources. Each paste operation can create "master shapes" that remain in the file even after you delete the visible content. These orphaned masters can make your file 10x larger than necessary.

## How It Works

VSDX files are ZIP archives containing XML. This tool:

1. Extracts the VSDX file
2. Scans page content to find which masters are actually referenced (via `USE("...")` patterns)
3. Removes unreferenced masters from `masters.xml` and deletes their XML files
4. Repacks the cleaned archive

## Installation

```bash
pip install vsdx-shrinker
```

Or with uv:

```bash
uv tool install vsdx-shrinker
```

## Usage

### Command Line

Shrink a file (creates backup):

```bash
vsdx-shrinker diagram.vsdx
```

Shrink to a new file:

```bash
vsdx-shrinker diagram.vsdx -o diagram_small.vsdx
```

Analyze without modifying:

```bash
vsdx-shrinker diagram.vsdx --analyze
```

### Python API

```python
from vsdx_shrinker import shrink_vsdx, analyze_vsdx

# Analyze a file
result = analyze_vsdx("diagram.vsdx")
print(f"Can save {result['potential_savings_mb']} MB")

# Shrink a file
result = shrink_vsdx("diagram.vsdx", output_path="diagram_small.vsdx")
print(f"Reduced by {result['reduction_percent']}%")
```

## CLI Options

```
usage: vsdx-shrinker [-h] [-o OUTPUT] [--no-backup] [--analyze] [-q] [--version] input

positional arguments:
  input                 Input .vsdx file path

options:
  -h, --help            show this help message and exit
  -o, --output OUTPUT   Output file path (default: overwrite input with backup)
  --no-backup           Do not create backup when overwriting input file
  --analyze             Only analyze the file, do not modify
  -q, --quiet           Suppress output except errors
  --version             show program's version number and exit
```

## Example Output

```
$ vsdx-shrinker presentation.vsdx --analyze
Analysis of: presentation.vsdx
  Total masters:      244
  Used masters:       6
  Unused masters:     238
  Potential savings:  79.17 MB

$ vsdx-shrinker presentation.vsdx -o presentation_clean.vsdx
Shrunk: presentation.vsdx
  Original size:    14.8 MB
  New size:         1.6 MB
  Reduction:        13.2 MB (89.2%)
  Masters removed:  238
  Output:           presentation_clean.vsdx
```

## License

MIT
