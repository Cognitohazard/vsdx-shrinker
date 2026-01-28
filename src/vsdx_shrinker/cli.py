"""Command-line interface for vsdx-shrinker."""

import argparse
import sys

from .core import shrink_vsdx, analyze_vsdx, VsdxFormatError

MAX_DISPLAY_ITEMS = 10


def main() -> int:
    """Main entry point for the CLI."""
    parser = argparse.ArgumentParser(
        prog='vsdx-shrinker',
        description='Remove unused master shapes from Visio (.vsdx) files to reduce file size.',
        epilog='Example: vsdx-shrinker diagram.vsdx -o diagram_small.vsdx',
    )

    parser.add_argument('input', help='Input .vsdx file path')
    parser.add_argument('-o', '--output', default=None,
                        help='Output file path (default: overwrite input with backup)')
    parser.add_argument('--no-backup', action='store_true',
                        help='Do not create backup when overwriting input file')
    parser.add_argument('--analyze', action='store_true',
                        help='Only analyze the file, do not modify')
    parser.add_argument('-q', '--quiet', action='store_true',
                        help='Suppress output except errors')
    parser.add_argument('--version', action='version', version='%(prog)s 0.1.0')

    args = parser.parse_args()

    try:
        if args.analyze:
            result = analyze_vsdx(args.input)
            if not args.quiet:
                _print_analysis(args.input, result)
        else:
            result = shrink_vsdx(args.input, args.output, backup=not args.no_backup)
            if not args.quiet:
                _print_shrink_result(args.input, result)
        return 0

    except VsdxFormatError as e:
        print(f"Format error: {e}", file=sys.stderr)
        return 2
    except (FileNotFoundError, ValueError) as e:
        print(f"Error: {e}", file=sys.stderr)
        return 1
    except Exception as e:
        print(f"Error: {e}", file=sys.stderr)
        return 1


def _print_analysis(input_path: str, result: dict) -> None:
    """Print analysis results."""
    print(f"Analysis of: {input_path}")
    print(f"  Total masters:      {result['total_masters']}")
    print(f"  Used masters:       {result['used_masters']}")
    print(f"  Unused masters:     {result['unused_masters']}")
    print(f"  Potential savings:  {result['potential_savings_mb']} MB")

    unused = result['unused_names']
    if unused:
        print(f"\nUnused masters ({len(unused)}):")
        for name in unused[:MAX_DISPLAY_ITEMS]:
            print(f"    {name}")
        if len(unused) > MAX_DISPLAY_ITEMS:
            print(f"    ... and {len(unused) - MAX_DISPLAY_ITEMS} more")


def _print_shrink_result(input_path: str, result: dict) -> None:
    """Print shrink operation results."""
    print(f"Shrunk: {input_path}")
    print(f"  Original size:    {result['original_size_mb']} MB")
    print(f"  New size:         {result['new_size_mb']} MB")
    print(f"  Reduction:        {result['reduction_mb']} MB ({result['reduction_percent']}%)")
    print(f"  Masters removed:  {result['masters_removed']}")
    print(f"  Output:           {result['output_path']}")


if __name__ == '__main__':
    sys.exit(main())
