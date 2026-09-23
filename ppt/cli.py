"""
Command Line Interface for Consulting PPT Generator.
"""

import argparse
import sys
from pathlib import Path
from ppt.parser.builder import DeckBuilder
from ppt.config.default_theme import THEMES


def main():
    parser = argparse.ArgumentParser(
        description="Generate professional consulting PPT presentations using python-pptx."
    )
    parser.add_argument(
        "-s", "--spec",
        required=False,
        help="Path to YAML or JSON presentation spec file.",
    )
    parser.add_argument(
        "-o", "--out",
        default="output.pptx",
        help="Path for generated PPTX output file (default: output.pptx).",
    )
    parser.add_argument(
        "--list-themes",
        action="store_true",
        help="List all supported consulting themes.",
    )

    args = parser.parse_args()

    if args.list_themes:
        print("Available Consulting Themes:")
        for t in THEMES.keys():
            print(f" - {t}")
        sys.exit(0)

    if not args.spec:
        print("Error: -s/--spec is required when generating presentation.", file=sys.stderr)
        parser.print_help()
        sys.exit(1)

    spec_path = Path(args.spec)
    if not spec_path.exists():
        print(f"Error: Spec file not found: {args.spec}", file=sys.stderr)
        sys.exit(1)

    print(f"Loading specification from: {spec_path}")
    builder = DeckBuilder(spec_path)
    print(f"Compiling deck '{builder.spec.title}' with {len(builder.spec.slides)} slides...")
    builder.build(args.out)
    print(f"Successfully generated: {args.out}")


if __name__ == "__main__":
    main()
