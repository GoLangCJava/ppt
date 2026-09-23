#!/usr/bin/env python3
"""
Demo Script: Generating Consulting Presentations via YAML and Python API.
"""

import sys
from pathlib import Path

# Add project root to sys.path
root_dir = Path(__file__).resolve().parent.parent
if str(root_dir) not in sys.path:
    sys.path.insert(0, str(root_dir))

from ppt.parser.builder import DeckBuilder


def main():
    yaml_path = root_dir / "examples" / "sample_deck.yaml"
    output_dir = root_dir / "examples" / "output"
    output_dir.mkdir(parents=True, exist_ok=True)
    output_file = output_dir / "consulting_solution_proposal.pptx"

    print("=" * 60)
    print("Consulting Slide Generator - Building Demo Presentation")
    print("=" * 60)
    print(f"Reading Spec: {yaml_path}")

    builder = DeckBuilder(yaml_path)
    print(f"Presentation Title: {builder.spec.title}")
    print(f"Theme: {builder.spec.theme}")
    print(f"Slide Count: {len(builder.spec.slides)}")

    print(f"Rendering slides to: {output_file} ...")
    prs = builder.build(output_file)

    print(f"✅ Successfully generated {len(prs.slides)} slides to {output_file}!")
    print("=" * 60)


if __name__ == "__main__":
    main()
