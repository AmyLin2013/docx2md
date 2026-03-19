# Copyright (c) 2025 AmyLin <zhi_lin@qq.com>
# Licensed under the MIT License. See LICENSE file for details.

"""CLI entry point for doc2md converter."""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

from converter.word2md import convert_word_to_markdown


def _auto_output_path(input_path: Path) -> Path:
    """Generate output .md path from input path."""
    return input_path.with_suffix(".md")


def _validate_format(input_path: Path) -> None:
    """Validate that the input file is a supported Word format."""
    ext = input_path.suffix.lower()
    if ext != ".docx":
        raise ValueError(f"Unsupported file format: {ext} (supported: .docx)")


def convert_file(input_path: str, output_path: str | None = None, **kwargs) -> str:
    """Convert a Word file to Markdown.

    Args:
        input_path: Path to input file (.docx).
        output_path: Optional output .md path. Auto-generated if None.
        **kwargs: Additional arguments passed to the converter.

    Returns:
        The output file path.
    """
    inp = Path(input_path)
    _validate_format(inp)
    out = Path(output_path) if output_path else _auto_output_path(inp)

    print(f"Converting: {inp.name} → {out.name}")
    convert_word_to_markdown(inp, out, **kwargs)
    print(f"  ✓ Saved to: {out}")
    return str(out)


def main():
    parser = argparse.ArgumentParser(
        prog="doc2md",
        description="Convert Word (.docx) files to Markdown.",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  doc2md document.docx                # → document.md
  doc2md document.docx -o output.md   # → output.md
  doc2md *.docx                       # batch convert all .docx files
  doc2md --no-images doc.docx         # skip image extraction
        """,
    )

    parser.add_argument(
        "input",
        nargs="+",
        help="Input file(s) to convert (.docx). Supports glob patterns.",
    )
    parser.add_argument(
        "-o", "--output",
        help="Output .md file path. Only valid with a single input file. "
             "If omitted, output is saved next to the input with .md extension.",
    )
    parser.add_argument(
        "--no-images",
        action="store_true",
        help="Do not extract embedded images to separate files.",
    )
    parser.add_argument(
        "--skip-cover",
        action="store_true",
        help="Remove the first page (cover page).",
    )
    parser.add_argument(
        "--toc-mode",
        choices=["none", "toc_only", "before_toc", "before_toc_keep_abstract"],
        default="none",
        help="How to handle the Table of Contents. "
             "none=keep all, toc_only=remove TOC only, "
             "before_toc=remove TOC+before, "
             "before_toc_keep_abstract=remove TOC+before but keep abstracts.",
    )
    parser.add_argument(
        "--stdout",
        action="store_true",
        help="Print Markdown to stdout instead of saving to file.",
    )

    args = parser.parse_args()

    # Validate
    if args.output and len(args.input) > 1:
        parser.error("--output can only be used with a single input file.")

    # Process each file
    success_count = 0
    error_count = 0

    for input_file in args.input:
        try:
            inp = Path(input_file)
            _validate_format(inp)

            if args.stdout:
                md = convert_word_to_markdown(
                    inp, extract_images=not args.no_images,
                    skip_cover=args.skip_cover,
                    toc_mode=args.toc_mode,
                )
                print(md)
            else:
                convert_file(
                    input_file, args.output if args.output else None,
                    extract_images=not args.no_images,
                    skip_cover=args.skip_cover,
                    toc_mode=args.toc_mode,
                )

            success_count += 1
        except Exception as e:
            print(f"  ✗ Error converting {input_file}: {e}", file=sys.stderr)
            error_count += 1

    # Summary
    if len(args.input) > 1:
        print(f"\nDone: {success_count} succeeded, {error_count} failed.")

    sys.exit(1 if error_count > 0 else 0)


if __name__ == "__main__":
    main()
