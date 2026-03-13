#!/usr/bin/env python
"""Batch Find-and-Tag CLI — find equipment tags in a drawing and place callout groups.

Usage:
    python batch_tag.py "FT-575;BS-564;FT-577"
    python batch_tag.py tags.txt
    python batch_tag.py tags.txt --cube-size 12 --direction left --text-height 6

The tag list can be:
  - A semicolon-delimited string on the command line
  - A text file with one tag per line
"""

import argparse
import asyncio
import os
import sys

# Ensure the autocad-mcp package is importable
sys.path.insert(0, os.path.join(os.path.dirname(__file__), "src"))

from autocad_mcp.backends.file_ipc import FileIPCBackend


def parse_tags(arg: str) -> list[str]:
    """Parse tags from a file path or semicolon-delimited string."""
    if os.path.isfile(arg):
        with open(arg, "r") as f:
            return [line.strip() for line in f if line.strip()]
    return [t.strip() for t in arg.split(";") if t.strip()]


async def main():
    parser = argparse.ArgumentParser(description="Batch find and tag equipment in AutoCAD")
    parser.add_argument("tags", help="Semicolon-delimited tag list or path to tags.txt file")
    parser.add_argument("--cube-size", type=float, default=24.0, help="Cube size (default: 24.0)")
    parser.add_argument("--direction", default="right", choices=["left", "right"],
                        help="Tag placement direction (default: right)")
    parser.add_argument("--text-height", type=float, default=8.0, help="Text height (default: 8.0)")
    args = parser.parse_args()

    tags = parse_tags(args.tags)
    if not tags:
        print("No tags provided.")
        sys.exit(1)

    print(f"Tags to process ({len(tags)}): {', '.join(tags)}")
    print(f"Parameters: cube_size={args.cube_size}, direction={args.direction}, text_height={args.text_height}")
    print()

    backend = FileIPCBackend()
    init_result = await backend.initialize()
    if not init_result.ok:
        print(f"Failed to connect to AutoCAD: {init_result.error}")
        sys.exit(1)
    print("Connected to AutoCAD.")

    print(f"Dispatching batch-find-and-tag for {len(tags)} tags...")
    result = await backend.batch_find_and_tag(
        tags,
        cube_size=args.cube_size,
        direction=args.direction,
        text_height=args.text_height,
    )

    if not result.ok:
        print(f"Error: {result.error}")
        sys.exit(1)

    payload = result.payload
    placed = payload.get("placed", 0)
    not_found = payload.get("not_found", [])
    results = payload.get("results", [])

    # Print summary table
    print()
    print(f"{'Tag':<20} {'Status':<12} {'Position':<30} {'Handles'}")
    print("-" * 90)
    for entry in results:
        tag = entry.get("tag", "?")
        status = entry.get("status", "?")
        pos = entry.get("position", "")
        if isinstance(pos, (list, tuple)):
            pos_str = f"({pos[0]:.1f}, {pos[1]:.1f})"
        elif isinstance(pos, dict):
            pos_str = f"({pos.get('x', 0):.1f}, {pos.get('y', 0):.1f})"
        else:
            pos_str = str(pos) if pos else "-"

        if status == "placed":
            handles = f"cube={entry.get('cube_handle', '?')}, leader={entry.get('leader_handle', '?')}"
        else:
            handles = "-"

        print(f"{tag:<20} {status:<12} {pos_str:<30} {handles}")

    print()
    print(f"Placed: {placed}/{len(tags)}")
    if not_found:
        print(f"Not found: {', '.join(not_found)}")


if __name__ == "__main__":
    asyncio.run(main())
