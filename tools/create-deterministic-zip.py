#!/usr/bin/env python3
"""Create a byte-stable ZIP archive from a directory tree."""

from __future__ import annotations

import argparse
import zipfile
from pathlib import Path


FIXED_TIMESTAMP = (2026, 7, 30, 0, 0, 0)


def create_archive(source: Path, output: Path) -> None:
    source = source.resolve(strict=True)
    output = output.resolve(strict=False)
    if not source.is_dir():
        raise ValueError(f"Source is not a directory: {source}")
    if not output.parent.is_dir():
        raise ValueError(f"Output directory does not exist: {output.parent}")

    files = sorted(
        (path for path in source.rglob("*") if path.is_file()),
        key=lambda path: path.relative_to(source).as_posix(),
    )
    if not files:
        raise ValueError(f"Source directory is empty: {source}")

    with zipfile.ZipFile(
        output,
        mode="w",
        compression=zipfile.ZIP_STORED,
        allowZip64=True,
        strict_timestamps=True,
    ) as archive:
        for path in files:
            relative_path = path.relative_to(source).as_posix()
            info = zipfile.ZipInfo(relative_path, date_time=FIXED_TIMESTAMP)
            info.compress_type = zipfile.ZIP_STORED
            info.create_system = 0
            info.external_attr = 0
            info.extra = b""
            info.comment = b""
            archive.writestr(info, path.read_bytes())


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--source", required=True, type=Path)
    parser.add_argument("--output", required=True, type=Path)
    args = parser.parse_args()
    create_archive(args.source, args.output)


if __name__ == "__main__":
    main()
