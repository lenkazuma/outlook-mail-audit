"""Shared, dependency-free helpers for local report generation."""
from __future__ import annotations
import argparse
import csv
from pathlib import Path
from typing import Any


def positive_int(value: str) -> int:
    number = int(value)
    if number < 1:
        raise argparse.ArgumentTypeError("must be a positive integer")
    return number


def csv_safe(value: Any) -> Any:
    """Prevent text cells from being treated as spreadsheet formulas."""
    if isinstance(value, str):
        candidate = value.lstrip()
        if value.startswith(("\t", "\r", "\n")) or candidate.startswith(("=", "+", "-", "@")):
            return "'" + value
    return value


def write_csv(path: Path, rows: list[dict[str, Any]]) -> None:
    if not rows:
        path.write_text("", encoding="utf-8")
        return
    fields = list(dict.fromkeys(key for row in rows for key in row))
    with path.open("w", encoding="utf-8-sig", newline="") as stream:
        writer = csv.DictWriter(stream, fieldnames=fields)
        writer.writerow({key: csv_safe(key) for key in fields})
        writer.writerows({key: csv_safe(value) for key, value in row.items()} for row in rows)
