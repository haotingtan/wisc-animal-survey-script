#!/usr/bin/env python3
"""Create a Markdown inventory report for one or more Excel workbooks.

By default, this summarizes the two survey workbooks supplied with this
project. Pass other .xlsx files with --inputs to reuse the script.
"""

from __future__ import annotations

import argparse
from datetime import date, datetime, time
from pathlib import Path
from typing import Any, Iterable

from openpyxl import load_workbook


DEFAULT_INPUTS = [
    Path(
        "/Users/haotingtan/Desktop/Research/1 Vet Survey/Spring - week 1/"
        "Anonymous vet responses 200128.xlsx"
    ),
    Path(
        "/Users/haotingtan/Desktop/Research/2 Farmer survey/Spring - week 5/"
        "Survey data USA text version for Haoting 250203.xlsx"
    ),
]


def cell_text(value: Any, limit: int = 140) -> str:
    """Render a cell value compactly for the report, without including data rows."""
    if value is None:
        return ""
    if isinstance(value, (datetime, date, time)):
        text = value.isoformat(sep=" ") if isinstance(value, datetime) else value.isoformat()
    else:
        text = str(value)
    text = " ".join(text.split())
    if len(text) > limit:
        text = text[: limit - 1] + "…"
    return text.replace("|", "\\|")


def nonempty(value: Any) -> bool:
    return value is not None and (not isinstance(value, str) or value.strip() != "")


def analyze_sheet(rows: Iterable[tuple[Any, ...]]) -> dict[str, Any]:
    populated_rows = 0
    populated_cells = 0
    max_populated_col = 0
    first_rows: list[list[str]] = []
    for row in rows:
        filled_columns = [index for index, value in enumerate(row, start=1) if nonempty(value)]
        if filled_columns:
            populated_rows += 1
            populated_cells += len(filled_columns)
            max_populated_col = max(max_populated_col, max(filled_columns))
        if len(first_rows) < 3:
            first_rows.append([cell_text(value) for value in row if nonempty(value)][:12])
    return {
        "populated_rows": populated_rows,
        "populated_cells": populated_cells,
        "populated_columns": max_populated_col,
        "preview": first_rows,
    }


def analyze_workbook(path: Path, data_only: bool) -> list[dict[str, Any]]:
    workbook = load_workbook(path, read_only=True, data_only=data_only)
    result = []
    try:
        for sheet in workbook.worksheets:
            metrics = analyze_sheet(sheet.iter_rows(values_only=True))
            result.append({"name": sheet.title, **metrics})
    finally:
        workbook.close()
    return result


def build_report(inputs: list[Path], data_only: bool) -> str:
    lines = ["# Excel workbook summary", ""]
    for path in inputs:
        if not path.is_file():
            raise FileNotFoundError(f"Input workbook not found: {path}")
        lines.extend([f"## {path.name}", "", f"Source: `{path}`", ""])
        for sheet in analyze_workbook(path, data_only):
            lines.extend(
                [
                    f"### {sheet['name']}",
                    "",
                    f"- Populated rows: {sheet['populated_rows']:,}",
                    f"- Populated columns: {sheet['populated_columns']:,}",
                    f"- Non-empty cells: {sheet['populated_cells']:,}",
                    "- First three rows (up to 12 non-empty cells per row; typically field names, question labels, and metadata):",
                    "",
                ]
            )
            if not any(any(cell for cell in row) for row in sheet["preview"]):
                lines.extend(["  _No values found._", ""])
            else:
                for row_number, row in enumerate(sheet["preview"], start=1):
                    values = " | ".join(row)
                    lines.extend([f"  - Row {row_number}: {values or '_blank_'}", ""])
    return "\n".join(lines)


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Summarize the worksheet structure of Excel workbooks."
    )
    parser.add_argument(
        "--inputs",
        nargs="+",
        type=Path,
        default=DEFAULT_INPUTS,
        metavar="FILE.xlsx",
        help="Excel files to inspect (defaults to the two survey workbooks).",
    )
    parser.add_argument(
        "--output",
        type=Path,
        default=Path("generated_summary.md"),
        help="Markdown report path (default: generated_summary.md).",
    )
    parser.add_argument(
        "--formula-text",
        action="store_true",
        help="Show formula text instead of cached formula results, when available.",
    )
    args = parser.parse_args()

    report = build_report(args.inputs, data_only=not args.formula_text)
    args.output.parent.mkdir(parents=True, exist_ok=True)
    args.output.write_text(report, encoding="utf-8")
    print(f"Wrote workbook summary to {args.output.resolve()}")


if __name__ == "__main__":
    main()
