"""Read actual XLSX/XML cell coordinates without trusting declared dimensions."""

import argparse
import json
import os
from pathlib import Path
from xml.etree import ElementTree as ET
from zipfile import ZipFile

from openpyxl.utils.cell import coordinate_to_tuple, range_boundaries

NS = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("folder")
    parser.add_argument("report", type=Path)
    args = parser.parse_args()
    summary = {"folder": args.folder, "workbooks": 0, "worksheets": 0,
               "cells_with_content": 0, "outside_declared_dimensions": 0,
               "formula_cells": 0, "formulas_without_saved_nonempty_result": 0,
               "dimension_issues": [], "read_errors": [], "scan_errors": []}
    for root, _, names in os.walk(args.folder, onerror=lambda exc: summary["scan_errors"].append(str(exc))):
        for name in names:
            if not name.lower().endswith((".xlsx", ".xlsm")) or name.startswith("~$"):
                continue
            path = os.path.join(root, name)
            try:
                with ZipFile(path) as archive:
                    summary["workbooks"] += 1
                    for part in archive.namelist():
                        if not part.startswith("xl/worksheets/") or not part.endswith(".xml") or "/_rels/" in part:
                            continue
                        summary["worksheets"] += 1
                        bounds = reference = None
                        outside = []
                        with archive.open(part) as source:
                            for _, node in ET.iterparse(source, events=("end",)):
                                if node.tag == NS + "dimension":
                                    reference = node.get("ref")
                                    try:
                                        bounds = range_boundaries(reference)
                                    except (ValueError, TypeError):
                                        bounds = None
                                elif node.tag == NS + "c":
                                    value = node.find(NS + "v")
                                    formula = node.find(NS + "f")
                                    inline = node.find(NS + "is")
                                    if value is not None or formula is not None or inline is not None:
                                        summary["cells_with_content"] += 1
                                        if bounds is not None and all(n is not None for n in bounds):
                                            row, column = coordinate_to_tuple(node.get("r"))
                                            if not (bounds[0] <= column <= bounds[2] and bounds[1] <= row <= bounds[3]):
                                                outside.append(node.get("r"))
                                    if formula is not None:
                                        summary["formula_cells"] += 1
                                        if value is None or value.text is None:
                                            summary["formulas_without_saved_nonempty_result"] += 1
                                    node.clear()
                                elif node.tag == NS + "row":
                                    node.clear()
                        if outside:
                            summary["outside_declared_dimensions"] += len(outside)
                            summary["dimension_issues"].append({
                                "file": path, "part": part, "declared": reference,
                                "count": len(outside), "sample_coordinates": outside[:10],
                            })
            except Exception as exc:
                summary["read_errors"].append({"file": path, "error": str(exc)})
    args.report.parent.mkdir(parents=True, exist_ok=True)
    args.report.write_text(json.dumps(summary, ensure_ascii=False, indent=2), encoding="utf-8")
    print(json.dumps({key: value for key, value in summary.items()
                      if key not in ("dimension_issues", "read_errors", "scan_errors")}, ensure_ascii=False))
    print("dimension_issue_sheets", len(summary["dimension_issues"]), "read_errors", len(summary["read_errors"]),
          "scan_errors", len(summary["scan_errors"]))


if __name__ == "__main__":
    main()
