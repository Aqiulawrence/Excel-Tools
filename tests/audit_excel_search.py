"""Read-only text search audit against independently read XML/BIFF text cells."""

import argparse
from collections import Counter
import hashlib
import json
import os
from pathlib import Path
import posixpath
import sys
from xml.etree import ElementTree as ET
from zipfile import ZipFile

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")
sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
import ExcelSearch as search
import xlrd

NS = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
RID = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"
TERMS = ("JOY", "USD", "品名")


def rich_text(node):
    # Phonetic guides are metadata, not part of the cell value.
    return "".join(child.text or "" for child in node.findall(NS + "t")) + "".join(
        child.text or "" for run in node.findall(NS + "r") for child in run.findall(NS + "t"))


def independent_text_cells(path):
    if path.lower().endswith(".xls"):
        book = xlrd.open_workbook(path, on_demand=True)
        try:
            for sheet in book.sheets():
                for row in range(sheet.nrows):
                    for col in range(sheet.ncols):
                        if sheet.cell_type(row, col) == xlrd.XL_CELL_TEXT:
                            # Independent address construction, without production helpers.
                            number, letters = col + 1, ""
                            while number:
                                number, digit = divmod(number - 1, 26)
                                letters = chr(65 + digit) + letters
                            yield sheet.name, f"{letters}{row + 1}", sheet.cell_value(row, col)
        finally:
            book.release_resources()
        return
    with ZipFile(path) as archive:
        shared = []
        if "xl/sharedStrings.xml" in archive.namelist():
            with archive.open("xl/sharedStrings.xml") as source:
                for _, node in ET.iterparse(source, events=("end",)):
                    if node.tag == NS + "si":
                        shared.append(rich_text(node).replace("x005F_", ""))
                        node.clear()
        relationships = {node.get("Id"): node.get("Target") for node in
                         ET.fromstring(archive.read("xl/_rels/workbook.xml.rels"))}
        sheets = ET.fromstring(archive.read("xl/workbook.xml")).find(NS + "sheets")
        for sheet in sheets:
            target = relationships[sheet.get(RID)]
            part = target.lstrip("/") if target.startswith("/") else posixpath.normpath("xl/" + target)
            if not part.startswith("xl/worksheets/"):
                continue
            with archive.open(part) as source:
                for _, node in ET.iterparse(source, events=("end",)):
                    if node.tag == NS + "c":
                        kind = node.get("t")
                        value = node.find(NS + "v")
                        text = None
                        if kind == "s" and value is not None and value.text is not None:
                            text = shared[int(value.text)]
                        elif kind == "inlineStr":
                            inline = node.find(NS + "is")
                            text = rich_text(inline) if inline is not None else ""
                        elif kind in ("str", "e") and value is not None:
                            text = value.text or ""
                        if text is not None:
                            yield sheet.get("name"), node.get("r"), text
                        node.clear()
                    elif node.tag == NS + "row":
                        node.clear()


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("folder")
    parser.add_argument("report", type=Path)
    parser.add_argument("--keywords", nargs="+", default=list(TERMS))
    parser.add_argument("--workers", type=int, nargs="+", default=[1, 12, 16])
    args = parser.parse_args()
    if not Path(args.folder).is_dir():
        parser.error("folder must be an existing directory")
    if any(not term.strip() for term in args.keywords):
        parser.error("keywords must not be empty")
    if any(count < 1 or count > 16 for count in args.workers):
        parser.error("workers must be between 1 and 16")
    terms = list(dict.fromkeys(term.strip() for term in args.keywords))
    scan_errors, files = [], []
    for root, _, names in os.walk(args.folder, onerror=lambda e: scan_errors.append(str(e))):
        files.extend(os.path.join(root, name) for name in names
                     if name.lower().endswith(search.EXCEL_EXTENSIONS) and not name.startswith("~$"))
    expected = {term: Counter() for term in terms}
    oracle_errors, text_cells = [], 0
    for path in files:
        try:
            for sheet, address, text in independent_text_cells(path):
                text_cells += 1
                lowered = text.lower()
                for term in terms:
                    if term.lower() in lowered:
                        expected[term][(path, sheet, address, text)] += 1
        except Exception as exc:
            oracle_errors.append({"file": path, "error": repr(exc)})
    print(json.dumps({"oracle_files": len(files), "text_cells": text_cells,
                      "oracle_errors": oracle_errors}, ensure_ascii=False), flush=True)
    app = search.QApplication([])
    app.setQuitOnLastWindowClosed(False)
    pool = search.SearchProcessPool()
    report = {"source_sha256": hashlib.sha256(Path(search.__file__).read_bytes()).hexdigest(),
              "folder": args.folder, "files": len(files), "oracle_text_cells": text_cells,
              "oracle_errors": oracle_errors, "scan_errors": scan_errors, "runs": []}
    try:
        # Changing pool size also checks rebuilding the persistent process pool.
        for index, term in enumerate(terms):
            workers = args.workers[index % len(args.workers)]
            worker = search.ExcelSearchWorker(pool)
            worker.setup(term, [args.folder], workers, search.DEFAULT_KEYWORDS)
            records, progress = [], []
            worker.batch_result_signal.connect(records.extend)
            worker.progress_signal.connect(lambda *args: progress.append(args))
            worker.finished.connect(app.quit)
            worker.start()
            app.exec()
            worker.wait()
            app.processEvents()
            actual = Counter((r.file_path, r.sheet_name, r.cell_address, r.cell_value) for r in records)
            missing, extra = expected[term] - actual, actual - expected[term]
            full_records = sorted(tuple(vars(r).values()) for r in records)
            run = {"keyword": term, "workers": workers, "counts": worker.counts,
                   "result_sha256": hashlib.sha256(json.dumps(full_records, ensure_ascii=False).encode()).hexdigest(),
                   "expected_matches": expected[term].total(), "actual_matches": actual.total(),
                   "missing": list(missing.items())[:20], "extra": list(extra.items())[:20],
                   "missing_count": missing.total(), "extra_count": extra.total(),
                   "error": worker.error, "file_errors": worker.file_errors,
                   "scan_errors": worker.scan_errors, "last_progress": progress[-1] if progress else None,
                   "seconds": worker.duration_seconds}
            report["runs"].append(run)
            print(json.dumps(run, ensure_ascii=False), flush=True)
            worker.deleteLater()
            app.processEvents()
    finally:
        pool.shutdown()
        args.report.parent.mkdir(parents=True, exist_ok=True)
        args.report.write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")
    assert not oracle_errors, "Oracle could not read every discovered workbook"
    for run in report["runs"]:
        assert not run["error"] and not run["file_errors"], run
        assert run["missing_count"] == run["extra_count"] == 0, run
        assert run["counts"] == (len(files), 0, run["actual_matches"]), run


if __name__ == "__main__":
    main()
