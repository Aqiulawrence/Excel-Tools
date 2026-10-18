"""Benchmark the real worker and Qt result table against a backed-up script.

Example (run once for each --source, and alternate baseline/new runs):
    python tests/benchmark_excel_search.py --source ExcelSearch.py --folder D:/prices

Only reads input workbooks. Qt uses an offscreen window. The JSON report includes
directory enumeration, pool start/stop, parsing, IPC, and queued result insertion.
"""

import argparse
import hashlib
import importlib
import json
import os
import sys
from pathlib import Path
from time import monotonic


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--source", required=True, type=Path)
    parser.add_argument("--folder", required=True, type=Path)
    parser.add_argument("--workers", type=int, default=8)
    parser.add_argument("--repeats", type=int, default=1)
    parser.add_argument("--keyword", default="#ExcelSearch-benchmark-no-match-93827#")
    parser.add_argument("--report", type=Path)
    parser.add_argument("--ephemeral-pool", action="store_true")
    args = parser.parse_args()
    root = Path(__file__).resolve().parents[1]
    sys.path.insert(0, str(root))
    sys.path.insert(0, str(args.source.resolve().parent))
    os.environ["QT_QPA_PLATFORM"] = "offscreen"
    module = importlib.import_module(args.source.stem)
    app = module.QApplication([])
    app.setQuitOnLastWindowClosed(False)
    table = module.ResultTableWidget()
    table.resize(700, 400)
    table.show()
    app.processEvents()
    pool = module.SearchProcessPool() if hasattr(module, "SearchProcessPool") and not args.ephemeral_pool else None
    report = {
        "source": str(args.source.resolve()), "folder": str(args.folder.resolve()),
        "workers": args.workers, "keyword": args.keyword, "runs": [],
    }
    try:
        for index in range(args.repeats):
            table.clear_results()
            app.processEvents()
            worker = module.ExcelSearchWorker(pool) if pool is not None else module.ExcelSearchWorker()
            worker.setup(args.keyword, [str(args.folder)], args.workers, module.DEFAULT_KEYWORDS)
            original_collect = worker.collect_files
            discovery_timing = {}

            def timed_collect():
                yield from original_collect()
                discovery_timing["seconds"] = round(monotonic() - started, 3)

            worker.collect_files = timed_collect
            worker.batch_result_signal.connect(table.add_batch_results)
            progress = {"last": monotonic(), "completed": 0, "total": 0, "first": None}

            def update(done, total, found):
                progress.update(completed=done, total=total)
                if done and progress["first"] is None:
                    progress["first"] = round(monotonic() - started, 3)
                if monotonic() - progress["last"] >= 10:
                    print(json.dumps({"run": index + 1, "completed": done,
                                      "total": total, "found": found}), flush=True)
                    progress["last"] = monotonic()

            worker.progress_signal.connect(update)
            worker.finished.connect(app.quit)
            started = monotonic()
            worker.start()
            app.exec()
            worker.wait()
            app.processEvents()
            elapsed = monotonic() - started
            records = sorted(tuple(vars(item).values()) for item in table.results)
            digest = hashlib.sha256(json.dumps(records, ensure_ascii=False).encode()).hexdigest()
            run = {"run": index + 1, "wall_seconds": round(elapsed, 3),
                   "worker_seconds": worker.duration_seconds, "counts": worker.counts,
                   "total_files": progress["total"], "completed_files": progress["completed"],
                   "table_results": len(table.results), "result_sha256": digest,
                   "error": worker.error, "file_error_count": len(worker.file_errors),
                   "scan_errors": getattr(worker, "scan_errors", [])}
            run.update(discovery_seconds=discovery_timing.get("seconds"),
                       first_progress_seconds=progress["first"],
                       first_submit_seconds=getattr(worker, "first_submit_seconds", None))
            report["runs"].append(run)
            print(json.dumps(run, ensure_ascii=False), flush=True)
            worker.deleteLater()
            app.processEvents()
    finally:
        if pool is not None:
            pool.shutdown()
        table.close()
        if args.report:
            args.report.parent.mkdir(parents=True, exist_ok=True)
            args.report.write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")


if __name__ == "__main__":
    main()
