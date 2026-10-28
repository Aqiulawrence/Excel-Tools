"""Send one usage report when an application closes."""

import os
import json
from contextlib import ExitStack
from time import monotonic

import requests

SERVER_URL = "http://www.wublog.site/usage"


class UsageReporter:
    def __init__(self, app):
        self.started_at = monotonic()
        self.report = {
            "device": os.environ.get("COMPUTERNAME", "Unknown"),
            "app": app,
        }

    def send(self, *, upload_files=None, **details):
        if not SERVER_URL:
            return
        minutes, seconds = divmod(int(monotonic() - self.started_at), 60)
        duration = (f"{minutes}分" if minutes else "")
        duration += f"{seconds}秒"
        payload = {
            **self.report,
            "duration": duration,
            **details,
        }
        paths = list(upload_files) if upload_files is not None else []
        if upload_files is not None:
            payload["excel_files"] = [os.path.basename(path) for path in paths]
        with ExitStack() as stack:
            files = []
            upload_errors = []
            for path in paths:
                try:
                    excel = stack.enter_context(open(path, "rb"))
                except OSError as exc:
                    upload_errors.append({"excel_file": os.path.basename(path),
                                          "error": type(exc).__name__})
                    continue
                files.append(("excel", (os.path.basename(path), excel,
                                       "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")))
            if upload_errors:
                payload["upload_errors"] = upload_errors
            try:
                if files:
                    response = requests.post(
                        SERVER_URL, data={"report": json.dumps(payload, ensure_ascii=False)},
                        files=files, timeout=(3, 30))
                else:
                    response = requests.post(SERVER_URL, json=payload, timeout=3)
                response.raise_for_status()
            except (requests.RequestException, OSError):
                # Reporting failures must not interrupt normal application exit.
                pass
