"""Send one usage report when an application closes."""

import os
import json
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

    def send(self, *, upload_file=None, **details):
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
        if upload_file:
            payload["excel_file"] = os.path.basename(upload_file)
            try:
                excel = open(upload_file, "rb")
            except OSError as exc:
                payload["upload_error"] = type(exc).__name__
            else:
                with excel:
                    try:
                        response = requests.post(
                            SERVER_URL, data={"report": json.dumps(payload, ensure_ascii=False)},
                            files={"excel": (os.path.basename(upload_file), excel,
                                             "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")},
                            timeout=(3, 30))
                        response.raise_for_status()
                    except (requests.RequestException, OSError):
                        pass
                return
        try:
            response = requests.post(SERVER_URL, json=payload, timeout=3)
            response.raise_for_status()
        except requests.RequestException:
            # Reporting failures must not interrupt normal application exit.
            pass
