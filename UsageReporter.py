"""Send one usage report when an application closes."""

import os
from time import monotonic

import requests

SERVER_URL = "http://www.wublog.site/usage"


class UsageReporter:
    def __init__(self, app):
        self.started_at = monotonic()
        self.report = {
            "app": app,
            "device": os.environ.get("COMPUTERNAME", "Unknown"),
        }

    def send(self, **details):
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
        try:
            response = requests.post(SERVER_URL, json=payload, timeout=3)
            response.raise_for_status()
        except requests.RequestException:
            # Reporting failures must not interrupt normal application exit.
            pass
