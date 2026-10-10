from flask import Flask, request, jsonify
import json
from datetime import datetime

app = Flask(__name__)

LATEST_VERSION = ""
UPDATE_URL = "https://origin.wublog.site/dist.zip"


def write_log(filename, data):
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    with open(filename, "a", encoding="utf-8") as f:
        f.write(f"[{timestamp}] {json.dumps(data, ensure_ascii=False)}\n")


@app.route('/update', methods=['POST'])
def update():
    write_log("./update.log", request.json)
    return jsonify({
        "version": LATEST_VERSION,
        "url": UPDATE_URL
    })


@app.route('/usage', methods=['POST'])
def usage():
    data = request.json
    if not isinstance(data, dict):
        return jsonify({"ok": False}), 400
    write_log("./usage.log", data)
    return jsonify({"ok": True})


if __name__ == '__main__':
    app.run(host='0.0.0.0', port=80)
