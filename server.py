from flask import Flask, request, jsonify
import json
import os
import re
import secrets
import tempfile
from datetime import datetime
from pathlib import Path

app = Flask(__name__)

VERSION_FILE = Path(__file__).with_name('version.txt')
TOKEN_FILE = Path(__file__).with_name('release_token.txt')
UPDATE_URL = "https://origin.wublog.site/dist.zip"


def write_log(filename, data):
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    with open(filename, "a", encoding="utf-8") as f:
        f.write(f"[{timestamp}] {json.dumps(data, ensure_ascii=False)}\n")


@app.route('/update', methods=['POST'])
def update():
    write_log("./update.log", request.json)
    return jsonify({
        "version": VERSION_FILE.read_text(encoding='utf-8').strip(),
        "url": UPDATE_URL
    })


@app.route('/publish', methods=['POST'])
def publish():
    try:
        token = TOKEN_FILE.read_text(encoding='utf-8-sig').strip()
    except OSError:
        return jsonify({"error": "Release token not configured"}), 503
    if not token:
        return jsonify({"error": "Release token not configured"}), 503
    if not secrets.compare_digest(
            request.headers.get('Authorization', '').encode('utf-8'),
            f'Bearer {token}'.encode('utf-8')):
        return jsonify({"error": "Unauthorized"}), 401
    data = request.get_json(silent=True)
    version = data.get('version') if isinstance(data, dict) else None
    if not isinstance(version, str) or not re.fullmatch(r'v\d+\.\d+\.\d+', version):
        return jsonify({"error": "Invalid version"}), 400
    # 原子替换，更新检查不会读到写入一半的版本号。
    with tempfile.NamedTemporaryFile(mode='w', encoding='utf-8',
                                     dir=VERSION_FILE.parent, delete=False) as f:
        f.write(version + '\n')
    try:
        os.replace(f.name, VERSION_FILE)
    finally:
        Path(f.name).unlink(missing_ok=True)
    return jsonify({"version": version})


@app.route('/usage', methods=['POST'])
def usage():
    data = request.json
    if not isinstance(data, dict):
        return jsonify({"ok": False}), 400
    write_log("./usage.log", data)
    return jsonify({"ok": True})


if __name__ == '__main__':
    app.run(host='0.0.0.0', port=80)
