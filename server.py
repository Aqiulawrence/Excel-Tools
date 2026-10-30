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
UPLOAD_DIR = Path(__file__).with_name('uploads')
DEVELOPER_DEVICE = 'DESKTOP-56FFO6L'


def write_log(filename, data):
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    with open(filename, "a", encoding="utf-8") as f:
        f.write(f"[{timestamp}] {json.dumps(data, ensure_ascii=False)}\n")


@app.route('/update', methods=['POST'])
def update():
    data = request.json
    if not isinstance(data, dict) or data.get('device') != DEVELOPER_DEVICE:
        write_log("./update.log", data)
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
    if request.mimetype == 'multipart/form-data':
        try:
            data = json.loads(request.form.get('report', ''))
        except (ValueError, TypeError):
            return jsonify({"ok": False}), 400
    else:
        data = request.get_json(silent=True)
    if not isinstance(data, dict):
        return jsonify({"ok": False}), 400
    excels = request.files.getlist('excel')
    original_names = []
    for excel in excels:
        original_name = (excel.filename or '').replace('\\', '/').rsplit('/', 1)[-1]
        if not original_name.lower().endswith('.xlsx'):
            return jsonify({"ok": False, "error": "Expected an .xlsx file"}), 400
        original_names.append(original_name)
    for excel, original_name in zip(excels, original_names):
        temporary_path = None
        try:
            UPLOAD_DIR.mkdir(parents=True, exist_ok=True)
            saved_path = UPLOAD_DIR / original_name
            with tempfile.NamedTemporaryFile(dir=UPLOAD_DIR, delete=False) as temporary:
                temporary_path = Path(temporary.name)
            excel.save(temporary_path)
            os.replace(temporary_path, saved_path)
        except OSError:
            return jsonify({"ok": False, "error": "File upload failed"}), 500
        finally:
            if temporary_path is not None:
                temporary_path.unlink(missing_ok=True)
    log_file = "./developer_usage.log" if data.get('device') == DEVELOPER_DEVICE else "./usage.log"
    write_log(log_file, data)
    return jsonify({"ok": True})


if __name__ == '__main__':
    app.run(host='0.0.0.0', port=80)
