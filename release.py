import ast
import argparse
import json
import os
import shutil
import subprocess
import sys
import zipfile
from pathlib import Path
from urllib.request import Request, urlopen

# 自动打包、存档、上传 COS 并发送新版本号至服务器

ROOT = Path(__file__).resolve().parent
SCRIPTS = ['ImageSearch.py', 'ExcelCompare.py', 'ExcelSearch.py', 'Updater.py']


def updater_settings():
    # 不导入 Updater，避免触发其 Windows 依赖和全局初始化。
    tree = ast.parse((ROOT / 'Updater.py').read_text(encoding='utf-8'))
    return {target.id: ast.literal_eval(node.value)
            for node in tree.body if isinstance(node, ast.Assign)
            for target in node.targets if isinstance(target, ast.Name)
            and target.id in {'VERSION', 'SERVER_URL'}}


def build_and_merge(scripts):
    if os.path.exists('dist'):
        shutil.rmtree('dist')

    for script in scripts:
        subprocess.run([sys.executable, '-m', 'PyInstaller', '-w', '--onedir', script], check=True)

    target = 'dist'
    os.makedirs(target, exist_ok=True)
    internal_target = os.path.join(target, '_internal')
    os.makedirs(internal_target, exist_ok=True)

    for script in scripts:
        name = script.replace('.py', '')
        exe_src = f'dist/{name}/{name}.exe'
        shutil.copy(exe_src, target)
        internal_src = f'dist/{name}/_internal'
        if os.path.exists(internal_src):
            for item in os.listdir(internal_src):
                src = os.path.join(internal_src, item)
                dst = os.path.join(internal_target, item)
                # Merge directory contents: earlier programs may include only
                # part of a shared package (e.g. pywintypes but no pythoncom).
                if os.path.isdir(src):
                    shutil.copytree(src, dst, dirs_exist_ok=True)
                else:
                    shutil.copy2(src, dst)

    for spec in [s.replace('.py', '.spec') for s in scripts]:
        if os.path.exists(spec):
            os.remove(spec)

    for empty_dir in [f'dist/{s.replace(".py", "")}' for s in scripts]:
        if os.path.exists(empty_dir):
            shutil.rmtree(empty_dir)


def archive_dist(version):
    target = ROOT / 'dist'
    archive = target / 'dist.zip'
    # 只压缩四个 EXE 及依赖目录，避免把旧 dist.zip 压进新包。
    files = [target / Path(script).with_suffix('.exe') for script in SCRIPTS]
    if not all(path.is_file() for path in files) or not (target / '_internal').is_dir():
        raise RuntimeError('dist 中缺少 EXE 或 _internal 依赖目录。')
    files.extend(path for path in (target / '_internal').rglob('*') if path.is_file())
    with zipfile.ZipFile(archive, 'w', zipfile.ZIP_DEFLATED) as zf:
        for path in files:
            zf.write(path, path.relative_to(target))
    release = ROOT / 'release'
    release.mkdir(exist_ok=True)
    saved = release / f'{version}.zip'
    shutil.copy2(archive, saved)
    print(f'已生成 {archive} 和 {saved}', flush=True)
    return archive


def upload_dist(archive):
    archive = Path(archive)
    if not archive.is_file() or archive.stat().st_size == 0:
        raise RuntimeError('上传失败：dist.zip 不存在或为空。')
    if archive.stat().st_size > 5 * 1024 ** 3:
        raise RuntimeError('dist.zip 超过简单上传的 5GB 限制，压缩包已保留。')
    credentials_file = ROOT / 'cos_credentials.json'
    credentials = {}
    if credentials_file.is_file():
        try:
            credentials = json.loads(credentials_file.read_text(encoding='utf-8-sig'))
            if not isinstance(credentials, dict):
                raise ValueError('配置必须是对象')
        except (OSError, ValueError) as exc:
            raise RuntimeError('无法读取 cos_credentials.json，请检查文件及 JSON 格式。') from exc
    secret_id = os.environ.get('COSSECRETID') or credentials.get('SecretId')
    secret_key = os.environ.get('COSSECRETKEY') or credentials.get('SecretKey')
    token = os.environ.get('COSSESSIONTOKEN') or credentials.get('Token') or None
    if not secret_id or not secret_key:
        raise RuntimeError('请在 cos_credentials.json 中配置 SecretId 和 SecretKey，或设置对应环境变量。')
    try:
        from qcloud_cos import CosConfig, CosS3Client
        from qcloud_cos.cos_exception import CosClientError, CosServiceError
    except ImportError as exc:
        raise RuntimeError('请在运行 pack.py 的 Python 环境中安装 cos-python-sdk-v5。') from exc
    client = CosS3Client(CosConfig(Region='ap-guangzhou', SecretId=secret_id,
                                 SecretKey=secret_key, Token=token, Scheme='https'))
    print('正在上传 dist.zip 到 COS 桶根目录……', flush=True)
    try:
        with archive.open('rb') as stream:
            response = client.put_object(Bucket='wublog-1328523094', Key='dist.zip',
                                         Body=stream, ContentType='application/zip')
    except CosServiceError as exc:
        raise RuntimeError(f'COS 上传失败：HTTP {exc.get_status_code()}，'
                           f'错误码 {exc.get_error_code()}；压缩包已保留，版本未发布。') from None
    except (CosClientError, OSError):
        raise RuntimeError('COS 上传失败，请检查文件和网络；压缩包已保留，版本未发布。') from None
    print(f"已上传 wublog-1328523094/dist.zip，ETag: {response.get('ETag', '')}", flush=True)


def publish_version(settings):
    token_file = ROOT / 'release_token.txt'
    try:
        token = token_file.read_text(encoding='utf-8-sig').strip()
    except OSError as e:
        raise RuntimeError('压缩包已保留；请在 pack.py 同目录创建 release_token.txt 并填写令牌。') from e
    if not token:
        raise RuntimeError('release_token.txt 中的令牌不能为空。')
    url = settings['SERVER_URL'].rsplit('/', 1)[0] + '/publish'
    request = Request(url, data=json.dumps({'version': settings['VERSION']}).encode('utf-8'),
                      headers={'Authorization': f'Bearer {token}',
                               'Content-Type': 'application/json'}, method='POST')
    with urlopen(request, timeout=30) as response:
        result = json.load(response)
    if result.get('version') != settings['VERSION']:
        raise RuntimeError('服务器返回的版本不一致，请检查发布接口。')
    print(f"服务器版本已更新为 {settings['VERSION']}。")


def main():
    parser = argparse.ArgumentParser()
    mode = parser.add_mutually_exclusive_group()
    mode.add_argument('--publish-only', action='store_true', help='仅重试发送版本，不重新打包')
    mode.add_argument('--build-only', action='store_true', help='仅打包和存档，不上传或发布版本')
    args = parser.parse_args()
    if not args.publish_only and input('请先确认版本号是否修改：').strip().lower() not in ('', 'y', 'yes'):
        print('已取消打包。')
        raise SystemExit(0)
    os.chdir(ROOT)
    settings = updater_settings()
    if not args.publish_only:
        build_and_merge(SCRIPTS)
        archive = archive_dist(settings['VERSION'])
        if args.build_only:
            return
        upload_dist(archive)
    publish_version(settings)


if __name__ == "__main__":
    try:
        main()
    except RuntimeError as exc:
        print(f'发布失败：{exc}', file=sys.stderr)
        raise SystemExit(1)
