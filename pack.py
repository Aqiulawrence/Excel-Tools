import os
import shutil
import subprocess

def build_and_merge(scripts, output_name="Merged-Tools"):
    for d in ['dist']:
        if os.path.exists(d):
            shutil.rmtree(d)

    for script in scripts:
        subprocess.run(['pyinstaller', '-w', '--onedir', script], check=True)

    target = f'dist/{output_name}'
    os.makedirs(target, exist_ok=True)
    internal_target = os.path.join(target, '_internal')
    os.makedirs(internal_target, exist_ok=True)

    for script in scripts:
        name = script.replace('.py', '')
        exe_src = f'dist/{name}/{name}.exe'
        if os.path.exists(exe_src):
            shutil.copy(exe_src, target)
        internal_src = f'dist/{name}/_internal'
        if os.path.exists(internal_src):
            for item in os.listdir(internal_src):
                src = os.path.join(internal_src, item)
                dst = os.path.join(internal_target, item)
                if not os.path.exists(dst):
                    if os.path.isdir(src):
                        shutil.copytree(src, dst)
                    else:
                        shutil.copy2(src, dst)

    for spec in [s.replace('.py', '.spec') for s in scripts]:
        if os.path.exists(spec):
            os.remove(spec)

    for empty_dir in [f'dist/{s.replace(".py", "")}' for s in scripts]:
        if os.path.exists(empty_dir):
            shutil.rmtree(empty_dir)

    print(f"完成！文件在: {target}")
    input()

if __name__ == "__main__":
    scripts = ['ImageSearch.py', 'ExcelCompare.py', 'ExcelSearch.py']
    build_and_merge(scripts, "Excel-Tools")
