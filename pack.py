import os
import shutil
import subprocess

for d in ['dist', 'build']:
    if os.path.exists(d):
        shutil.rmtree(d)

print("打包ImageSearch...")
subprocess.run(['pyinstaller', '-w', '--onedir', 'ImageSearch.py'], check=True)

print("打包ExcelCompare...")  
subprocess.run(['pyinstaller', '-w', '--onedir', 'ExcelCompare.py'], check=True)

target = 'dist/Excel-Tools'
os.makedirs(target, exist_ok=True)

shutil.copy('dist/ImageSearch/ImageSearch.exe', target)
shutil.copy('dist/ExcelCompare/ExcelCompare.exe', target)

internal_target = os.path.join(target, '_internal')
os.makedirs(internal_target, exist_ok=True)

# 合并两个_internal文件夹
for src_dir in ['dist/ImageSearch/_internal', 'dist/ExcelCompare/_internal']:
    if os.path.exists(src_dir):
        for item in os.listdir(src_dir):
            src = os.path.join(src_dir, item)
            dst = os.path.join(internal_target, item)
            if not os.path.exists(dst):
                if os.path.isdir(src):
                    shutil.copytree(src, dst)
                else:
                    shutil.copy2(src, dst)

# 清理
for spec in ['ImageSearch.spec', 'ExcelCompare.spec']:
    if os.path.exists(spec):
        os.remove(spec)

for empty_dir in ['dist/ImageSearch', 'dist/ExcelCompare']:
    if os.path.exists(empty_dir):
        shutil.rmtree(empty_dir)

print(f"完成！文件在: {target}")
