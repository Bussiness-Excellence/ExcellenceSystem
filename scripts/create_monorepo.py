import os
import shutil

target = r"E:\crm extractor\ExcellenceSystem"
webapp = os.path.join(target, "webapp")
scripts = os.path.join(target, "scripts")

os.makedirs(webapp, exist_ok=True)
os.makedirs(scripts, exist_ok=True)

webapp_files = [
    r"E:\crm extractor\excellence-crm\src",
    r"E:\crm extractor\excellence-crm\public",
    r"E:\crm extractor\excellence-crm\api",
    r"E:\crm extractor\excellence-crm\package.json",
    r"E:\crm extractor\excellence-crm\package-lock.json",
    r"E:\crm extractor\excellence-crm\vercel.json",
    r"E:\crm extractor\excellence-crm\.env",
    r"E:\crm extractor\excellence-crm\.gitignore"
]

for item in webapp_files:
    if os.path.exists(item):
        dest = os.path.join(webapp, os.path.basename(item))
        if os.path.isdir(item):
            shutil.copytree(item, dest, dirs_exist_ok=True)
        else:
            shutil.copy2(item, dest)
    else:
        print(f"Warning: {item} not found")

def copy_scripts(src_dir, dest_dir, exts):
    if not os.path.exists(src_dir): return
    for f in os.listdir(src_dir):
        if any(f.endswith(ext) for ext in exts):
            if f not in ["App.js", "Dashboard.js"]:
                shutil.copy2(os.path.join(src_dir, f), dest_dir)

copy_scripts(r"E:\pulpoextractor", scripts, [".py"])
copy_scripts(r"E:\crm extractor", scripts, [".py"])
copy_scripts(r"E:\crm extractor\excellence-crm", scripts, [".py", ".bat", ".js"])

print("Copy completed successfully.")
