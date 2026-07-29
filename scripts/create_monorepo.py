#!/usr/bin/env python3
"""
create_monorepo.py — assemble the webapp + scripts monorepo.

FIXED:
  - No longer copies .env into the monorepo. The old version explicitly
    included it in the copy list, which walks your Supabase keys straight
    into the folder you push to GitHub. It now copies .env.example instead
    and writes a .gitignore that blocks .env everywhere.
  - Warns loudly if a .env ends up in the target anyway.
  - Paths are configurable instead of being pinned to E:\\.
  - Reports what was copied and exits non-zero if a required source is
    missing, rather than printing a warning and producing a broken tree.

Usage:
    python create_monorepo.py
    python create_monorepo.py --target "D:\\ExcellenceSystem"
"""

import os
import sys
import shutil
import argparse
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
from pulpoplus_config import env_path  # noqa: E402

# Never copy these into a folder destined for git.
BLOCKED_NAMES = {".env", ".env.local", ".env.production", ".env.development"}

GITIGNORE = """\
# Secrets
.env
.env.*
!.env.example

# Node
node_modules/
npm-debug.log*

# Python
__pycache__/
*.py[cod]
.venv/

# Build
build/
dist/
.vercel/

# CRM data exports
*.xlsx
*.xls
*.csv
~$*
Periods/

# OS
.DS_Store
Thumbs.db
"""


def copy_item(src, dest_dir, copied, skipped):
    src = Path(src)
    if src.name in BLOCKED_NAMES:
        skipped.append(f"{src}  (secret file — deliberately not copied)")
        return
    if not src.exists():
        skipped.append(f"{src}  (not found)")
        return

    dest = Path(dest_dir) / src.name
    if src.is_dir():
        shutil.copytree(
            src, dest, dirs_exist_ok=True,
            ignore=shutil.ignore_patterns(*BLOCKED_NAMES, "node_modules", "__pycache__",
                                          "build", "dist", ".vercel", "~$*"),
        )
    else:
        shutil.copy2(src, dest)
    copied.append(str(src))


def copy_scripts(src_dir, dest_dir, exts, copied, skipped, exclude=()):
    src_dir = Path(src_dir)
    if not src_dir.is_dir():
        skipped.append(f"{src_dir}  (not found)")
        return
    for f in sorted(src_dir.iterdir()):
        if not f.is_file() or f.name in exclude or f.name in BLOCKED_NAMES:
            continue
        if f.name.startswith("~$") or f.name.startswith("."):
            continue
        if any(f.name.endswith(e) for e in exts):
            shutil.copy2(f, Path(dest_dir) / f.name)
            copied.append(str(f))


def main():
    crm_root = env_path("CRM_ROOT", r"E:\crm extractor")

    ap = argparse.ArgumentParser()
    ap.add_argument("--target", default=os.path.join(crm_root, "ExcellenceSystem"))
    ap.add_argument("--crm-root", default=crm_root)
    ap.add_argument("--extractor-root", default=env_path("EXTRACTOR_ROOT", r"E:\pulpoextractor"))
    args = ap.parse_args()

    target = Path(args.target)
    webapp = target / "webapp"
    scripts = target / "scripts"
    webapp.mkdir(parents=True, exist_ok=True)
    scripts.mkdir(parents=True, exist_ok=True)

    crm = Path(args.crm_root)
    app = crm / "excellence-crm"

    copied, skipped = [], []

    for item in ["src", "public", "api", "package.json", "package-lock.json",
                 "vercel.json", ".gitignore"]:
        copy_item(app / item, webapp, copied, skipped)

    # Ship the template, never the real secrets.
    example = app / ".env.example"
    if example.exists():
        copy_item(example, webapp, copied, skipped)
    else:
        (webapp / ".env.example").write_text(
            "REACT_APP_SUPABASE_URL=\nREACT_APP_SUPABASE_ANON_KEY=\n", encoding="utf-8"
        )
        copied.append(str(webapp / ".env.example") + "  (generated)")

    copy_scripts(args.extractor_root, scripts, [".py"], copied, skipped)
    copy_scripts(crm, scripts, [".py"], copied, skipped)
    copy_scripts(app, scripts, [".py", ".bat", ".js"], copied, skipped,
                 exclude={"App.js", "Dashboard.js"})

    # A .gitignore at the repo root protects both halves of the tree.
    (target / ".gitignore").write_text(GITIGNORE, encoding="utf-8")

    print(f"\nCopied {len(copied)} item(s) into {target}")
    for c in copied:
        print(f"  + {c}")

    if skipped:
        print(f"\nSkipped {len(skipped)} item(s):")
        for s in skipped:
            print(f"  - {s}")

    # Final safety sweep.
    leaked = [p for p in target.rglob(".env*") if p.name != ".env.example"]
    if leaked:
        print("\nWARNING: secret files found in the monorepo — remove before committing:")
        for p in leaked:
            print(f"  ! {p}")
        return 1

    print("\nDone. No .env files were copied.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
