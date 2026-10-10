"""Install the bundled skill without overwriting an existing skill."""
from __future__ import annotations

import argparse
import importlib.util
import json
import os
import sys
import tempfile
from pathlib import Path

sys.dont_write_bytecode = True
spec = importlib.util.spec_from_file_location("release", Path(__file__).with_name("build-release.py"))
release = importlib.util.module_from_spec(spec)
spec.loader.exec_module(release)


def default_skills_dir():
    home = Path(os.environ["CODEX_HOME"]).expanduser() if os.environ.get("CODEX_HOME") else Path.home()/".codex"
    return home/"skills"


def install(root, skills_dir):
    root, parent = Path(root).resolve(), Path(skills_dir).expanduser().resolve()
    target = parent/release.SKILL_NAME
    if target.exists() or target.is_symlink():
        raise ValueError("Skill already exists; no files were overwritten. Use --skills-dir for an isolated test.")
    contents = release.read_files(root, release.INSTALL_FILES)
    contents["package-manifest.json"] = (json.dumps(release.manifest_for(contents), ensure_ascii=False, indent=2)+"\n").encode("utf-8")
    parent.mkdir(parents=True, exist_ok=True)
    with tempfile.TemporaryDirectory(prefix=".source-data-audit-install-", dir=parent) as tmp:
        stage = Path(tmp)/release.SKILL_NAME
        stage.mkdir()
        for rel, data in contents.items():
            destination = stage/rel
            destination.parent.mkdir(parents=True, exist_ok=True)
            destination.write_bytes(data)
        if target.exists() or target.is_symlink():
            raise ValueError("Skill destination appeared during installation; no files were overwritten")
        stage.rename(target)
    return target


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--skills-dir", type=Path, default=default_skills_dir())
    args = parser.parse_args()
    try:
        target = install(release.ROOT, args.skills_dir)
    except (ValueError, OSError) as exc:
        print(f"Install failed: {exc}", file=sys.stderr)
        return 2
    print(f"Installed: {target}\nReopen Codex and invoke $source-data-duplicate-audit.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
