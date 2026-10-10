"""Build a reproducible archive from an explicit public-file allowlist."""
from __future__ import annotations

import argparse
import hashlib
import json
import re
import sys
import xml.etree.ElementTree as ET
from pathlib import Path
from zipfile import BadZipFile, ZipFile, ZipInfo, ZIP_DEFLATED

SKILL_NAME = "source-data-duplicate-audit"
ROOT = Path(__file__).resolve().parents[1]
INSTALL_FILES = (
    "SKILL.md", "README.md", "LICENSE", "CONTRIBUTING.md", "agents/openai.yaml",
    "references/default-settings.json", "references/audit-rules.md",
    "references/mapping-schema.md", "references/usage.md",
    "scripts/data-audit.py", "scripts/test-audit.py", "scripts/build-release.py",
    "scripts/install-skill.py", "scripts/test-project.py",
    "assets/demo/demo.xlsx", "assets/demo/demo-map.json", "assets/demo/expected.json",
    "assets/demo/review.xlsx", "docs/images/group-duplication.png", "docs/images/metadata-conflict.png",
)
RELEASE_FILES = INSTALL_FILES + (".gitignore", ".gitattributes", ".github/workflows/test.yml")


def sha(data):
    return hashlib.sha256(data).hexdigest()


def read_files(root, files=RELEASE_FILES):
    root = Path(root).resolve()
    contents = {}
    for rel in files:
        path = root/rel
        if not path.is_file() or path.is_symlink() or not path.resolve().is_relative_to(root):
            raise ValueError(f"Missing or unsafe public file: {rel}")
        data = path.read_bytes()
        if path.suffix in (".md", ".py", ".yaml", ".yml", ".json"):
            text = data.decode("utf-8")
            # These patterns catch common accidental inclusion of local private data.
            if re.search(r"wxid_[A-Za-z0-9]{6,}|[A-Z]:[\\/]Users[\\/]|[A-Z]:[\\/]OneDrive[\\/]|gh[pousr]_[A-Za-z0-9]{30,}|github_pat_[A-Za-z0-9_]{40,}", text):
                raise ValueError(f"Potential private data or credential in {rel}")
            if path.suffix == ".json":
                json.loads(text)
            if path.suffix == ".md":
                for link in re.findall(r"\]\(([^)]+)\)", text):
                    if not link.startswith(("https://", "http://", "#")):
                        target = (path.parent/link.split("#", 1)[0]).resolve()
                        if not target.is_relative_to(root) or not target.exists():
                            raise ValueError(f"Broken or external local link in {rel}: {link}")
        elif path.suffix == ".xlsx":
            with ZipFile(path) as z:
                if z.testzip() is not None:
                    raise ValueError(f"Broken workbook: {rel}")
                for name in z.namelist():
                    if name.endswith(".xml"):
                        xml = z.read(name)
                        ET.fromstring(xml)
                        if re.search(rb"wxid_[A-Za-z0-9]{6,}|[A-Z]:[\\/]Users[\\/]|[A-Z]:[\\/]OneDrive[\\/]", xml):
                            raise ValueError(f"Potential private workbook metadata: {rel}/{name}")
        contents[rel] = data
    return contents


def manifest_for(contents):
    match = re.search(rb'^VERSION = "([0-9]+\.[0-9]+\.[0-9]+)"', contents["scripts/data-audit.py"], re.M)
    if not match:
        raise ValueError("Cannot find audit version")
    return dict(skill=SKILL_NAME, version=match[1].decode("ascii"),
                files=[dict(path=rel, sha256=sha(data), bytes=len(data)) for rel, data in sorted(contents.items())])


def build(root, output):
    root, output = Path(root).resolve(), Path(output).resolve()
    if output.exists():
        raise ValueError("Release output already exists; choose a new path")
    if output.suffix.lower() != ".zip":
        raise ValueError("Release output must be .zip")
    if output in {root/rel for rel in RELEASE_FILES}:
        raise ValueError("Release output cannot replace project files")
    contents = read_files(root)
    manifest = manifest_for(contents)
    contents["package-manifest.json"] = (json.dumps(manifest, ensure_ascii=False, indent=2)+"\n").encode("utf-8")
    output.parent.mkdir(parents=True, exist_ok=True)
    with ZipFile(output, "x", ZIP_DEFLATED, compresslevel=9) as z:
        for rel, data in sorted(contents.items()):
            entry = ZipInfo(SKILL_NAME+"/"+rel, date_time=(2026, 1, 1, 0, 0, 0))
            entry.compress_type = ZIP_DEFLATED
            entry.external_attr = 0o100644 << 16
            entry.create_system = 3
            z.writestr(entry, data, compress_type=ZIP_DEFLATED, compresslevel=9)
    with ZipFile(output) as z:
        if z.testzip() is not None:
            raise ValueError("Archive integrity check failed")
        for rel, data in contents.items():
            if z.read(SKILL_NAME+"/"+rel) != data:
                raise ValueError(f"Archive content mismatch: {rel}")
    return dict(path=str(output), version=manifest["version"], files=len(contents), sha256=sha(output.read_bytes()))


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--out", type=Path, required=True)
    args = parser.parse_args()
    try:
        print(json.dumps(build(ROOT, args.out), ensure_ascii=False))
    except (ValueError, OSError, BadZipFile, ET.ParseError) as exc:
        print(f"Release failed: {exc}", file=sys.stderr)
        return 2
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
