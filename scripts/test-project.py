"""Test installation, release isolation and bundled demo behavior."""
from __future__ import annotations

import importlib.util
import json
import os
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch
from zipfile import ZipFile

sys.dont_write_bytecode = True
ROOT = Path(__file__).resolve().parents[1]


def module(name, filename):
    spec = importlib.util.spec_from_file_location(name, ROOT/"scripts"/filename)
    result = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(result)
    return result


release = module("release", "build-release.py")
installer = module("installer", "install-skill.py")


class ProjectTests(unittest.TestCase):
    def test_public_files_and_document_links(self):
        contents = release.read_files(ROOT)
        self.assertEqual(set(contents), set(release.RELEASE_FILES))
        self.assertEqual(len(release.manifest_for(contents)["files"]), len(contents))

    def test_installer_environment_and_no_overwrite(self):
        with tempfile.TemporaryDirectory() as tmp:
            home = Path(tmp)/"codex"
            with patch.dict(os.environ, {"CODEX_HOME": str(home)}):
                self.assertEqual(installer.default_skills_dir(), home/"skills")
            target = installer.install(ROOT, home/"skills")
            self.assertTrue((target/"SKILL.md").exists())
            manifest = json.loads((target/"package-manifest.json").read_text(encoding="utf-8"))
            for entry in manifest["files"]:
                self.assertEqual(release.sha((target/entry["path"]).read_bytes()), entry["sha256"])
            sentinel = target/"user-notes.txt"; sentinel.write_text("keep", encoding="utf-8")
            with self.assertRaisesRegex(ValueError, "already exists"):
                installer.install(ROOT, home/"skills")
            self.assertEqual(sentinel.read_text(encoding="utf-8"), "keep")

    def test_reproducible_archive_and_extracted_install(self):
        with tempfile.TemporaryDirectory() as tmp:
            tmp = Path(tmp)
            a, b = tmp/"a.zip", tmp/"b.zip"
            release.build(ROOT, a); release.build(ROOT, b)
            self.assertEqual(a.read_bytes(), b.read_bytes())
            with self.assertRaisesRegex(ValueError, "already exists"):
                release.build(ROOT, a)
            with ZipFile(a) as z:
                self.assertTrue(all(n.startswith(release.SKILL_NAME+"/") and ".." not in Path(n).parts for n in z.namelist()))
                self.assertFalse(any("__pycache__" in n or n.startswith(release.SKILL_NAME+"/local-data/") for n in z.namelist()))
                z.extractall(tmp/"extract")
            root = tmp/"extract"/release.SKILL_NAME
            self.assertEqual((root/"scripts/data-audit.py").read_bytes(), (ROOT/"scripts/data-audit.py").read_bytes())
            install = subprocess.run([sys.executable, "-I", "-B", "-X", "utf8", str(root/"scripts/install-skill.py"), "--skills-dir", str(tmp/"skills")], capture_output=True, text=True, encoding="utf-8")
            self.assertEqual(install.returncode, 0, install.stderr)

    def test_installed_audit_matches_all_bundled_examples(self):
        with tempfile.TemporaryDirectory() as tmp:
            tmp = Path(tmp)
            installed = installer.install(ROOT, tmp/"skills")
            result = tmp/"result"
            command = [sys.executable, "-I", "-B", "-X", "utf8", str(installed/"scripts/data-audit.py"),
                       "audit", "--input", str(installed/"assets/demo/demo.xlsx"), "--map", str(installed/"assets/demo/demo-map.json"), "--out", str(result)]
            proc = subprocess.run(command, capture_output=True, text=True, encoding="utf-8")
            self.assertEqual(proc.returncode, 0, proc.stderr)
            analysis = json.loads((result/"analysis.json").read_text(encoding="utf-8"))
            expected = json.loads((installed/"assets/demo/expected.json").read_text(encoding="utf-8"))["expected"]
            actual = {(f["type"], f["group_a"], f["group_b"]) for f in analysis["findings"]}
            self.assertTrue(all(tuple(e) in actual for e in expected))
            self.assertEqual(analysis["status"], "complete_for_mapped_rules")
            verification = json.loads((result/"verification.json").read_text(encoding="utf-8"))
            self.assertTrue(verification["source_hash_unchanged"])


if __name__ == "__main__":
    unittest.main(verbosity=2)
