"""Behavioral tests for the portable audit; Python standard library only."""
from __future__ import annotations

import copy
import importlib.util
import itertools
import json
import random
import sys
import tempfile
import unittest
from decimal import Decimal
from pathlib import Path
from zipfile import ZipFile, ZIP_DEFLATED

sys.dont_write_bytecode = True
spec = importlib.util.spec_from_file_location("data_audit", Path(__file__).with_name("data-audit.py"))
audit = importlib.util.module_from_spec(spec)
spec.loader.exec_module(audit)


def items(values):
    return [dict(value=Decimal(str(v)), raw=str(v), coord=f"A{i+2}", pos=i,
                 precision=audit.decimal_places(Decimal(str(v)))) for i, v in enumerate(values)]


def group(gid, values, **kwargs):
    ii = items(values)
    return dict(id=gid, sheet="S", label=gid, panel="P", items=ii, slots=[i["value"] for i in ii],
                n=len(ii), role="primary", ordering="observed", comparison_set="metric", **kwargs)


def mapping(**kwargs):
    return dict(schema_version="1.0", reviewed=True, groups=kwargs.pop("groups", []),
                excluded=[], records=[], reuse=[], derived_checks=[], **kwargs)


class MatchingTests(unittest.TestCase):
    def test_exact_multiset_does_not_reuse_one_point(self):
        aa, bb = items([1, 1, 1, 2]), items([1, 2, 3])
        pp = audit.exact_pairs(aa, bb)
        self.assertEqual(len(pp), 2)
        self.assertEqual(len({j for _, j in pp}), 2)

    def test_relative_error_zero_sign_and_scale(self):
        self.assertEqual(audit.relative_error(Decimal(0), Decimal(0)), 0)
        self.assertEqual(audit.relative_error(Decimal(-1), Decimal(1)), 2)
        aa, bb = items([-4, -2, 0, 1, 8]), items([-4.02, -2.01, 0, 1.003, 8.03])
        for factor in (Decimal("0.00001"), Decimal("1000000")):
            a2 = items([x["value"]*factor for x in aa])
            b2 = items([x["value"]*factor for x in bb])
            self.assertEqual(audit.near_pairs(aa, bb, Decimal("0.01")), audit.near_pairs(a2, b2, Decimal("0.01")))

    def test_near_matches_optimal_cardinality_on_small_inputs(self):
        rng = random.Random(14)
        for _ in range(25):
            aa = items([rng.choice([1, 1.004, 1.009, 2, 2.01, -1]) for _ in range(4)])
            bb = items([rng.choice([1, 1.004, 1.009, 2, 2.01, -1]) for _ in range(4)])
            possible = [(i, j) for i in range(4) for j in range(4)
                        if audit.relative_error(aa[i]["value"], bb[j]["value"]) <= Decimal("0.006")]
            best = 0
            for n in range(1, 5):
                if any(len({i for i, _ in p}) == n and len({j for _, j in p}) == n
                       for p in itertools.combinations(possible, n)):
                    best = n
            found = audit.near_pairs(aa, bb, Decimal("0.006"))
            self.assertEqual(len(found), best)
            self.assertEqual(len(found), len({j for _, j in found}))

    def test_threshold_override_changes_behavior(self):
        a = group("a", [1.123456, 2.234567, 4.345678, 7.456789, 12.678912])
        b = group("b", [i["value"]*Decimal("1.003") for i in a["items"]])
        settings = audit.settings_for()
        loose, _ = audit.detect_groups([a, b], mapping(), settings)
        settings["near_relative_tolerance"] = 0.001
        strict, _ = audit.detect_groups([a, b], mapping(), settings)
        self.assertIn("near_ordered", {f["type"] for f in loose})
        self.assertNotIn("near_ordered", {f["type"] for f in strict})

    def test_unordered_declaration_disables_ordered_finding(self):
        a = group("a", [1.123, 2.234, 4.345, 7.456, 9.567])
        b = copy.deepcopy(a); b["id"] = "b"; b["ordering"] = "unordered"
        findings, _ = audit.detect_groups([a, b], mapping(), audit.settings_for())
        self.assertIn("exact_unordered", {f["type"] for f in findings})
        self.assertNotIn("exact_ordered", {f["type"] for f in findings})

    def test_effective_precision_not_display_or_padding(self):
        self.assertEqual(audit.decimal_places(Decimal("1.230000")), 2)
        self.assertEqual(audit.decimal_places(Decimal("1.00000")), 0)
        self.assertEqual(audit.decimal_places(Decimal("0.0000123")), 7)

    def test_comparison_set_scope_and_pair_limit(self):
        a = group("a", [1.123, 2.234, 4.345, 7.456, 9.567])
        b = copy.deepcopy(a); b["id"] = "b"; b["comparison_set"] = "other-metric"
        findings, _ = audit.detect_groups([a, b], mapping(), audit.settings_for())
        self.assertNotIn("exact_ordered", {f["type"] for f in findings})
        b["comparison_set"] = a["comparison_set"]
        c = copy.deepcopy(a); c["id"] = "c"
        settings = audit.settings_for(); settings["max_comparison_pairs"] = 1
        with self.assertRaisesRegex(audit.AuditError, "Comparison-pair limit"):
            audit.detect_groups([a, b, c], mapping(), settings)

    def test_namespace_declarations_used_in_attribute_values_survive(self):
        original = (b'<styleSheet xmlns="'+audit.M.encode()+b'" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" '
                    b'xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac" mc:Ignorable="x14ac"><fonts/></styleSheet>')
        result = audit.xml_bytes(audit.ET.fromstring(original), original)
        self.assertIn(b'xmlns:x14ac=', result)
        self.assertIn(b'Ignorable="x14ac"', result)
        audit.ET.fromstring(result)


class WorkbookTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.path = Path(self.temp.name)
        self.source = self.path/"source.xlsx"
        audit.make_workbook([dict(name="S", rows=[["A", "B"]]+[[Decimal(str(x)), Decimal(str(x))]
                            for x in [1.123456, 2.234567, 4.345678, 7.456789, 9.567891]])], self.source)
        self.book = audit.Book(self.source)
        self.map = mapping(groups=[dict(id="a", sheet="S", ranges=["A2:A6"]),
                                   dict(id="b", sheet="S", ranges=["B2:B6"])])

    def tearDown(self):
        self.temp.cleanup()

    def run_case(self, m=None, name="result"):
        file = self.path/f"{name}-map.json"
        audit.write_json(file, m or self.map)
        return audit.run_audit(self.source, file, self.path/name)

    def test_reviewed_required(self):
        self.map["reviewed"] = False
        with self.assertRaises(audit.AuditError):
            audit.mapping_groups(self.book, self.map)

    def test_primary_overlap_rejected_auxiliary_allowed(self):
        self.map["groups"].append(dict(id="c", sheet="S", ranges=["A2:A6"]))
        with self.assertRaises(audit.AuditError):
            audit.mapping_groups(self.book, self.map)
        self.map["groups"][-1]["role"] = "auxiliary"
        groups, coverage = audit.mapping_groups(self.book, self.map)
        self.assertEqual(len(groups), 3)
        self.assertTrue(all(c["status"] == "group" for c in coverage))

    def test_coverage_gap_fails_without_review(self):
        self.map["groups"].pop()
        with self.assertRaisesRegex(audit.AuditError, "Coverage incomplete"):
            self.run_case()
        self.assertTrue((self.path/"result/incomplete.json").exists())
        self.assertFalse((self.path/"result/review.xlsx").exists())

    def test_unknown_mapping_and_settings_rejected(self):
        self.map["groups"][0]["orderng"] = "unordered"
        with self.assertRaisesRegex(audit.AuditError, "Unknown mapping"):
            audit.mapping_groups(self.book, self.map)
        file = self.path/"settings.json"; audit.write_json(file, {"near_covrage": 0.9})
        with self.assertRaisesRegex(audit.AuditError, "Unknown settings"):
            audit.settings_for(file)

    def test_formula_cache_and_local_named_range_preserved(self):
        parts = dict(self.book.parts)
        tree = audit.ET.fromstring(parts["xl/worksheets/sheet1.xml"])
        cell = tree.find("m:sheetData/m:row/m:c[@r='A2']", audit.NS)
        formula = audit.ET.Element(audit.tag("f")); formula.text = "1.123456"; cell.insert(0, formula)
        parts["xl/worksheets/sheet1.xml"] = audit.xml_bytes(tree)
        wb = audit.ET.fromstring(parts["xl/workbook.xml"])
        names = audit.ET.SubElement(wb, audit.tag("definedNames"))
        audit.ET.SubElement(names, audit.tag("definedName"), name="Local", localSheetId="0").text = "'S'!$A$2"
        parts["xl/workbook.xml"] = audit.xml_bytes(wb)
        parts["customXml/item1.xml"] = b"<metadata>keep me</metadata>"
        with ZipFile(self.source, "w", ZIP_DEFLATED) as z:
            for name, data in parts.items():
                z.writestr(name, data)
        self.book = audit.Book(self.source)
        with self.assertRaisesRegex(audit.AuditError, "Formula must be"):
            audit.mapping_groups(self.book, self.map)
        self.map["groups"][0]["include_formulas"] = True
        result = self.run_case()
        saved = audit.Book(self.path/"result/review.xlsx")
        self.assertEqual(saved.cell("S", "A2")["formula"], "1.123456")
        self.assertEqual(saved.cell("S", "A2")["raw"], self.book.cell("S", "A2")["raw"])
        self.assertEqual(saved.parts["customXml/item1.xml"], parts["customXml/item1.xml"])
        wb2 = audit.ET.fromstring(saved.parts["xl/workbook.xml"])
        self.assertEqual(list(wb2.find(audit.tag("sheets")))[0].get("name"), "S")
        self.assertEqual(wb2.find("m:definedNames/m:definedName", audit.NS).get("localSheetId"), "0")
        verification = audit.load_json(self.path/"result/verification.json")
        self.assertTrue(verification["source_hash_unchanged"])
        self.assertTrue(verification["original_values_and_formulas_unchanged"])
        self.assertEqual(verification["highlighted_cells"], 10)
        self.assertEqual(result["status"], "complete_for_mapped_rules")

    def test_package_xml_default_namespace_and_opc_parts(self):
        self.run_case()
        saved = audit.Book(self.path/"result/review.xlsx")
        types = saved.parts["[Content_Types].xml"]
        self.assertIn(b'<Types xmlns="'+audit.C.encode()+b'"', types)
        root = audit.ET.fromstring(types)
        for entry in root:
            if entry.tag.endswith("Override"):
                self.assertIn(entry.get("PartName").lstrip("/"), saved.parts)
        wb = audit.ET.fromstring(saved.parts["xl/workbook.xml"])
        self.assertEqual(wb.find("m:bookViews/m:workbookView", audit.NS).get("activeTab"), "1")

    def test_no_metadata_reported_as_not_assessable(self):
        result = self.run_case()
        self.assertEqual(result["record_checks"], [])
        self.assertIn("无法评估", (self.path/"result/report.md").read_text(encoding="utf-8"))

    def test_existing_output_and_input_target_rejected(self):
        self.run_case()
        with self.assertRaises(audit.AuditError):
            self.run_case()
        with self.assertRaises(audit.AuditError):
            audit.highlighted_copy(self.book, {}, self.source)

    def test_resource_limit_does_not_produce_success(self):
        file = self.path/"map.json"; audit.write_json(file, self.map)
        setting = self.path/"settings.json"; audit.write_json(setting, {"max_matching_cells": 2})
        with self.assertRaisesRegex(audit.AuditError, "Matching matrix limit"):
            audit.run_audit(self.source, file, self.path/"limit", setting)
        self.assertFalse((self.path/"limit/review.xlsx").exists())


class RecordTests(unittest.TestCase):
    def test_longitudinal_key_and_missing_identity(self):
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp)/"records.xlsx"
            rows = [["ID", "visit", "sex", "v1", "v2"],
                    ["S1", "Day1", "F", 1.1234, 2.2345], ["S1", "Day2", "F", 3.3456, 4.4567],
                    [None, "Day1", "M", 5.5678, 6.6789]]
            audit.make_workbook([dict(name="R", rows=rows)], path)
            rec = dict(sheet="R", first_row=2, last_row=4,
                       fields=dict(id="A", visit="B", sex="C", v1="D", v2="E"),
                       id_fields=["id", "visit"], stable_fields=["sex"],
                       value_fields=["v1", "v2"], unique_observation=True)
            mm = mapping(); mm["records"] = [rec]
            findings, checks = audit.record_findings(audit.Book(path), mm, audit.settings_for())
            self.assertNotIn("duplicate_id", {f["type"] for f in findings})
            self.assertEqual(checks[0]["identity_status"], "partial")
            self.assertEqual(checks[0]["missing_id_rows"], [4])
            rec["id_fields"] = []; rec["stable_fields"] = []
            _, checks = audit.record_findings(audit.Book(path), mm, audit.settings_for())
            self.assertEqual(checks[0]["identity_status"], "not_assessable")


class DemoTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.temp = tempfile.TemporaryDirectory()
        cls.path = Path(cls.temp.name)/"demo"
        audit.create_demo(cls.path)
        cls.result = audit.run_audit(cls.path/"demo.xlsx", cls.path/"demo-map.json", cls.path/"result")

    @classmethod
    def tearDownClass(cls):
        cls.temp.cleanup()

    def test_all_rule_types_have_expected_examples(self):
        expected = audit.load_json(self.path/"expected.json")["expected"]
        actual = {(f["type"], f["group_a"], f["group_b"]) for f in self.result["findings"]}
        self.assertEqual({e[0] for e in expected}, set(audit.TYPES))
        for e in expected:
            with self.subTest(example=e):
                self.assertIn(tuple(e), actual)

    def test_low_variance_and_gaps_are_not_ordered_copies(self):
        for f in self.result["findings"]:
            pair = {f["group_a"], f["group_b"]}
            if pair == {"low-variance-a", "low-variance-b"}:
                self.assertNotIn(f["type"], ("near_ordered", "near_unordered"))
            if pair == {"gap-a", "gap-b"}:
                self.assertNotIn(f["type"], ("exact_ordered", "near_ordered", "offset", "scale", "affine"))
            if pair == {"counts-a", "counts-b"} and f["type"] == "exact_unordered":
                self.assertEqual(f["level"], "low")

    def test_shared_control_keeps_evidence_but_is_explained(self):
        mm = audit.load_json(self.path/"demo-map.json")
        mm["reuse"] = [dict(group_a="base", group_b="exact-copy", reason="Shared control", evidence="Synthetic protocol")]
        groups, _ = audit.mapping_groups(audit.Book(self.path/"demo.xlsx"), mm)
        findings, _ = audit.detect_groups(groups, mm, audit.settings_for())
        relevant = [f for f in findings if {f["group_a"], f["group_b"]} == {"base", "exact-copy"}]
        self.assertTrue(relevant)
        self.assertTrue(all(f["level"] == "explained" for f in relevant))

    def test_derived_check_error_localization(self):
        check = [c for c in self.result["derived_checks"] if c["id"] == "difference-C"][0]
        self.assertEqual(check["mismatching"], 0)
        errors = [f for f in self.result["findings"] if f["type"] == "derived_mismatch"]
        self.assertEqual(len(errors), 1)
        self.assertEqual(errors[0]["target_cell"], "D7")


if __name__ == "__main__":
    unittest.main(verbosity=2)
