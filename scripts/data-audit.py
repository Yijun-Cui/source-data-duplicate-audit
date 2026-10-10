"""Portable numeric source-data audit. Python 3.10+, standard library only.

XLSX input is read through ZIP/XML; exact values use Decimal. A reviewed,
workbook-specific mapping is mandatory. Source worksheet bytes are retained
except highlighted style IDs. No workbook instructions or macros are executed.
"""
from __future__ import annotations

import argparse
import copy
import hashlib
import json
import math
import posixpath
import re
import statistics
import sys
from collections import Counter, defaultdict, deque
from datetime import datetime, timezone
from decimal import Decimal, InvalidOperation, localcontext
from itertools import combinations
from pathlib import Path
from zipfile import ZIP_DEFLATED, BadZipFile, ZipFile
from xml.sax.saxutils import quoteattr
import xml.etree.ElementTree as ET

VERSION = "1.0.0"
M = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
P = "http://schemas.openxmlformats.org/package/2006/relationships"
C = "http://schemas.openxmlformats.org/package/2006/content-types"
NS = {"m": M}
DEFAULT_PATH = Path(__file__).resolve().parents[1] / "references/default-settings.json"
LEVELS = {"priority": (6, "F8C9CD", "优先复核"), "metadata": (5, "FFF0A6", "记录复核"),
          "review": (4, "FFE2B8", "多点或结构复核"), "precision": (3, "E4DDF4", "高精度线索"),
          "near": (2, "D7EAFB", "近似线索"), "explained": (1, "DDEEDB", "已声明复用或派生"),
          "low": (0, None, "低信息量重合")}
TYPES = {
    "exact_ordered": "整组同序重复", "exact_unordered": "整组乱序重复", "exact_partial": "部分精确重复",
    "high_precision": "高精度相同值", "near_ordered": "同序近似重复", "near_unordered": "无序近似重合",
    "offset": "精确平移", "scale": "精确缩放", "affine": "精确线性变换", "near_affine": "近似线性变换",
    "near_offset": "近似平移", "near_scale": "近似缩放", "arithmetic": "等差序列",
    "near_arithmetic": "近等差序列", "geometric": "等比序列", "near_geometric": "近等比序列",
    "terminal": "末位数字集中", "duplicate_record": "成套记录重复", "near_record": "成套记录近似",
    "duplicate_id": "唯一记录键重复", "metadata_conflict": "稳定字段冲突", "derived_mismatch": "派生关系不一致",
}


class AuditError(ValueError):
    pass


def load_json(path):
    return json.loads(Path(path).read_text(encoding="utf-8-sig"))


def write_json(path, value):
    Path(path).write_text(json.dumps(value, ensure_ascii=False, indent=2), encoding="utf-8")


def digest(path):
    return hashlib.sha256(Path(path).read_bytes()).hexdigest()


def tag(name):
    return f"{{{M}}}{name}"


def xml_bytes(root, original=None):
    root_ns = root.tag[1:].split("}", 1)[0] if root.tag.startswith("{") else M
    ET.register_namespace("", root_ns)
    ET.register_namespace("r", R)
    encoded = ET.tostring(root, encoding="utf-8", xml_declaration=True)
    if original:
        # Preserve namespace declarations used only in attribute values such as
        # mc:Ignorable. ElementTree otherwise drops these unused declarations.
        match = re.search(rb'<[^?!][^>]*>', original)
        declarations = re.findall(rb'\b(xmlns(?::[\w]+)?)="([^"]*)"', match.group()) if match else []
        newroot = re.search(rb'<[^?!][^>]*>', encoded)
        extras = b''.join(b' '+name+b'='+quoteattr(value.decode('utf-8')).encode('utf-8') for name, value in declarations if not re.search(rb'\b'+re.escape(name)+rb'=', newroot.group()))
        encoded = encoded[:newroot.end()-1]+extras+encoded[newroot.end()-1:]
    return encoded


def col_number(text):
    n = 0
    for c in text.upper():
        if not "A" <= c <= "Z":
            raise AuditError(f"Invalid column: {text}")
        n = n * 26 + ord(c) - 64
    return n


def col_name(n):
    result = ""
    while n:
        n, r = divmod(n - 1, 26)
        result = chr(65 + r) + result
    return result


def address(coord):
    match = re.fullmatch(r"\$?([A-Za-z]+)\$?([1-9]\d*)", coord)
    if not match:
        raise AuditError(f"Invalid cell address: {coord}")
    c, r = col_number(match[1]), int(match[2])
    if c > 16384 or r > 1048576:
        raise AuditError(f"Out-of-bounds cell: {coord}")
    return r, c


def range_cells(ref, one_dimensional=False):
    ends = ref.split(":")
    if len(ends) > 2:
        raise AuditError(f"Invalid range: {ref}")
    r1, c1 = address(ends[0]); r2, c2 = address(ends[-1])
    if r2 < r1 or c2 < c1:
        raise AuditError(f"Reversed range: {ref}")
    if one_dimensional and r1 != r2 and c1 != c2:
        raise AuditError(f"A group must be one-dimensional: {ref}")
    if (r2-r1+1)*(c2-c1+1) > 1000000:
        raise AuditError(f"Range too large: {ref}")
    return [f"{col_name(c)}{r}" for r in range(r1, r2+1) for c in range(c1, c2+1)]


def selected_cells(spec):
    if bool(spec.get("cells")) == bool(spec.get("ranges")):
        raise AuditError("Specify exactly one of cells or ranges")
    if spec.get("cells"):
        cells = [f"{col_name(address(c)[1])}{address(c)[0]}" for c in spec["cells"]]
    else:
        cells = [c for ref in spec["ranges"] for c in range_cells(ref, True)]
    if len(set(cells)) != len(cells):
        raise AuditError("Repeated cells within one mapping entry")
    return cells


def decimal_places(value):
    return max(0, -value.normalize().as_tuple().exponent)


def related_path(owner, target):
    return target.lstrip("/") if target.startswith("/") else posixpath.normpath(posixpath.join(posixpath.dirname(owner), target))


class Book:
    def __init__(self, path):
        self.path = Path(path)
        if not self.path.is_file() or self.path.suffix.lower() != ".xlsx":
            raise AuditError("Input must be an existing .xlsx file; convert .xls/.xlsm explicitly first")
        self.sha256 = digest(path)
        with ZipFile(path) as z:
            self.parts = {n: z.read(n) for n in z.namelist()}
        if any(n.startswith("_xmlsignatures/") for n in self.parts):
            raise AuditError("Digitally signed workbook: review-copy modification is unsupported")
        wb = ET.fromstring(self.parts["xl/workbook.xml"])
        rels = {r.get("Id"): r.get("Target") for r in ET.fromstring(self.parts["xl/_rels/workbook.xml.rels"])}
        shared = []
        if "xl/sharedStrings.xml" in self.parts:
            shared = ["".join(t.text or "" for t in e.findall("m:t", NS)+e.findall("m:r/m:t", NS))
                      for e in ET.fromstring(self.parts["xl/sharedStrings.xml"])]
        styles = ET.fromstring(self.parts["xl/styles.xml"])
        formats = {int(e.get("numFmtId")): e.get("formatCode", "") for e in styles.findall("m:numFmts/m:numFmt", NS)}
        self.styles = styles
        self.style_xfs = list(styles.find(tag("cellXfs")))
        self.sheets = {}
        for node in wb.find(tag("sheets")):
            name = node.get("name")
            part = related_path("xl/workbook.xml", rels[node.get(f"{{{R}}}id")])
            tree = ET.fromstring(self.parts[part]); cells = {}
            for e in tree.findall("m:sheetData/m:row/m:c", NS):
                coord = e.get("r"); kind = e.get("t", "n"); v = e.find(tag("v"))
                formula = e.find(tag("f")); literal = v.text if v is not None else None
                style_id = int(e.get("s", "0"))
                num_id = int(self.style_xfs[style_id].get("numFmtId", "0"))
                fmt = formats.get(num_id, "")
                date_style = num_id in set(range(14, 23)) | set(range(45, 48)) or bool(re.search(r"[ymdhis]", re.sub(r'"[^\"]*"|\[[^\]]*\]|\\.', "", fmt), re.I))
                number = None; text = literal
                if kind == "s" and literal is not None:
                    text = shared[int(literal)]
                elif kind == "inlineStr":
                    ss = e.find(tag("is"))
                    text = "".join(t.text or "" for t in ss.findall("m:t", NS)+ss.findall("m:r/m:t", NS)) if ss is not None else ""
                elif kind == "n" and literal is not None:
                    try:
                        number = Decimal(literal)
                    except InvalidOperation as exc:
                        raise AuditError(f"Invalid numeric value at {name}!{coord}") from exc
                    if not number.is_finite():
                        raise AuditError(f"Non-finite numeric value at {name}!{coord}")
                cells[coord] = dict(coord=coord, number=number, raw=literal, text=text, type=kind,
                                    formula=formula.text if formula is not None else None,
                                    has_formula=formula is not None, date=date_style or kind == "d", style=style_id)
            self.sheets[name] = dict(part=part, state=node.get("state", "visible"), cells=cells,
                merges=[e.get("ref") for e in tree.findall("m:mergeCells/m:mergeCell", NS)])

    def cell(self, sheet, coord):
        if sheet not in self.sheets:
            raise AuditError(f"Unknown sheet: {sheet}")
        return self.sheets[sheet]["cells"].get(coord, dict(coord=coord, number=None, raw=None, text=None,
            type="blank", formula=None, has_formula=False, date=False, style=0))

    def inventory(self):
        sheets = []
        for name, s in self.sheets.items():
            populated = []
            for c in s["cells"].values():
                if c["number"] is not None or c["text"] is not None or c["has_formula"]:
                    item = {k: v for k, v in c.items() if k != "number"}
                    border_id = int(self.style_xfs[c["style"]].get("borderId", "0"))
                    borders = list(self.styles.find(tag("borders")))
                    item["borders"] = {side: borders[border_id].find(tag(side)).get("style")
                        for side in ("left", "right", "top", "bottom")
                        if borders[border_id].find(tag(side)) is not None and borders[border_id].find(tag(side)).get("style")}
                    item["numeric"] = c["number"] is not None
                    item["numeric_text"] = c["number"] is None and not c["has_formula"] and bool(re.fullmatch(r"[+-]?(?:\d+(?:\.\d*)?|\.\d+)(?:[eE][+-]?\d+)?", str(c["text"] or "").strip()))
                    populated.append(item)
            sheets.append(dict(name=name, state=s["state"], merges=s["merges"], cells=populated,
                numeric_count=sum(c["number"] is not None for c in s["cells"].values()),
                formula_count=sum(c["has_formula"] for c in s["cells"].values())))
        return dict(version=VERSION, source=self.path.name, sha256=self.sha256, sheets=sheets)


def validate_mapping(mapping):
    def keys(spec, allowed, location):
        if not isinstance(spec, dict):
            raise AuditError(f"Object required: {location}")
        unknown = set(spec)-set(allowed.split())
        if unknown:
            raise AuditError(f"Unknown mapping keys at {location}: {sorted(unknown)}")

    keys(mapping, "schema_version reviewed groups excluded records reuse derived_checks", "root")
    allowed = {
        "groups": "id sheet panel label cells ranges ordering role comparison_set kind include_formulas",
        "excluded": "sheet cells ranges reason",
        "records": "sheet first_row last_row fields id_fields stable_fields value_fields unique_observation",
        "reuse": "group_a group_b reason evidence",
        "derived_checks": "id source_sheet source_range target_sheet target_range operation lag status reason relative_tolerance",
    }
    for section, fields in allowed.items():
        if not isinstance(mapping.get(section, []), list):
            raise AuditError(f"List required: {section}")
        for i, spec in enumerate(mapping.get(section, [])):
            keys(spec, fields, f"{section}[{i}]")
    for g in mapping.get("groups", []):
        if g.get("role", "primary") not in ("primary", "auxiliary") or g.get("ordering", "observed") not in ("observed", "unordered", "time"):
            raise AuditError("Invalid group role or ordering")
        if "include_formulas" in g and not isinstance(g["include_formulas"], bool):
            raise AuditError("include_formulas must be boolean")
        if not isinstance(g.get("comparison_set", "measurements"), str) or not g.get("comparison_set", "measurements"):
            raise AuditError("comparison_set must be a nonempty string")
    for rec in mapping.get("records", []):
        first, last = rec.get("first_row"), rec.get("last_row")
        if not isinstance(first, int) or not isinstance(last, int) or not 1 <= first <= last <= 1048576:
            raise AuditError("Invalid record row bounds")
        if not isinstance(rec.get("fields"), dict) or not rec["fields"]:
            raise AuditError("Record fields must be a nonempty object")
        for col in rec["fields"].values():
            if not isinstance(col, str) or not re.fullmatch(r"[A-Z]+", col):
                raise AuditError("Record columns must use uppercase column letters")
            address(f"{col}1")
        for name in ("id_fields", "stable_fields", "value_fields"):
            values = rec.get(name, [])
            if not isinstance(values, list) or any(not isinstance(v, str) for v in values) or len(set(values)) != len(values):
                raise AuditError(f"Invalid record {name}")
        if "unique_observation" in rec and not isinstance(rec["unique_observation"], bool):
            raise AuditError("unique_observation must be boolean")


def mapping_groups(book, mapping):
    validate_mapping(mapping)
    if mapping.get("schema_version") != "1.0" or mapping.get("reviewed") is not True:
        raise AuditError("Mapping must have schema_version=1.0 and reviewed=true after layout review")
    excluded = {}; owners = {}; groups = []; ids = set()
    for spec in mapping.get("excluded", []):
        if not spec.get("reason"):
            raise AuditError("Every exclusion needs a reason")
        for c in selected_cells(spec):
            book.cell(spec["sheet"], c)
            excluded[(spec["sheet"], c)] = spec["reason"]
    for rec in mapping.get("records", []):
        for field, col in rec["fields"].items():
            if field not in rec.get("value_fields", []):
                for r in range(rec["first_row"], rec["last_row"]+1):
                    excluded.setdefault((rec["sheet"], f"{col}{r}"), "记录身份或分类字段")
    for spec in mapping.get("groups", []):
        gid = spec.get("id")
        if not gid or gid in ids:
            raise AuditError("Group IDs must be unique and nonempty")
        ids.add(gid); sheet = spec["sheet"]; coords = selected_cells(spec)
        values = []; slots = []
        for pos, c in enumerate(coords):
            cell = book.cell(sheet, c); n = cell["number"]
            if cell["has_formula"] and not spec.get("include_formulas", False):
                raise AuditError(f"Formula must be excluded or explicitly included as cached: {sheet}!{c}")
            if cell["has_formula"] and n is None:
                raise AuditError(f"Formula cache is not numeric: {sheet}!{c}")
            if cell["date"]:
                raise AuditError(f"Date/time cell included as measurement: {sheet}!{c}")
            if n is None and cell["text"] not in (None, ""):
                raise AuditError(f"Non-numeric content inside data group: {sheet}!{c}")
            slots.append(n)
            if n is not None:
                key = (sheet, c)
                if key in excluded:
                    raise AuditError(f"Measurement also excluded: {sheet}!{c}")
                if spec.get("role", "primary") == "primary":
                    if key in owners:
                        raise AuditError(f"Overlapping primary groups: {sheet}!{c}")
                    owners[key] = gid
                values.append(dict(coord=c, value=n, raw=cell["raw"], pos=pos,
                                   precision=decimal_places(n), cached=cell["has_formula"]))
        if not values:
            raise AuditError(f"Empty numeric group: {gid}")
        groups.append({**spec, "coords": coords, "slots": slots, "items": values, "n": len(values),
            "comparison_set": spec.get("comparison_set", "measurements"),
            "role": spec.get("role", "primary"), "ordering": spec.get("ordering", "observed")})
    if not groups:
        raise AuditError("No measurement groups mapped")
    coverage = []
    for name, s in book.sheets.items():
        for c in s["cells"].values():
            if c["number"] is not None or c["has_formula"] or c["type"] == "e":
                key = (name, c["coord"])
                status = "group" if key in owners else "excluded" if key in excluded else "unassigned"
                coverage.append(dict(sheet=name, cell=c["coord"], status=status,
                    group=owners.get(key, ""), reason=excluded.get(key, ""), raw=c["raw"],
                    formula=c["has_formula"], numeric=c["number"] is not None))
    for spec in mapping.get("reuse", []):
        if spec.get("group_a") not in ids or spec.get("group_b") not in ids or not spec.get("reason") or not spec.get("evidence"):
            raise AuditError("Reuse declarations need two valid groups, reason and evidence")
    return groups, coverage


def relative_error(a, b):
    den = max(abs(a), abs(b))
    return abs(a-b)/den if den else Decimal(0)


def scale(values):
    med = statistics.median(values)
    mad = Decimal("1.4826") * statistics.median([abs(v-med) for v in values])
    if mad:
        return mad
    mean = sum(values)/len(values)
    return (sum((v-mean)**2 for v in values)/len(values)).sqrt()


def exact_pairs(a, b):
    lookup = defaultdict(deque)
    for j, item in enumerate(b):
        lookup[item["value"]].append(j)
    pairs = []
    for i, item in enumerate(a):
        if lookup[item["value"]]:
            pairs.append((i, lookup[item["value"]].popleft()))
    return pairs


def near_pairs(a, b, tol):
    ai = sorted(range(len(a)), key=lambda i: a[i]["value"])
    bi = sorted(range(len(b)), key=lambda i: b[i]["value"])
    back = [bytearray(len(b)+1) for _ in range(len(a)+1)]
    previous = [(0, Decimal(0))]*(len(b)+1)
    for x, i in enumerate(ai, 1):
        current = [(0, Decimal(0))]*(len(b)+1)
        for y, j in enumerate(bi, 1):
            best, move = previous[y], 1
            if current[y-1] > best:
                best, move = current[y-1], 2
            err = relative_error(a[i]["value"], b[j]["value"])
            if err <= tol:
                candidate = (previous[y-1][0]+1, previous[y-1][1]-err)
                if candidate > best:
                    best, move = candidate, 3
            current[y] = best; back[x][y] = move
        previous = current
    pairs = []; x, y = len(a), len(b)
    while x and y:
        move = back[x][y]
        if move == 3:
            pairs.append((ai[x-1], bi[y-1])); x -= 1; y -= 1
        elif move == 1:
            x -= 1
        else:
            y -= 1
    return list(reversed(pairs))


def numeric_stats(a, b, pairs):
    av = [i["value"] for i in a]; bv = [i["value"] for i in b]
    errors = [relative_error(av[i], bv[j]) for i, j in pairs]
    pooled = ((scale(av)**2+scale(bv)**2)/2).sqrt()
    mae = sum(abs(av[i]-bv[j]) for i, j in pairs)/len(pairs)
    return dict(similarity=float(1-sum(errors)/len(errors)),
                scale_error=float(mae/pooled) if pooled else (0. if not mae else None))


def settings_for(path=None):
    settings = load_json(DEFAULT_PATH)
    if path:
        override = load_json(path)
        unknown = set(override)-set(settings)
        if unknown:
            raise AuditError(f"Unknown settings: {sorted(unknown)}")
        settings.update(override)
    for key, value in settings.items():
        if key in ("terminal_positions", "near_records"):
            continue
        if not isinstance(value, (int, float)) or isinstance(value, bool) or not math.isfinite(value) or value <= 0:
            raise AuditError(f"Invalid positive setting: {key}")
        if any(word in key for word in ("coverage", "similarity", "concentration", "cv", "tolerance", "error", "ratio")) and value > 1:
            raise AuditError(f"Ratio must be in (0,1]: {key}")
    for key in ("min_group_size", "min_partial_count", "high_precision_decimals", "multi_precision_count", "terminal_min_n", "max_comparison_pairs", "max_matching_cells"):
        if not isinstance(settings[key], int):
            raise AuditError(f"Integer setting required: {key}")
    if not isinstance(settings["near_records"], bool) or not isinstance(settings["terminal_positions"], list) or not settings["terminal_positions"] or any(not isinstance(x, int) or not 1 <= x <= 12 for x in settings["terminal_positions"]):
        raise AuditError("Invalid terminal positions or near_records")
    return settings


def group_pairs(a, b, pairs):
    result = []
    for i, j in pairs:
        aa, bb = a["items"][i], b["items"][j]
        result.append(dict(sheet_a=a["sheet"], cell_a=aa["coord"], raw_a=aa["raw"],
                           sheet_b=b["sheet"], cell_b=bb["coord"], raw_b=bb["raw"],
                           relative_error=float(relative_error(aa["value"], bb["value"]))))
    return result


def detect_groups(groups, mapping, settings):
    findings = []; diagnostics = Counter(); reuse = {}
    for r in mapping.get("reuse", []):
        reuse[frozenset((r["group_a"], r["group_b"]))] = r
    by_set = defaultdict(list)
    for g in groups:
        by_set[g["comparison_set"]].append(g)
    pair_count = sum(len(gg)*(len(gg)-1)//2 for gg in by_set.values())
    if pair_count > settings["max_comparison_pairs"]:
        raise AuditError(f"Comparison-pair limit exceeded: {pair_count}; partition comparison_set or raise the recorded limit")
    eligible = (pair for gg in by_set.values() for pair in combinations(gg, 2))
    tol = Decimal(str(settings["near_relative_tolerance"]))

    def emit(kind, level, a, b=None, pairs=None, note="", **stats):
        matches = group_pairs(a, b or a, pairs or [])
        related = reuse.get(frozenset((a["id"], b["id"]))) if b else None
        if related:
            level = "explained"; note += f" 已声明复用：{related['reason']}。证据：{related['evidence']}。"
        findings.append(dict(type=kind, level=level, group_a=a["id"], group_b=b["id"] if b else "",
            label_a=a.get("label", a["id"]), label_b=b.get("label", b["id"]) if b else "",
            panel_a=a.get("panel", ""), panel_b=b.get("panel", "") if b else "",
            n_a=a["n"], n_b=b["n"] if b else None, matches=matches, note=note, **stats))

    for a, b in eligible:
        diagnostics["group_pairs"] += 1
        aa, bb = a["items"], b["items"]; av = [i["value"] for i in aa]; bv = [i["value"] for i in bb]
        ep = exact_pairs(aa, bb); n = len(ep)
        hp = [(i, j) for i, j in ep if min(aa[i]["precision"], bb[j]["precision"]) >= settings["high_precision_decimals"]]
        if hp:
            emit("high_precision", "review" if len(hp) >= settings["multi_precision_count"] else "precision", a, b, hp,
                "存储数值完全相同。单点只作线索；计算后数值可能由离散分子和共同分母产生。",
                coverage_a=len(hp)/a["n"], coverage_b=len(hp)/b["n"])
        if min(a["n"], b["n"]) < settings["min_group_size"]:
            continue
        full = n == a["n"] == b["n"]
        diversity = full and len(set(av)) >= 4 and Counter(av).most_common(1)[0][1]/len(av) < .60
        informative = sum(av[i] != av[i].to_integral_value() for i, _ in ep)
        low = (informative < 3 and not diversity) or len(set(av)) == 1 or len(set(bv)) == 1
        same_panel_slices = a["sheet"] == b["sheet"] and a.get("panel") == b.get("panel") and a.get("kind") == b.get("kind") == "time_slice"
        if same_panel_slices:
            low = True
        if full:
            same_order = a["slots"] == b["slots"]
            same_order = same_order and a["ordering"] != "unordered" and b["ordering"] != "unordered"
            emit("exact_ordered" if same_order else "exact_unordered", "low" if low else "priority", a, b, ep,
                 "保留重复次数。空白位置参与原顺序核对。" + (" 整数、常量或同一时间矩阵重合，单列复核。" if low else ""), coverage_a=1., coverage_b=1.)
        elif n >= settings["min_partial_count"] and n/min(a["n"], b["n"]) >= settings["min_partial_coverage"]:
            emit("exact_partial", "low" if low else "review", a, b, ep,
                 "两组覆盖率分别列出；整数和零值匹配的信息量较低。", coverage_a=n/a["n"], coverage_b=n/b["n"])
        pure_counts = all(v == v.to_integral_value() for v in av+bv)
        if len(set(av)) == 1 or len(set(bv)) == 1:
            continue
        if not pure_counts and a["n"]*b["n"] > settings["max_matching_cells"]:
            raise AuditError(f"Matching matrix limit exceeded for {a['id']}/{b['id']}; no rules may be reported as complete")
        npairs = near_pairs(aa, bb, tol) if not pure_counts else []
        if len(npairs) > n and len(npairs) >= settings["min_partial_count"] and len(npairs)/min(a["n"], b["n"]) >= settings["near_coverage"]:
            st = numeric_stats(aa, bb, npairs)
            if st["scale_error"] is not None and st["scale_error"] <= settings["near_scale_ratio"] and st["similarity"] >= settings["near_similarity"]:
                emit("near_unordered", "low" if same_panel_slices else "near", a, b, npairs,
                    "一对一无序匹配。分布相近也可命中，不能证明复制；轨迹乱序命中不证明时间对应。",
                    coverage_a=len(npairs)/a["n"], coverage_b=len(npairs)/b["n"], **st)
        aligned = len(a["slots"]) == len(b["slots"]) and [i["pos"] for i in aa] == [i["pos"] for i in bb]
        if not aligned or a["ordering"] == "unordered" or b["ordering"] == "unordered":
            continue
        ordered_pairs = [(i, i) for i in range(a["n"]) if relative_error(av[i], bv[i]) <= tol]
        if not pure_counts and av != bv and len(ordered_pairs)/a["n"] >= settings["near_coverage"]:
            st = numeric_stats(aa, bb, [(i, i) for i in range(a["n"])])
            if st["scale_error"] is not None and st["scale_error"] <= settings["near_scale_ratio"] and st["similarity"] >= settings["near_similarity"]:
                emit("near_ordered", "low" if same_panel_slices else "near", a, b, ordered_pairs,
                     "保留原表位置。实际动物或受试者配对仍需 ID。", coverage_a=len(ordered_pairs)/a["n"], coverage_b=len(ordered_pairs)/b["n"], **st)
        if av == bv or a["n"] < 5:
            continue
        differences = [y-x for x, y in zip(av, bv)]
        nonzero = next((i for i, x in enumerate(av) if x != av[0]), None)
        same_ratio = all(y*av[0] == bv[0]*x for x, y in zip(av, bv)) if av[0] else False
        exact_linear = nonzero is not None and all((y-bv[0])*(av[nonzero]-av[0]) == (bv[nonzero]-bv[0])*(x-av[0]) for x, y in zip(av, bv))
        pairs = [(i, i) for i in range(a["n"])]
        level = "low" if same_panel_slices or pure_counts else "review"
        if len(set(differences)) == 1:
            emit("offset", level, a, b, pairs, "固定平移。核对背景扣除和派生指标。", offset=str(differences[0]))
        elif same_ratio:
            emit("scale", level, a, b, pairs, "固定缩放。核对单位换算和归一化。", ratio=str(bv[0]/av[0]))
        elif exact_linear:
            slope = (bv[nonzero]-bv[0])/(av[nonzero]-av[0])
            emit("affine", level, a, b, pairs, "固定线性变换。核对计算关系。", slope=str(slope), intercept=str(bv[0]-slope*av[0]))
        else:
            xx, yy = list(map(float, av)), list(map(float, bv))
            xmean, ymean = statistics.mean(xx), statistics.mean(yy)
            varx = sum((x-xmean)**2 for x in xx)
            sd = statistics.pstdev(yy)
            if varx and sd:
                slope = sum((x-xmean)*(y-ymean) for x, y in zip(xx, yy))/varx
                intercept = ymean-slope*xmean
                nr = math.sqrt(statistics.mean((y-slope*x-intercept)**2 for x, y in zip(xx, yy)))/sd
                if nr <= settings["affine_normalized_error"]:
                    emit("near_affine", level, a, b, pairs, "拟合残差除以组内标准差。单调趋势和派生公式可解释。", normalized_error=nr, slope=slope, intercept=intercept)
                off = statistics.median([y-x for x, y in zip(xx, yy)])
                er = math.sqrt(statistics.mean((y-x-off)**2 for x, y in zip(xx, yy)))/sd
                if er <= settings["structure_normalized_error"]:
                    emit("near_offset", level, a, b, pairs, "平移后残差按组内标准差归一化。", normalized_error=er, offset=off)
                if all(xx):
                    ratios = [y/x for x, y in zip(xx, yy)]; ratio = statistics.median(ratios)
                    cv = statistics.pstdev(ratios)/abs(ratio) if ratio else math.inf
                    er = math.sqrt(statistics.mean((y-ratio*x)**2 for x, y in zip(xx, yy)))/sd
                    if cv <= settings["ratio_cv"] and er <= settings["structure_normalized_error"]:
                        emit("near_scale", level, a, b, pairs, "比值变异和归一化残差均达标。", ratio=ratio, ratio_cv=cv, normalized_error=er)
    for g in groups:
        items = g["items"]; v = [i["value"] for i in items]
        if g["role"] != "auxiliary":
            for d in settings["terminal_positions"]:
                valid = [i for i, item in enumerate(items) if item["precision"] >= d]
                if len(valid) < settings["terminal_min_n"]:
                    continue
                digits = [int(abs(v[i])*Decimal(10)**d) % 10 for i in valid]
                digit, count = Counter(digits).most_common(1)[0]
                diagnostics["terminal_tests"] += 1
                if count/len(valid) >= settings["terminal_concentration"]:
                    pairs = [(i, i) for i, dd in zip(valid, digits) if dd == digit]
                    emit("terminal", "precision", g, pairs=pairs, note=f"小数第 {d} 位数字 {digit} 占 {count}/{len(valid)}。未补零或取整；核对分辨率和计算规则。", position=d, concentration=count/len(valid))
        if len(v) < 5 or len(set(v)) == 1 or g["ordering"] == "unordered" or any(x is None for x in g["slots"]):
            continue
        pairs = [(i, i) for i in range(len(v))]; integer = all(x == x.to_integral_value() for x in v)
        level = "low" if integer or g.get("kind") == "time_slice" else "review"
        differences = [v[i+1]-v[i] for i in range(len(v)-1)]
        if len(set(differences)) == 1:
            emit("arithmetic", level, g, pairs=pairs, note="原顺序等差。核对轴值、设定值或排序过程。", step=str(differences[0]))
        else:
            x = list(map(float, v)); dif = list(map(float, differences)); mean = statistics.mean(dif)
            er = regression_error(x)
            cv = statistics.pstdev(dif)/abs(mean) if mean else math.inf
            if cv <= settings["structure_cv"] and er <= settings["structure_normalized_error"]:
                emit("near_arithmetic", level, g, pairs=pairs, note="差分变异和线性拟合残差均达标。", difference_cv=cv, normalized_error=er)
        if all(x > 0 for x in v):
            exact_geo = all(v[i+1]*v[0] == v[i]*v[1] for i in range(len(v)-1))
            ratios = [float(v[i+1]/v[i]) for i in range(len(v)-1)]
            er = regression_error([math.log(float(x)) for x in v])
            cv = statistics.pstdev(ratios)/abs(statistics.mean(ratios))
            if exact_geo:
                emit("geometric", level, g, pairs=pairs, note="原顺序等比。核对稀释梯度或计算关系。", ratio=str(v[1]/v[0]))
            elif cv <= settings["ratio_cv"] and er <= settings["structure_normalized_error"]:
                emit("near_geometric", level, g, pairs=pairs, note="比值变异和对数线性残差均达标。", ratio_cv=cv, normalized_error=er)
    return findings, dict(diagnostics)


def regression_error(y):
    sd = statistics.pstdev(y)
    if not sd:
        return math.inf
    xx = list(range(len(y))); xm = statistics.mean(xx); ym = statistics.mean(y)
    slope = sum((x-xm)*(v-ym) for x, v in zip(xx, y))/sum((x-xm)**2 for x in xx)
    return math.sqrt(statistics.mean((v-ym-slope*(x-xm))**2 for x, v in zip(xx, y)))/sd


def record_findings(book, mapping, settings):
    findings = []; checks = []
    for spec in mapping.get("records", []):
        sheet = spec["sheet"]; fields = spec["fields"]; ids = spec.get("id_fields", [])
        stable = spec.get("stable_fields", []); metrics = spec.get("value_fields", [])
        if len(set(fields.values())) != len(fields):
            raise AuditError("Record fields must map to different columns")
        if any(k not in fields for k in ids+stable+metrics) or (stable and not ids):
            raise AuditError("Record fields or stable-field identity are incomplete")
        rows = []
        for r in range(spec["first_row"], spec["last_row"]+1):
            row = {k: book.cell(sheet, f"{col}{r}") for k, col in fields.items()}
            if not any(c["text"] is not None or c["number"] is not None for c in row.values()):
                continue
            rows.append((r, row))
        byid = defaultdict(list); missing_ids = []
        for r, row in rows:
            key = tuple(str(row[k]["text"] or "").strip() for k in ids)
            if ids and all(key):
                byid[key].append((r, row))
            elif ids:
                missing_ids.append(r)
        def emit(kind, ra, rb, keys, note):
            matches = []
            for k in keys:
                ca, cb = ra[1][k], rb[1][k]
                matches.append(dict(sheet_a=sheet, cell_a=ca["coord"], raw_a=ca["raw"] or ca["text"],
                    sheet_b=sheet, cell_b=cb["coord"], raw_b=cb["raw"] or cb["text"], relative_error=None))
            findings.append(dict(type=kind, level="metadata" if kind in ("duplicate_id", "metadata_conflict") else "review",
                group_a=f"{sheet}:row{ra[0]}", group_b=f"{sheet}:row{rb[0]}", label_a="配对记录", label_b="配对记录",
                n_a=len(keys), n_b=len(keys), matches=matches, note=note))
        conflicts = 0; repeated = sum(len(v)-1 for v in byid.values())
        for key, duplicate_rows in byid.items():
            for ra, rb in combinations(duplicate_rows, 2):
                if spec.get("unique_observation", False):
                    emit("duplicate_id", ra, rb, ids, "映射声明该记录键应唯一。纵向观测需把访视或时间加入键。")
                conflicting = [k for k in stable if ra[1][k]["text"] not in (None, "") and rb[1][k]["text"] not in (None, "") and ra[1][k]["text"] != rb[1][k]["text"]]
                if conflicting:
                    conflicts += 1
                    emit("metadata_conflict", ra, rb, ids+conflicting, "相同记录键的稳定字段冲突："+"、".join(conflicting)+"。稳定性由映射声明。")
        valid = [(r, row) for r, row in rows if len(metrics) >= 2 and all(row[k]["number"] is not None for k in metrics)]
        if len(valid)*(len(valid)-1)//2 > settings["max_comparison_pairs"]:
            raise AuditError("Record comparison limit exceeded")
        metric_scales = {k: scale([row[k]["number"] for _, row in valid]) for k in metrics} if valid else {}
        exact_count = 0; near_count = 0
        for ra, rb in combinations(valid, 2):
            av = [ra[1][k]["number"] for k in metrics]; bv = [rb[1][k]["number"] for k in metrics]
            if av == bv:
                exact_count += 1
                emit("duplicate_record", ra, rb, metrics, "同一记录表的多个指标同时相同。核对合法复测和独立样本定义。")
            elif settings["near_records"]:
                good = sum(relative_error(x, y) <= Decimal(str(settings["near_relative_tolerance"])) for x, y in zip(av, bv))/len(metrics)
                similarity = float(1-statistics.mean(relative_error(x, y) for x, y in zip(av, bv)))
                guard = all((abs(x-y)/metric_scales[k] <= Decimal(str(settings["near_scale_ratio"]))) if metric_scales[k] else x == y for k, x, y in zip(metrics, av, bv))
                if good >= settings["near_coverage"] and similarity >= settings["near_similarity"] and guard:
                    near_count += 1
                    emit("near_record", ra, rb, metrics, "按每个指标的表内波动尺度核对；指标列由映射明确声明。")
        missing_stable = [r for r, row in rows if any(row[k]["text"] in (None, "") for k in stable)]
        checks.append(dict(sheet=sheet, rows=len(rows), repeated_keys=repeated, stable_conflicts=conflicts,
            identity_status="not_assessable" if not ids else "partial" if missing_ids else "checked",
            missing_id_rows=missing_ids, missing_stable_rows=missing_stable, stable_fields=stable,
            stable_status="not_assessable" if not ids or not stable else "partial" if missing_ids or missing_stable else "checked",
            paired_rows=len(valid), omitted_paired_rows=len(rows)-len(valid), exact_record_pairs=exact_count,
            near_record_pairs=near_count if settings["near_records"] else "disabled"))
    return findings, checks


def metadata_summary(checks):
    if not checks:
        return "未提供记录表映射，样本身份冲突无法评估。"
    identity = sum(c["identity_status"] == "checked" for c in checks)
    stable = sum(c["stable_status"] == "checked" for c in checks)
    missing = sum(len(c["missing_id_rows"]) for c in checks)
    return (f"共 {len(checks)} 个记录表映射；身份键完整核对 {identity} 个，稳定字段完整核对 {stable} 个。"
            f"缺失身份键 {missing} 行。未声明字段或缺失值无法评估，详见规则与限制。")


def derived_findings(book, mapping):
    findings = []; checks = []
    for spec in mapping.get("derived_checks", []):
        if spec.get("operation") != "lag_difference" or spec.get("status") not in ("confirmed", "hypothesis"):
            raise AuditError("Derived check requires lag_difference and confirmed/hypothesis status")
        lag = spec.get("lag", 2)
        if not isinstance(lag, int) or lag < 1:
            raise AuditError("Derived lag must be a positive integer")
        sc = range_cells(spec["source_range"], True); tc = range_cells(spec["target_range"], True)
        if len(sc) != len(tc) or not spec.get("reason"):
            raise AuditError("Derived ranges must have the same length and a documented reason")
        tol = Decimal(str(spec.get("relative_tolerance", 0)))
        if not tol.is_finite() or not 0 <= tol <= 1:
            raise AuditError("Derived relative_tolerance must be in [0,1]")
        values = [book.cell(spec["source_sheet"], c)["number"] for c in sc]
        matches = 0
        for i, target in enumerate(tc):
            dest = book.cell(spec["target_sheet"], target)
            if values[i] is None or (i >= lag and values[i-lag] is None) or dest["number"] is None:
                raise AuditError("Blank or non-numeric cell inside declared derived relationship")
            expected = values[i] - (values[i-lag] if i >= lag else Decimal(0))
            if relative_error(expected, dest["number"]) <= tol:
                matches += 1
            else:
                findings.append(dict(type="derived_mismatch", level="review", group_a=spec.get("id", "derived"), group_b="",
                    label_a="派生核对", label_b="", n_a=len(tc), n_b=None,
                    matches=[dict(sheet_a=spec["source_sheet"], cell_a=sc[i], raw_a=str(values[i]),
                        sheet_b=spec["target_sheet"], cell_b=target, raw_b=dest["raw"], relative_error=None)],
                    note=f"{spec['status']}：{spec['reason']}；预期 {expected}，现存 {dest['raw']}。",
                    expected=str(expected), target_sheet=spec["target_sheet"], target_cell=target,
                    previous_source_cell=sc[i-lag] if i >= lag else "baseline=0"))
        checks.append(dict(id=spec.get("id", "derived"), status=spec["status"], total=len(tc), matching=matches,
                           mismatching=len(tc)-matches, reason=spec["reason"]))
    return findings, checks


def report_tables(result):
    summary = [
        ["项目", "结果与解释"], ["版本", result["version"]], ["输入", result["source"]], ["SHA-256", result["sha256"]],
        ["检查范围", f"{len(result['sheet_inventory'])} 个工作表；{result['numeric_count']:,} 个数值；{len(result['groups'])} 个比较单元。比较单元不等于独立样本 n。"],
        ["覆盖", "每个数值已归属主要分组或有明确排除理由。"],
        ["优先复核记录", sum(f["level"] == "priority" for f in result["findings"])],
        ["元数据", metadata_summary(result["record_checks"])],
        ["规则命中", "同一组对可能命中多个规则。记录数不能相加为独立问题数。"],
        ["解释", "完整重复、多点匹配和近似重合需要结合指标、图注、样本 ID 与计算公式复核。"],
        ["颜色", "红：完整重复；橙：多点或结构；紫：高精度单点；蓝：近似；黄：记录；绿：声明复用。低信息量不额外着色。"],
        ["源数据", "副本保留原始值、公式、布局和附件。源数据页仅增加指定位置的填充样式。"],
        ["读取顺序", "先看审查总览，再筛选查重明细的分类。逐点配对中的坐标链接可跳回源数据。"],
    ]
    for kind, count in Counter(f["type"] for f in result["findings"] if f["level"] != "low").items():
        summary.append([TYPES[kind], count])
    findings = [["编号", "分类", "类型", "组 A", "组 B", "n A", "n B", "匹配点数", "覆盖 A", "覆盖 B", "相似度", "尺度误差", "说明"]]
    low = [findings[0][:]]; pairs = [["编号", "类型", "Sheet A", "Cell A", "存储值 A", "Sheet B", "Cell B", "存储值 B", "相对误差"]]
    for f in sorted(result["findings"], key=lambda f: -LEVELS[f["level"]][0]):
        row = [f["id"], LEVELS[f["level"]][2], TYPES[f["type"]], f["group_a"], f["group_b"], f["n_a"], f["n_b"], len(f["matches"]),
               f.get("coverage_a"), f.get("coverage_b"), f.get("similarity"), f.get("scale_error"), f["note"]]
        (low if f["level"] == "low" else findings).append(row)
        if f["level"] != "low":
            for p in f["matches"]:
                pairs.append([f["id"], TYPES[f["type"]], p["sheet_a"], p["cell_a"], p["raw_a"], p["sheet_b"], p["cell_b"], p["raw_b"], p["relative_error"]])
    coverage = [["记录类别", "工作表", "组 ID/单元格", "归属/分组", "数量/存储值", "说明"]]
    for g in result["groups"]:
        coverage.append(["组", g["sheet"], g["id"], g.get("label", ""), g["n"], f"{g['role']} / {g['comparison_set']} / "+", ".join(g["coords"])])
    for c in result["coverage"]:
        coverage.append(["数值归属", c["sheet"], c["cell"], c["status"], c["raw"], c["group"] or c["reason"]])
    rules = [["项目", "值/边界"], ["设置版本", VERSION], ["映射 SHA-256", result["mapping_sha256"]]]
    rules += [[k, json.dumps(v, ensure_ascii=False)] for k, v in result["settings"].items()]
    rules += [["近似误差", "|a-b|/max(|a|,|b|)；0 对 0 为 0。误差除以组内稳健尺度，避免低方差误判。"],
        ["高精度", "按存储数值完全相等，使用有效小数位。显示小数位和计算后小数位不代表测量精度。"],
        ["末位规则", "仅纳入有效小数位足够的值。每个位置分别统计，不补零、不强制保留两位。"],
        ["局限", "阈值是筛查设置。未计算巧合概率或多重检验 p 值。无法评估的字段单独说明。"],
        ["结构核对", "Codex 需另查图号与 sheet、指标单位、分组边界和不同子图的样本数。脚本不推断具体图号含义。"]]
    for c in result["record_checks"]:
        rules.append(["记录核对", json.dumps(c, ensure_ascii=False)])
    for c in result["derived_checks"]:
        rules.append(["派生核对", json.dumps(c, ensure_ascii=False)])
    return [("审查总览", summary, [24, 115]), ("查重明细", findings, [12, 27, 23, 28, 28, 10, 10, 12, 12, 12, 12, 12, 95]),
            ("逐点配对", pairs, [12, 23, 20, 12, 27, 20, 12, 27, 12]),
            ("分组覆盖", coverage, [19, 21, 25, 35, 27, 105]), ("规则与限制", rules, [33, 115]),
            ("低信息量记录", low, [12, 27, 23, 28, 28, 10, 10, 12, 12, 12, 12, 12, 95])]


def sheet_xml(rows, widths, links=False):
    root = ET.Element(tag("worksheet"), {"xmlns:r": R})
    ET.SubElement(root, tag("dimension"), ref=f"A1:{col_name(max(map(len, rows)))}{len(rows)}")
    views = ET.SubElement(root, tag("sheetViews")); view = ET.SubElement(views, tag("sheetView"), workbookViewId="0", showGridLines="0")
    ET.SubElement(view, tag("pane"), ySplit="1", topLeftCell="A2", activePane="bottomLeft", state="frozen")
    ET.SubElement(root, tag("sheetFormatPr"), defaultRowHeight="22")
    cols = ET.SubElement(root, tag("cols"))
    for i, width in enumerate(widths, 1):
        ET.SubElement(cols, tag("col"), min=str(i), max=str(i), width=str(width), customWidth="1")
    data = ET.SubElement(root, tag("sheetData"))
    for r, row in enumerate(rows, 1):
        height = "30" if r == 1 else "52" if len(widths) == 2 else "38"
        rr = ET.SubElement(data, tag("row"), r=str(r), ht=height, customHeight="1")
        for c, value in enumerate(row, 1):
            if value is None:
                continue
            cell = ET.SubElement(rr, tag("c"), r=f"{col_name(c)}{r}", s="1" if r == 1 else "0")
            if isinstance(value, (int, float, Decimal)) and not isinstance(value, bool):
                ET.SubElement(cell, tag("v")).text = str(value)
            else:
                cell.set("t", "inlineStr"); ss = ET.SubElement(cell, tag("is")); tt = ET.SubElement(ss, tag("t"))
                tt.text = str(value); tt.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    if len(widths) > 2:
        ET.SubElement(root, tag("autoFilter"), ref=f"A1:{col_name(len(widths))}{len(rows)}")
    if links and len(rows) > 1:
        hh = ET.SubElement(root, tag("hyperlinks"))
        for r, row in enumerate(rows[1:], 2):
            for col, sh, addr in [(4, row[2], row[3]), (7, row[5], row[6])]:
                quoted = str(sh).replace("'", "''")
                ET.SubElement(hh, tag("hyperlink"), ref=f"{col_name(col)}{r}", location=f"'{quoted}'!{addr}")
    return root


def highlighted_copy(book, result, path):
    path = Path(path)
    if path.exists() or path.resolve() == book.path.resolve():
        raise AuditError("Output exists or equals input")
    parts = dict(book.parts); styles = ET.fromstring(parts["xl/styles.xml"])
    fills = styles.find(tag("fills")); xfs = styles.find(tag("cellXfs")); fonts = styles.find(tag("fonts"))
    highlights = {}; colors = {}; clones = {}
    for f in result["findings"]:
        level = f["level"]
        if level == "low":
            continue
        for p in f["matches"]:
            targets = [(p["sheet_a"], p["cell_a"]), (p["sheet_b"], p["cell_b"])]
            if f["type"] == "derived_mismatch":
                targets = [(f["target_sheet"], f["target_cell"])]
            for key in targets:
                if key not in highlights or LEVELS[level][0] > LEVELS[highlights[key]][0]:
                    highlights[key] = level
    for key, level in highlights.items():
        name, coord = key; original = book.cell(name, coord)
        if coord not in book.sheets[name]["cells"]:
            raise AuditError(f"Highlight address absent from workbook: {name}!{coord}")
        color = LEVELS[level][1]
        if color not in colors:
            fill = ET.SubElement(fills, tag("fill")); pattern = ET.SubElement(fill, tag("patternFill"), patternType="solid")
            ET.SubElement(pattern, tag("fgColor"), rgb="FF"+color); ET.SubElement(pattern, tag("bgColor"), indexed="64")
            colors[color] = len(fills)-1
        old = original["style"]; stylekey = (old, color)
        if stylekey not in clones:
            xf = copy.deepcopy(xfs[old]); xf.set("fillId", str(colors[color])); xf.set("applyFill", "1")
            clones[stylekey] = len(xfs); xfs.append(xf)
        part = book.sheets[name]["part"]
        pattern = rb'<(?:[\w]+:)?c\b[^>]*\br="'+coord.encode()+rb'"[^>]*>'
        match = re.search(pattern, parts[part])
        if match is None:
            raise AuditError(f"Unsupported worksheet cell serialization at {name}!{coord}")
        oldtag = match.group()
        newtag = re.sub(rb'\bs="\d+"', f's="{clones[stylekey]}"'.encode(), oldtag) if re.search(rb'\bs="\d+"', oldtag) else oldtag[:-1]+f' s="{clones[stylekey]}">'.encode()
        parts[part] = parts[part][:match.start()]+newtag+parts[part][match.end():]
    # New report styles are appended, so original indexes and theme remain valid.
    font = ET.SubElement(fonts, tag("font")); ET.SubElement(font, tag("sz"), val="10"); ET.SubElement(font, tag("name"), val="Microsoft YaHei")
    normal_font = len(fonts)-1
    headfont = copy.deepcopy(font); ET.SubElement(headfont, tag("b")); ET.SubElement(headfont, tag("color"), rgb="FFFFFFFF"); fonts.append(headfont)
    headfill = ET.SubElement(fills, tag("fill")); pf = ET.SubElement(headfill, tag("patternFill"), patternType="solid"); ET.SubElement(pf, tag("fgColor"), rgb="FF374151")
    report_styles = []
    for fid, fillid, fmt in [(normal_font, 0, 0), (normal_font+1, len(fills)-1, 0), (normal_font, 0, 10)]:
        xf = ET.SubElement(xfs, tag("xf"), numFmtId=str(fmt), fontId=str(fid), fillId=str(fillid), borderId="0", xfId="0", applyFont="1", applyAlignment="1", applyFill="1", applyNumberFormat="1")
        ET.SubElement(xf, tag("alignment"), vertical="center", wrapText="1")
        report_styles.append(len(xfs)-1)
    for section in (fonts, fills, xfs):
        section.set("count", str(len(section)))
    parts["xl/styles.xml"] = xml_bytes(styles, book.parts["xl/styles.xml"])
    wb = ET.fromstring(parts["xl/workbook.xml"]); sheets = wb.find(tag("sheets"))
    rels = ET.fromstring(parts["xl/_rels/workbook.xml.rels"]); types = ET.fromstring(parts["[Content_Types].xml"])
    sheetid = max(int(s.get("sheetId")) for s in sheets); rel_ids = {e.get("Id") for e in rels}
    report_names = []
    for i, (name, rows, widths) in enumerate(report_tables(result), 1):
        unique = name; suffix = 1
        while unique in book.sheets or unique in report_names:
            unique = f"{name}-{suffix}"; suffix += 1
        report_names.append(unique)
        part = f"xl/worksheets/audit-{i}.xml"
        while part in parts:
            part = part.replace(".xml", "-new.xml")
        tree = sheet_xml(rows, widths, links=name == "逐点配对")
        for c in tree.findall("m:sheetData/m:row/m:c", NS):
            row, col = address(c.get("r"))
            percent = row > 1 and ((name in ("查重明细", "低信息量记录") and col in (9, 10, 11, 12)) or (name == "逐点配对" and col == 9))
            c.set("s", str(report_styles[2 if percent else int(c.get("s"))]))
        parts[part] = xml_bytes(tree)
        rid = f"rIdAudit{i}"
        while rid in rel_ids:
            rid += "New"
        rel_ids.add(rid)
        ET.SubElement(rels, f"{{{P}}}Relationship", Id=rid, Type=R+"/worksheet", Target=part[3:])
        sheet = ET.Element(tag("sheet"), name=unique, sheetId=str(sheetid+i)); sheet.set(f"{{{R}}}id", rid)
        # Append reports to preserve defined-name localSheetId and source indexes.
        sheets.append(sheet)
        ET.SubElement(types, f"{{{C}}}Override", PartName="/"+part, ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml")
    for view in wb.findall("m:bookViews/m:workbookView", NS):
        view.set("activeTab", str(len(book.sheets))); view.set("firstSheet", "0")
    parts["xl/workbook.xml"] = xml_bytes(wb, book.parts["xl/workbook.xml"]); parts["xl/_rels/workbook.xml.rels"] = xml_bytes(rels); parts["[Content_Types].xml"] = xml_bytes(types)
    ET.register_namespace("", M)
    # Save to a new path. Original worksheet cells are not numerically rewritten.
    path = Path(path)
    if path.exists() or path.resolve() == book.path.resolve():
        raise AuditError("Output exists or equals input")
    with ZipFile(path, "w", ZIP_DEFLATED) as z:
        for part, data in parts.items():
            z.writestr(part, data)
    saved = Book(path)
    for name, source in book.sheets.items():
        def no_cell_styles(data):
            return re.sub(rb'<(?:[\w]+:)?c\b[^>]*>', lambda m: re.sub(rb' s="\d+"', b'', m.group()), data)
        if no_cell_styles(book.parts[source["part"]]) != no_cell_styles(saved.parts[source["part"]]):
            raise AuditError(f"Source worksheet changed beyond permitted style IDs: {name}")
        for coord, cell in source["cells"].items():
            final = saved.cell(name, coord)
            if (cell["raw"], cell["text"], cell["formula"], cell["type"]) != (final["raw"], final["text"], final["formula"], final["type"]):
                raise AuditError(f"Source value/formula changed: {name}!{coord}")
            if (name, coord) in highlights:
                sid = final["style"]; fillid = int(saved.style_xfs[sid].get("fillId", "0"))
                color = list(saved.styles.find(tag("fills")))[fillid].find("m:patternFill/m:fgColor", NS).get("rgb")
                if color != "FF"+LEVELS[highlights[(name, coord)]][1]:
                    raise AuditError("Saved highlight verification failed")
    if digest(book.path) != book.sha256:
        raise AuditError("Input file changed during audit")
    return dict(source_hash_unchanged=True, original_sheet_xml_preserved_except_style=True,
                original_values_and_formulas_unchanged=True, highlighted_cells=len(highlights), report_sheets=report_names,
                output_sha256=digest(path))


def run_audit(input_path, map_path, out_dir, settings_path=None):
    outdir = Path(out_dir)
    if outdir.exists() and any(outdir.iterdir()):
        raise AuditError("Output directory must be new or empty")
    book = Book(input_path); mapping = load_json(map_path); settings = settings_for(settings_path)
    outdir.mkdir(parents=True, exist_ok=True)
    with localcontext() as ctx:
        ctx.prec = 60
        groups, coverage = mapping_groups(book, mapping)
        write_json(outdir/"coverage.json", coverage)
        unknown = [c for c in coverage if c["status"] == "unassigned"]
        if unknown:
            write_json(outdir/"incomplete.json", dict(status="incomplete", reason="Unassigned numeric/formula/error cells", cells=unknown))
            raise AuditError(f"Coverage incomplete: {len(unknown)} cells; inspect coverage.json and use a new output directory after fixing mapping")
        findings, diagnostics = detect_groups(groups, mapping, settings)
        records, record_checks = record_findings(book, mapping, settings)
        derived, derived_checks = derived_findings(book, mapping)
    findings += records+derived
    for i, finding in enumerate(findings, 1):
        finding["id"] = f"F{i:05d}"
    result = dict(version=VERSION, status="complete_for_mapped_rules", source=book.path.name, sha256=book.sha256,
        mapping_sha256=digest(map_path), created_utc=datetime.now(timezone.utc).isoformat(), settings=settings,
        groups=[{k: v for k, v in g.items() if k not in ("items", "slots")} for g in groups],
        coverage=coverage, sheet_inventory=[dict(sheet=name, state=s["state"]) for name, s in book.sheets.items()],
        numeric_count=sum(c["numeric"] for c in coverage), findings=findings, diagnostics=diagnostics,
        record_checks=record_checks, derived_checks=derived_checks)
    verification = highlighted_copy(book, result, outdir/"review.xlsx")
    write_json(outdir/"analysis.json", result)
    write_json(outdir/"mapping-used.json", mapping)
    write_json(outdir/"verification.json", verification)
    main = sorted([f for f in findings if f["level"] not in ("low", "explained")], key=lambda f: -LEVELS[f["level"]][0])
    lines = ["# 源数据查重报告", "", f"输入文件：`{book.path.name}`。规则版本：`{VERSION}`。", "",
        f"本次检查 {len(book.sheets)} 个工作表、{result['numeric_count']:,} 个数值和 {len(groups)} 个比较单元。数值覆盖检查已通过。", "",
        "规则命中可能重叠。优先结合图注、样本 ID 和计算公式复核，不能把记录数当作独立问题数。", "",
        "| 编号 | 分类 | 类型 | 比较对象 | 点数 | 两组覆盖率 |", "| --- | --- | --- | --- | ---: | --- |"]
    for f in main[:30]:
        cov = " / ".join(f"{f[k]:.1%}" for k in ("coverage_a", "coverage_b") if k in f)
        objects = f"{f['group_a']} / {f['group_b']}".replace("|", "\\|")
        lines.append(f"| {f['id']} | {LEVELS[f['level']][2]} | {TYPES[f['type']]} | {objects} | {len(f['matches'])} | {cov} |")
    lines += ["", f"完整明细见 `review.xlsx` 和 `analysis.json`。另保留 {sum(f['level']=='low' for f in findings):,} 条低信息量记录。", "",
        "元数据核验："+metadata_summary(record_checks), "",
        "结构核对仍需 Codex 检查图号、指标单位、分组边界和样本数。数值归属完整不代表实验元数据完整。", "",
        "验证：输入哈希未变。原始工作表 XML 仅改变高亮样式索引，原始数值、公式和其他工作表内容保持一致。", "",
        "近似阈值、末位筛查和归一化误差是筛查口径。缺少背景分布时，不计算所谓造假概率。"]
    (outdir/"report.md").write_text("\n".join(lines)+"\n", encoding="utf-8")
    return result


def draft_map(book):
    groups = []
    for name, s in book.sheets.items():
        populated = [c for c in s["cells"].values() if c["text"] not in (None, "")]
        if not populated:
            continue
        header_row = min(address(c["coord"])[0] for c in populated)
        if s["merges"]:
            continue
        maxrow = max(address(c)[0] for c in s["cells"])
        for c in populated:
            row, col = address(c["coord"])
            if row != header_row or c["number"] is not None:
                continue
            coords = [f"{col_name(col)}{r}" for r in range(row+1, maxrow+1)]
            numeric_rows = [address(x)[0] for x in coords if book.cell(name, x)["number"] is not None]
            if len(numeric_rows) >= 4:
                groups.append(dict(id=f"G{len(groups)+1:04d}", sheet=name, panel=name, label=c["text"],
                    ranges=[f"{col_name(col)}{row+1}:{col_name(col)}{max(numeric_rows)}"], ordering="observed"))
    return dict(schema_version="1.0", reviewed=False, groups=groups, excluded=[], records=[], reuse=[], derived_checks=[])


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    sub = parser.add_subparsers(dest="command", required=True)
    for name in ("inspect", "draft-map"):
        p = sub.add_parser(name); p.add_argument("--input", required=True, type=Path); p.add_argument("--out", required=True, type=Path)
    p = sub.add_parser("audit"); p.add_argument("--input", required=True, type=Path); p.add_argument("--map", required=True, type=Path)
    p.add_argument("--out", required=True, type=Path); p.add_argument("--settings", type=Path)
    p = sub.add_parser("demo"); p.add_argument("--out", required=True, type=Path)
    args = parser.parse_args()
    try:
        if args.command == "audit":
            result = run_audit(args.input, args.map, args.out, args.settings)
            print(json.dumps(dict(status=result["status"], findings=len(result["findings"]), output=str(args.out)), ensure_ascii=False))
        elif args.command == "demo":
            create_demo(args.out)
            print(f"Demo created: {args.out}")
        else:
            if args.out.exists():
                raise AuditError("Inventory/mapping output already exists")
            book = Book(args.input); args.out.parent.mkdir(parents=True, exist_ok=True)
            write_json(args.out, book.inventory() if args.command == "inspect" else draft_map(book))
            print(f"Written: {args.out}")
    except (AuditError, KeyError, InvalidOperation, OSError, BadZipFile, ET.ParseError, OverflowError) as exc:
        print(f"Audit failed: {exc}", file=sys.stderr)
        return 2
    return 0


def make_workbook(sheet_defs, path):
    styles = ET.Element(tag("styleSheet"))
    fonts = ET.SubElement(styles, tag("fonts"), count="2")
    for header in (False, True):
        font = ET.SubElement(fonts, tag("font")); ET.SubElement(font, tag("sz"), val="11")
        ET.SubElement(font, tag("name"), val="Microsoft YaHei")
        if header:
            ET.SubElement(font, tag("b")); ET.SubElement(font, tag("color"), rgb="FFFFFFFF")
    fills = ET.SubElement(styles, tag("fills"), count="3")
    for kind in ("none", "gray125", "solid"):
        fill = ET.SubElement(fills, tag("fill")); pattern = ET.SubElement(fill, tag("patternFill"), patternType=kind)
        if kind == "solid":
            ET.SubElement(pattern, tag("fgColor"), rgb="FF374151")
    borders = ET.SubElement(styles, tag("borders"), count="1"); ET.SubElement(borders, tag("border"))
    sx = ET.SubElement(styles, tag("cellStyleXfs"), count="1")
    ET.SubElement(sx, tag("xf"), numFmtId="0", fontId="0", fillId="0", borderId="0")
    cx = ET.SubElement(styles, tag("cellXfs"), count="2")
    for fid, fill in ((0, 0), (1, 2)):
        xf = ET.SubElement(cx, tag("xf"), numFmtId="0", fontId=str(fid), fillId=str(fill), borderId="0", xfId="0", applyAlignment="1")
        ET.SubElement(xf, tag("alignment"), vertical="center", wrapText="1")
    cs = ET.SubElement(styles, tag("cellStyles"), count="1"); ET.SubElement(cs, tag("cellStyle"), name="Normal", xfId="0", builtinId="0")
    wb = ET.Element(tag("workbook"))
    views = ET.SubElement(wb, tag("bookViews")); ET.SubElement(views, tag("workbookView"), activeTab="0")
    sheets = ET.SubElement(wb, tag("sheets"))
    rels = ET.Element(f"{{{P}}}Relationships")
    types = ET.Element(f"{{{C}}}Types")
    ET.SubElement(types, f"{{{C}}}Default", Extension="rels", ContentType="application/vnd.openxmlformats-package.relationships+xml")
    ET.SubElement(types, f"{{{C}}}Default", Extension="xml", ContentType="application/xml")
    ET.SubElement(types, f"{{{C}}}Override", PartName="/xl/workbook.xml", ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml")
    ET.SubElement(types, f"{{{C}}}Override", PartName="/xl/styles.xml", ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml")
    rootrels = ET.Element(f"{{{P}}}Relationships")
    ET.SubElement(rootrels, f"{{{P}}}Relationship", Id="rId1", Type=R+"/officeDocument", Target="xl/workbook.xml")
    parts = {"xl/styles.xml": xml_bytes(styles), "_rels/.rels": xml_bytes(rootrels)}
    for i, spec in enumerate(sheet_defs, 1):
        name = spec["name"]; part = f"xl/worksheets/sheet{i}.xml"
        rows = spec["rows"]; widths = spec.get("widths", [20]*max(map(len, rows)))
        tree = sheet_xml(rows, widths)
        if spec.get("merges"):
            merge = ET.SubElement(tree, tag("mergeCells"), count=str(len(spec["merges"])))
            for ref in spec["merges"]:
                ET.SubElement(merge, tag("mergeCell"), ref=ref)
        parts[part] = xml_bytes(tree)
        node = ET.SubElement(sheets, tag("sheet"), name=name, sheetId=str(i)); node.set(f"{{{R}}}id", f"rId{i}")
        ET.SubElement(rels, f"{{{P}}}Relationship", Id=f"rId{i}", Type=R+"/worksheet", Target=part[3:])
        ET.SubElement(types, f"{{{C}}}Override", PartName="/"+part, ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml")
    ET.SubElement(rels, f"{{{P}}}Relationship", Id="rIdStyles", Type=R+"/styles", Target="styles.xml")
    parts["xl/workbook.xml"] = xml_bytes(wb); parts["xl/_rels/workbook.xml.rels"] = xml_bytes(rels); parts["[Content_Types].xml"] = xml_bytes(types)
    with ZipFile(path, "w", ZIP_DEFLATED) as z:
        for name, content in parts.items():
            z.writestr(name, content)


def create_demo(outdir):
    outdir = Path(outdir)
    if outdir.exists() and any(outdir.iterdir()):
        raise AuditError("Demo directory must be new or empty")
    outdir.mkdir(parents=True, exist_ok=True)
    base = list(map(Decimal, ["1.123456", "2.234567", "4.345678", "7.456789", "9.567891", "12.678912", "16.789123", "21.891234", "28.912345", "35.923456"]))
    perm = [4, 0, 8, 2, 9, 5, 1, 7, 3, 6]
    noise = list(map(Decimal, ["0.005", "-0.007", "0.011", "-0.006", "0.004", "-0.009", "0.015", "-0.004", "0.002", "-0.001"]))
    datasets = {
        "base": base, "exact-copy": base[:], "shuffled-copy": [base[i] for i in perm],
        "partial-copy": base[:6]+list(map(Decimal, ["101.234567", "113.987654", "124.556677", "155.887766"])),
        "near-copy": [v*Decimal("1.003") for v in base],
        "near-shuffled": [base[i]*Decimal("1.004") for i in perm],
        "offset-copy": [v+Decimal("3.5") for v in base], "scale-copy": [v*2 for v in base],
        "affine-copy": [v*Decimal("1.7")+Decimal("2.25") for v in base],
        "near-affine": [v*Decimal("1.7")+Decimal("2.25")+n/10 for v, n in zip(base, noise)],
        "near-offset": [v+Decimal("3.5")+n for v, n in zip(base, noise)],
        "near-scale": [v*(Decimal("2")+n) for v, n in zip(base, noise)],
        "arithmetic": [Decimal("100.25")+Decimal("0.75")*i for i in range(10)],
        "near-arithmetic": [Decimal("105.25")+Decimal("0.75")*i+n for i, n in enumerate(noise)],
        "geometric": [Decimal("0.125")*Decimal("1.5")**i for i in range(10)],
        "near-geometric": [Decimal("0.125")*Decimal("1.5")**i*(1+n/100) for i, n in enumerate(noise)],
        "terminal": list(map(Decimal, ["1.17", "3.27", "7.37", "13.47", "5.57", "2.67", "11.77", "17.87", "23.97", "19.07"])),
        "low-variance-a": [Decimal("100")+n/10 for n in noise],
        "low-variance-b": [Decimal("100.5")+noise[i]/10 for i in perm],
        "counts-a": list(map(Decimal, [0, 0, 1, 0, 1, 0, 0, 2, 0, 1])),
        "counts-b": list(map(Decimal, [1, 0, 0, 0, 0, 1, 0, 2, 1, 0])),
        "precision-a": [Decimal("71.123456"), Decimal("82.234567"), Decimal("91.987654"), Decimal("104.777777"), Decimal("118.444444"), Decimal("136.555555")],
        "precision-b": [Decimal("71.123456"), Decimal("82.234567"), Decimal("151.111111"), Decimal("165.333333"), Decimal("189.666666"), Decimal("209.888888")],
        "gap-a": [Decimal("11.123456"), None, Decimal("12.234567"), Decimal("13.345678"), Decimal("14.456789"), Decimal("15.567891")],
        "gap-b": [Decimal("11.123456"), Decimal("12.234567"), None, Decimal("13.345678"), Decimal("14.456789"), Decimal("15.567891")],
    }
    names = list(datasets); rows = [names]+[[datasets[n][r] if r < len(datasets[n]) else None for n in names] for r in range(10)]
    mapping = dict(schema_version="1.0", reviewed=True, groups=[], excluded=[], records=[], reuse=[], derived_checks=[])
    for col, gid in enumerate(names, 1):
        mapping["groups"].append(dict(id=gid, sheet="Columns", panel="Synthetic examples", label=gid,
            ranges=[f"{col_name(col)}2:{col_name(col)}{len(datasets[gid])+1}"], ordering="observed"))
    records = [["sample_id", "group", "sex", "value_1", "value_2"],
        ["S01", "Control", "M", Decimal("1.125678"), Decimal("3.445566")],
        ["S02", "Control", "M", Decimal("1.125678"), Decimal("3.445566")],
        ["S03", "Control", "M", Decimal("7.123456"), Decimal("8.234567")],
        ["S03", "Drug", "F", Decimal("9.987654"), Decimal("12.876543")],
        ["S04", "Drug", "F", Decimal("1.113"), Decimal("5.771")],
        ["S05", "Drug", "F", Decimal("1.114"), Decimal("5.779")]]
    mapping["records"].append(dict(sheet="Records", first_row=2, last_row=7,
        fields={"sample_id": "A", "group": "B", "sex": "C", "value_1": "D", "value_2": "E"},
        id_fields=["sample_id"], stable_fields=["group", "sex"], value_fields=["value_1", "value_2"], unique_observation=True))
    for col in ("D", "E"):
        mapping["groups"].append(dict(id=f"record-metric-{col}", sheet="Records", label=f"记录指标 {col}", ranges=[f"{col}2:{col}7"], comparison_set="record_metrics"))
    cumulative = list(map(Decimal, [0, 2, 6, 11, 13, 20, 28, 35, 39, 44]))
    derived = [v-(cumulative[i-2] if i >= 2 else 0) for i, v in enumerate(cumulative)]
    bad = derived[:]; bad[5] += 3
    derived_rows = [["Time", "Cumulative", "Two-step", "Two-step-with-error"]]+[[i, cumulative[i], derived[i], bad[i]] for i in range(10)]
    mapping["excluded"].append(dict(sheet="Derived", ranges=["A2:A11"], reason="时间轴"))
    for col in ("B", "C", "D"):
        mapping["groups"].append(dict(id=f"derived-{col}", sheet="Derived", label=col, ranges=[f"{col}2:{col}11"], comparison_set="derived"))
    for col, status in (("C", "confirmed"), ("D", "hypothesis")):
        mapping["derived_checks"].append(dict(id=f"difference-{col}", source_sheet="Derived", source_range="B2:B11", target_sheet="Derived", target_range=f"{col}2:{col}11", operation="lag_difference", lag=2, status=status, reason="模拟数据声明的两步差分"))
    horizontal = [["Fig. 1A", "Control", None, None, None, None, "Drug", None, None, None, None],
        ["Strain X"]+base[:5]+base[:5], ["Strain Y"]+[v*2 for v in base[:5]]+[v*2 for v in base[:5]]]
    for r, strain in ((2, "X"), (3, "Y")):
        for c1, c2, condition in (("B", "F", "Control"), ("G", "K", "Drug")):
            mapping["groups"].append(dict(id=f"row-{strain}-{condition}", sheet="Horizontal", panel="Fig. 1A", label=f"Strain {strain}/{condition}", ranges=[f"{c1}{r}:{c2}{r}"], comparison_set="horizontal"))
    time_rows = [["Days", "Control", None, None, None, "Drug", None, None, None]]
    for t in range(6):
        time_rows.append([t]+[Decimal("20")+Decimal(t)/10+Decimal(i)/20 for i in range(4)]+[Decimal("24")+Decimal(t)/10+Decimal(i)/20 for i in range(4)])
        for start, end, condition in (("B", "E", "Control"), ("F", "I", "Drug")):
            mapping["groups"].append(dict(id=f"time-{condition}-{t}", sheet="Time", panel="Time matrix", label=f"{condition}/day={t}", ranges=[f"{start}{t+2}:{end}{t+2}"], kind="time_slice", comparison_set="time_slices"))
    for c in range(2, 10):
        mapping["groups"].append(dict(id=f"trajectory-{col_name(c)}", sheet="Time", panel="Time matrix", label=col_name(c), ranges=[f"{col_name(c)}2:{col_name(c)}7"], kind="trajectory", role="auxiliary", comparison_set="trajectories", ordering="time"))
    mapping["excluded"].append(dict(sheet="Time", ranges=["A2:A7"], reason="天数轴"))
    make_workbook([dict(name="Columns", rows=rows), dict(name="Records", rows=records),
        dict(name="Derived", rows=derived_rows), dict(name="Horizontal", rows=horizontal, merges=["B1:F1", "G1:K1"]),
        dict(name="Time", rows=time_rows, merges=["B1:E1", "F1:I1"])], outdir/"demo.xlsx")
    write_json(outdir/"demo-map.json", mapping)
    expected = [
        ["exact_ordered", "base", "exact-copy"], ["exact_unordered", "base", "shuffled-copy"],
        ["exact_partial", "base", "partial-copy"], ["high_precision", "precision-a", "precision-b"],
        ["near_ordered", "base", "near-copy"], ["near_unordered", "base", "near-shuffled"],
        ["offset", "base", "offset-copy"], ["scale", "base", "scale-copy"], ["affine", "base", "affine-copy"],
        ["near_affine", "base", "near-affine"],
        ["near_offset", "base", "near-offset"], ["near_scale", "base", "near-scale"],
        ["arithmetic", "arithmetic", ""], ["near_arithmetic", "near-arithmetic", ""],
        ["geometric", "geometric", ""], ["near_geometric", "near-geometric", ""],
        ["terminal", "terminal", ""], ["duplicate_record", "Records:row2", "Records:row3"],
        ["near_record", "Records:row6", "Records:row7"], ["duplicate_id", "Records:row4", "Records:row5"],
        ["metadata_conflict", "Records:row4", "Records:row5"], ["derived_mismatch", "difference-D", ""],
    ]
    write_json(outdir/"expected.json", dict(expected=expected, negative_controls=["low-variance-a/low-variance-b must not be near duplicates", "gap-a/gap-b must not be ordered duplicates", "counts-a/counts-b stays low information"]))


if __name__ == "__main__":
    raise SystemExit(main())
