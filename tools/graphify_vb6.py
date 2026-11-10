"""VB6 structural extractor for graphify.

graphify's built-in AST extractor has no Visual Basic 6 grammar, so this script
plays the role of graphify's "Part A" for this repo: it reads the four main .vbp
projects, parses every module they list (.frm/.bas/.cls/.ctl), and writes a
graphify extraction JSON (nodes + edges) that the normal graphify build /
cluster / report / export pipeline can consume.

What it extracts (all EXTRACTED, confidence 1.0, except where noted):
  project  --contains-->  module              (from the .vbp listing)
  module   --contains-->  procedure           (Sub / Function / Property)
  proc     --calls------>  proc               (name resolved inside the project)
  proc     --references->  form/module        (frmX.Show, Load frmX, modX.Y ...)
  proc     --instantiates-> ProgID            (CreateObject("Lib.Class"))
  ProgID   --implemented_by-> class module    (when Lib is one of our projects)
  proc     --uses_table-->  db table          (FROM/JOIN/INTO/UPDATE in SQL strings, INFERRED 0.85)
  module   --declares---->  native DLL        (Declare ... Lib "x")

Usage (from the repo root):
  python tools/graphify_vb6.py            # writes graphify-out/.graphify_ast.json
Then run tools/build_graph.ps1 (or see CLAUDE.md) to build the graph.
"""
from __future__ import annotations

import json
import re
import sys
from collections import defaultdict
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent

# Main project files. "HostForms_SA v2 - Copy" and the *-L-WRFABROA / *-Luis-PC
# .vbp variants are developer-machine copies and are intentionally excluded.
PROJECTS = [
    "New Lab-I System v2/HostInterface.vbp",
    "ComSet/ComSet.vbp",
    "HostFomrs_RE/HostForms_RE.vbp",
    "HostForms_SA v2/HostForms_SA.vbp",
]

PROC_RE = re.compile(
    r"^\s*(?:(Public|Private|Friend|Global)\s+)?(?:Static\s+)?"
    r"(Sub|Function|Property\s+Get|Property\s+Let|Property\s+Set)\s+([A-Za-z_]\w*)",
    re.I,
)
END_PROC_RE = re.compile(r"^\s*End\s+(Sub|Function|Property)\b", re.I)
DECLARE_RE = re.compile(r'^\s*(?:Public\s+|Private\s+)?Declare\s+(?:PtrSafe\s+)?(?:Sub|Function)\s+(\w+)\s+Lib\s+"([^"]+)"', re.I)
SQL_TABLE_RE = re.compile(r"\b(?:FROM|JOIN|INTO|UPDATE)\s+\[?([A-Za-z_][\w]*)\]?", re.I)
IDENT_RE = re.compile(r"(?<![\w.])([A-Za-z_]\w*)(?:\s*\.\s*([A-Za-z_]\w*))?")
LOAD_RE = re.compile(r"\b(?:Load|Unload)\s+([A-Za-z_]\w*)", re.I)

SQL_STOP = {
    "select", "where", "set", "values", "dual", "the", "a", "an", "to", "and", "or",
    "table", "on", "as", "by", "order", "group", "inner", "left", "right", "outer",
    "me", "file", "date", "now", "string", "list", "form", "line", "port", "machine", "successful",
}


def slug(s: str) -> str:
    return re.sub(r"[^a-z0-9]+", "_", s.lower()).strip("_")


def stem_id(rel: Path) -> str:
    parts = list(rel.with_suffix("").parts)
    return "_".join(slug(p) for p in parts)


def read_text(p: Path) -> str:
    return p.read_text(encoding="cp1252", errors="replace")


def strip_comment(line: str) -> tuple[str, list[str]]:
    """Return (code without comment and with string bodies blanked, list of string literals)."""
    out, strings, buf = [], [], []
    in_str = False
    i = 0
    while i < len(line):
        c = line[i]
        if in_str:
            if c == '"':
                if i + 1 < len(line) and line[i + 1] == '"':
                    buf.append('"')
                    i += 2
                    continue
                in_str = False
                strings.append("".join(buf))
                buf = []
                out.append('""')
            else:
                buf.append(c)
        else:
            if c == '"':
                in_str = True
            elif c == "'":
                break
            else:
                out.append(c)
        i += 1
    code = "".join(out)
    if re.match(r"^\s*Rem\b", code, re.I):
        return "", []
    return code, strings


def logical_lines(text: str):
    """Yield (line_no, raw_line) joining ' _' continuations."""
    acc, start = "", None
    for n, raw in enumerate(text.splitlines(), 1):
        if start is None:
            start = n
        if raw.rstrip().endswith(" _"):
            acc += raw.rstrip()[:-1] + " "
            continue
        acc += raw
        yield start, acc
        acc, start = "", None
    if acc:
        yield start, acc


def parse_vbp(vbp: Path):
    name, exe, ptype, members = None, None, None, []
    for line in read_text(vbp).splitlines():
        if "=" not in line:
            continue
        k, v = line.split("=", 1)
        v = v.strip()
        if k == "Name":
            name = v.strip('"')
        elif k == "ExeName32":
            exe = v.strip('"')
        elif k == "Type":
            ptype = v
        elif k in ("Form", "UserControl"):
            members.append((k, None, v))
        elif k in ("Module", "Class"):
            mod, _, f = v.partition(";")
            members.append((k, mod.strip(), f.strip()))
    return name, exe, ptype, members


def main() -> int:
    nodes: dict[str, dict] = {}
    edges: list[dict] = []
    seen_edges: set[tuple] = set()

    def add_node(nid, label, source_file, kind="code", loc=None, **extra):
        if nid not in nodes:
            nodes[nid] = {
                "id": nid, "label": label, "file_type": kind,
                "source_file": source_file, "source_location": loc, **extra,
            }
        return nid

    def add_edge(src, dst, rel, source_file, loc=None, conf="EXTRACTED", score=1.0):
        key = (src, dst, rel)
        if src == dst or key in seen_edges:
            return
        seen_edges.add(key)
        edges.append({
            "source": src, "target": dst, "relation": rel, "confidence": conf,
            "confidence_score": score, "source_file": source_file,
            "source_location": loc, "weight": 1.0,
        })

    project_lib_classes: dict[str, dict[str, str]] = {}  # lib name -> class name -> file node id
    progid_refs: list[tuple] = []

    for vbp_rel in PROJECTS:
        vbp = ROOT / vbp_rel
        if not vbp.exists():
            print(f"skip missing {vbp_rel}", file=sys.stderr)
            continue
        pname, exe, ptype, members = parse_vbp(vbp)
        pid = add_node(f"project_{slug(pname)}", f"{pname} ({exe}, {ptype})", vbp_rel)

        # pass 1: modules + procedures
        modules = []  # (mod_name, file_node_id, rel_path, lines, procs)
        for kind, mod_name, f in members:
            fpath = (vbp.parent / f).resolve()
            try:
                rel = fpath.relative_to(ROOT)
            except ValueError:
                rel = None
            if not fpath.exists() or rel is None:
                # referenced from outside the repo (e.g. ..\..\LIS Main Application\vbGraph)
                ext_id = f"external_{slug(Path(f).stem)}"
                add_node(ext_id, f"{Path(f).name} (outside repo)", vbp_rel, kind="concept")
                add_edge(pid, ext_id, "contains", vbp_rel)
                continue
            text = read_text(fpath)
            m = re.search(r'Attribute VB_Name = "([^"]+)"', text)
            mod_name = m.group(1) if m else (mod_name or fpath.stem)
            fid = stem_id(rel)
            add_node(fid, f"{mod_name} ({fpath.suffix.lstrip('.').lower()})", rel.as_posix())
            add_edge(pid, fid, "contains", vbp_rel)
            if kind == "Class":
                project_lib_classes.setdefault(pname.lower(), {})[mod_name.lower()] = fid
            # skip the form designer block
            body_start = max(text.find("Attribute VB_Name"), 0)
            first_code_line = text[:body_start].count("\n") + 1
            procs = []
            cur = None
            for ln, raw in logical_lines(text):
                if ln < first_code_line:
                    continue
                code, strings = strip_comment(raw)
                pm = PROC_RE.match(code)
                if pm and cur is None:
                    vis = (pm.group(1) or "Public").lower()
                    if fpath.suffix.lower() in (".frm", ".ctl") and not pm.group(1):
                        vis = "public"
                    pname_ = pm.group(3)
                    cur = {"name": pname_, "vis": vis, "start": ln, "body": []}
                    continue
                if cur is not None:
                    if END_PROC_RE.match(code):
                        procs.append(cur)
                        cur = None
                    else:
                        cur["body"].append((ln, code, strings))
                else:
                    dm = DECLARE_RE.match(code)
                    if dm:
                        dll = dm.group(2).lower()
                        if not dll.endswith((".dll", ".drv", ".ocx")):
                            dll += ".dll"
                        did = add_node(f"dll_{slug(dll)}", dll, "", kind="concept")
                        add_edge(fid, did, "declares", rel.as_posix(), f"L{ln}")
            for p in procs:
                p["id"] = f"{fid}_{slug(p['name'])}"
                add_node(p["id"], f"{mod_name}.{p['name']}()", rel.as_posix(), loc=f"L{p['start']}")
                add_edge(fid, p["id"], "contains", rel.as_posix(), f"L{p['start']}")
            modules.append((mod_name, fid, rel, procs))

        # name indexes for this project
        mod_by_name = {m[0].lower(): m for m in modules}
        public_procs = defaultdict(list)
        for mod_name, fid, rel, procs in modules:
            for p in procs:
                if p["vis"] != "private":
                    public_procs[p["name"].lower()].append(p["id"])

        # pass 2: references inside procedure bodies
        for mod_name, fid, rel, procs in modules:
            local = {p["name"].lower(): p["id"] for p in procs}
            src = rel.as_posix()
            for p in procs:
                for ln, code, strings in p["body"]:
                    loc = f"L{ln}"
                    for lm in LOAD_RE.finditer(code):
                        t = mod_by_name.get(lm.group(1).lower())
                        if t:
                            add_edge(p["id"], t[1], "references", src, loc)
                    for im in IDENT_RE.finditer(code):
                        a, b = im.group(1).lower(), (im.group(2) or "").lower()
                        if a in mod_by_name and a != mod_name.lower():
                            tmod = mod_by_name[a]
                            tp = {q["name"].lower(): q["id"] for q in tmod[3]}
                            if b and b in tp:
                                add_edge(p["id"], tp[b], "calls", src, loc)
                            else:
                                add_edge(p["id"], tmod[1], "references", src, loc)
                            continue
                        if a in local and local[a] != p["id"]:
                            add_edge(p["id"], local[a], "calls", src, loc)
                        elif a in public_procs and len(public_procs[a]) == 1 and a not in local:
                            add_edge(p["id"], public_procs[a][0], "calls", src, loc)
                    if re.search(r"CreateObject\s*\(", code, re.I):
                        for s in strings:
                            if re.fullmatch(r"[A-Za-z_]\w*\.[A-Za-z_]\w*", s):
                                gid = add_node(f"progid_{slug(s)}", f'ProgID "{s}"', "", kind="concept")
                                add_edge(p["id"], gid, "instantiates", src, loc)
                                progid_refs.append((gid, s))
                    for s in strings:
                        for tm in SQL_TABLE_RE.finditer(s):
                            t = tm.group(1)
                            if t.lower() in SQL_STOP or len(t) < 3:
                                continue
                            tid = add_node(f"table_{slug(t)}", f"table {t}", "", kind="concept")
                            add_edge(p["id"], tid, "uses_table", src, loc, "INFERRED", 0.85)

    # link ProgIDs to our own class modules (cross-project edges)
    for gid, s in progid_refs:
        lib, cls = (x.lower() for x in s.split(".", 1))
        fid = project_lib_classes.get(lib, {}).get(cls)
        if fid:
            add_edge(gid, fid, "implemented_by", nodes[gid]["source_file"])

    out = ROOT / "graphify-out" / ".graphify_ast.json"
    out.parent.mkdir(exist_ok=True)
    out.write_text(json.dumps({"nodes": list(nodes.values()), "edges": edges,
                               "input_tokens": 0, "output_tokens": 0}, indent=1), encoding="utf-8")
    print(f"VB6 AST: {len(nodes)} nodes, {len(edges)} edges -> {out.relative_to(ROOT)}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
