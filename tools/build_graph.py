"""Rebuild the graphify knowledge graph for this VB6 repo.

Run from the repo root with the interpreter that has graphify installed:
    & (Get-Content graphify-out\\.graphify_python) tools\\build_graph.py
or, if graphify was installed with `uv tool install graphifyy`:
    & "$(uv tool dir)\\graphifyy\\Scripts\\python.exe" tools\\build_graph.py

Steps: VB6 extraction (tools/graphify_vb6.py) -> graphify build -> cluster ->
auto-label communities -> GRAPH_REPORT.md + graph.json -> graph.html.
No LLM / API key is needed; log and data .txt files are intentionally not
fed to graphify (they are instrument traffic dumps, not documentation).
"""
from __future__ import annotations

import collections
import json
import re
import shutil
import subprocess
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(Path(__file__).resolve().parent))

import graphify_vb6  # noqa: E402
from graphify.analyze import god_nodes, suggest_questions, surprising_connections  # noqa: E402
from graphify.build import build_from_json  # noqa: E402
from graphify.cluster import cluster, score_all  # noqa: E402
from graphify.export import to_json  # noqa: E402
from graphify.report import generate  # noqa: E402

PROJECT_PREFIX = {
    "new_lab_i_system_v2": "Interface",
    "comset": "ComSet",
    "hostfomrs_re": "RE",
    "hostforms_sa_v2": "SA",
}


def friendly(mod: str) -> str:
    m = re.sub(r"^(frm|mod|cls|mdi)", "", mod)
    m = re.sub(r"(?<=[a-z])(?=[A-Z])", " ", m)
    return m.strip() or mod


def label_communities(G, comms):
    labels, used = {}, collections.Counter()
    for k, members in comms.items():
        mods = collections.Counter(
            G.nodes[n].get("label", "").split(".")[0].split(" (")[0] for n in members
        )
        top = mods.most_common(1)[0][0]
        proj = collections.Counter(
            next((v for p, v in PROJECT_PREFIX.items() if n.startswith(p)), "Shared")
            for n in members
        ).most_common(1)[0][0]
        if top.startswith("table "):
            name = "Database Tables"
        elif top.startswith("ProgID"):
            name = "COM ProgIDs"
        else:
            name = friendly(top)
        lab = f"{proj}: {name}"
        used[lab] += 1
        labels[k] = lab if used[lab] == 1 else f"{lab} ({used[lab]})"
    return labels


def main() -> int:
    out = ROOT / "graphify-out"
    out.mkdir(exist_ok=True)
    graphify_vb6.main()
    ast = json.loads((out / ".graphify_ast.json").read_text(encoding="utf-8"))
    extraction = {"nodes": ast["nodes"], "edges": ast["edges"], "hyperedges": [],
                  "input_tokens": 0, "output_tokens": 0}

    G = build_from_json(extraction, root=str(ROOT), directed=True)
    if G.number_of_nodes() == 0:
        print("ERROR: empty graph")
        return 1
    comms = cluster(G)
    coh = score_all(G, comms)
    labels = label_communities(G, comms)
    gods = god_nodes(G)
    surprises = surprising_connections(G, comms)
    questions = suggest_questions(G, comms, labels)
    code_files = sorted({G.nodes[n].get("source_file") for n in G
                         if G.nodes[n].get("file_type") == "code"})
    detection = {"total_files": len(code_files), "total_words": 0,
                 "files": {"code": code_files}, "scan_root": str(ROOT)}
    report = generate(G, comms, coh, labels, gods, surprises, detection,
                      {"input": 0, "output": 0}, str(ROOT), suggested_questions=questions)

    # force=True semantics: remove old graph.json so the shrink-guard does not
    # block a legitimate rebuild after files were removed from a .vbp.
    (out / "graph.json").unlink(missing_ok=True)
    to_json(G, comms, str(out / "graph.json"), community_labels=labels)
    (out / "GRAPH_REPORT.md").write_text(report, encoding="utf-8")
    (out / ".graphify_labels.json").write_text(
        json.dumps({str(k): v for k, v in labels.items()}, ensure_ascii=False), encoding="utf-8")
    print(f"Graph: {G.number_of_nodes()} nodes, {G.number_of_edges()} edges, {len(comms)} communities")

    exe = shutil.which("graphify")
    if exe:
        subprocess.run([exe, "export", "html"], cwd=ROOT, check=False)
    else:
        print("graphify CLI not on PATH - skipped graph.html (run: graphify export html)")
    (out / ".graphify_ast.json").unlink(missing_ok=True)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
