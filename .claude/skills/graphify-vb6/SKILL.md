---
name: graphify-vb6
description: Rebuild or query the graphify knowledge graph of this VB6 repo. Use when asked to rebuild/update the graph, after editing .frm/.bas/.cls/.vbp files, or to answer "what calls X / what uses table Y / how does A reach B" questions. Plain `/graphify .` does NOT parse VB6 - use this instead.
---

# graphify for this VB6 repo

graphify has no built-in Visual Basic 6 parser. On this repo, `/graphify .` only
finds .cls files (and parses them wrongly) plus log .txt noise. Instead, this repo
ships its own extractor:

- `tools/graphify_vb6.py` parses the four main `.vbp` projects and every module
  they list. It produces procedures, calls, form references, `CreateObject` ProgIDs
  (linked across projects), `Declare ... Lib` DLLs and SQL table usage.
- `tools/build_graph.py` runs that extractor, then graphify's build, cluster, report
  and HTML export.

## Rebuild (PowerShell, from the repo root)

```powershell
$py = Get-Content graphify-out\.graphify_python   # interpreter that has graphify
if (-not $py) { $py = Join-Path (uv tool dir).Trim() "graphifyy\Scripts\python.exe" }
$env:PYTHONIOENCODING = "utf-8"
& $py tools\build_graph.py
```

This writes `graphify-out/graph.json`, `GRAPH_REPORT.md` and `graph.html`. No API key
and no LLM are needed. The build is deterministic and takes a few seconds.

If graphify is missing, install it with `uv tool install graphifyy`, or with
`pip install graphifyy` if uv is not available.

## Query

```powershell
$env:PYTHONIOENCODING = "utf-8"
graphify query "how does frmReceiveData save results" --budget 1500
graphify path "frmReceiveData.SaveResult()" "table tbLabInstrumentRValues"
graphify explain "modMain.setHostForms()"
```

Node labels look like this:
- Procedures: `Module.Proc()`
- Modules: `frmX (frm)`
- Tables: `table tbName`
- ProgIDs: `ProgID "Lib.Class"`
- Projects: `HostInterface (Interface.exe, Exe)`

Always confirm a graph answer against the source file and line (`source_location`)
before relying on it.

## Limits
- Call edges are matched by name inside one project. A late-bound call through a
  COM object variable (for example `HostSA.showResultEntry` or
  `MsgDef.segment_OBX`) shows up as an `instantiates` edge to the ProgID, not as a
  call edge to the plugin.
- `uses_table` edges are INFERRED from SQL string literals. Table names built from
  variables are missed.
- The instrument plugin DLLs (`Plugins/*.dll`) are binaries. Their source is not in
  this repo.
- If you add a new `.vbp` project, add it to `PROJECTS` in `tools/graphify_vb6.py`.
