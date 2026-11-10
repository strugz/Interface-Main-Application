# CLAUDE.md: Lab-I Laboratory Instrument Interface

This is a legacy **Visual Basic 6** system (32-bit, Windows). It connects clinical
laboratory analyzers to a database and a HIS. Read `SYSTEM_MAP.md` for the
architecture, data flow, databases, registry settings and known risks.

## Layout
- `New Lab-I System v2/`: **Interface.exe** (`HostInterface.vbp`). This is the main
  app and also the deployment folder.
- `ComSet/`: **ComSet.dll**, the serial-port settings component.
- `HostForms_SA v2/`: **HostForms_SA.dll**, for result entry, master data and
  Crystal reports.
- `HostFomrs_RE/`: **HostForms_RE.dll**, for SQL Server request entry. The folder name
  is misspelled; keep it as is.
- `HostForms_SA v2 - Copy/`, `*_bak*`, `*.bak` and the `*-L-WRFABROA.vbp` /
  `*-Luis-PC.vbp` files are backups. **Never edit them.**
- `tools/`: graphify tooling for VB6 (`graphify_vb6.py`, `build_graph.py`).
- `graphify-out/`: the knowledge graph (`graph.json`, `GRAPH_REPORT.md`,
  `graph.html`).

The `.vbp` file is the source of truth for which modules are compiled. Several
`.frm`/`.bas` files in the folders are not listed in it and are dead.

## Related projects (outside this folder)

| What | Where | Link to this repo |
|---|---|---|
| **Instrument driver plugins (source)** | `D:\1 Project\LIS Driver` (its own git repo and CLAUDE.md) | Builds the `Host_<Name>.dll` files that Interface.exe loads from `New Lab-I System v2\Plugins\` as `CreateObject("HOST_<NAME>.MessageDef")` |
| **Analyzer LIS / interface manuals** | `C:\Users\JayBryanCAbaoag\Documents\VuexJaysWayFile\VuexJaysWayFile\Project\Machine Specification` | Vendor protocol specs, one per analyzer. Example: `RAPIDPoint_500e_Interface_Specifications_Manual_-_English_-_Rev_B_*.pdf` for Host_RAPID |

How they fit together:
- **This repo is the host.** It runs the serial link and calls a plugin through the
  fixed `MessageDef` contract (`InstrumentID`, `MessageFormat`, `HandleInput`,
  `segment_OBX`, `segment_oBI`, ...). It then saves what the plugin returns. See
  `.claude/skills/instrument-plugin`.
- **LIS Driver is the plugin side.**
  - It holds `drivers/Host_<Name>/` (VB6 source) and `output/Host_<Name>/` (the
    built DLL).
  - `docs/host-plugin-interface.md` is the contract, written from the plugin side.
  - Build with `.\build.ps1 Host_<Name>`; test with `.\test.ps1`.
- **Field layouts.** The OBX/ORI layouts LIS Driver documents are the ones
  `SaveResult` reads here: `OBX(3)` exam, `(5)` result, `(6)` unit, `(14)` date;
  `ORI(2)` name, `(3)` sample number, `(6)` sequence/accession.
- **Changing that contract affects both repos.** It is tied to `SaveResult`,
  `processData`/`HandleInput` in `frmReceiveData.frm` and `frmProcess.frm`. Check
  both repos before changing it.
- **Deploying a driver.**
  1. Copy `LIS Driver\output\Host_<Name>\Host_<Name>.dll` into
     `New Lab-I System v2\Plugins\`.
  2. Register it with `%SystemRoot%\SysWOW64\regsvr32`.
  3. Select it in Settings → Program Settings.
- **Debugging an analyzer problem.**
  1. Read the manual in `Machine Specification`.
  2. Read the driver in `LIS Driver\drivers\Host_<Name>`.
  3. Read the raw traffic in `New Lab-I System v2\ComTransLog\<InstrumentID>yyyymmdd.txt`
     (needs `ComSet\Settings\LogTrans=1`).
  - **Host_RAPID.** The source of the working driver is
    `LIS Driver\drivers\Host_RAPID_RIGHT`. The folder name has `_RIGHT`, but the
    project and DLL are `Host_RAPID`.
    - Build it with `.\build.ps1 Host_RAPID_RIGHT`. The result is
      `output\Host_RAPID\Host_RAPID.dll`, v3.1 or later.
    - It needs FullHandshake = 1. It answers `SMP_NEW_AV` with `SMP_REQ`, echoing
      the analyzer's own `iIID`, and ACKs `SMP_NEW_DATA`.
    - Its InstrumentID is set in its settings dialog (Settings → plugin
      settings). The default is `RAPID`, which Lab-I's RAPID branch in
      `frmReceiveData.processData` depends on.
    - Older builds you may find are **not** that source:
      - `New Lab-I System v2\Plugins\Host_RAPID.dll`, v1.0.11. It hardcodes IID
        `59835` in `SMP_REQ`, so a RAPIDPoint with any other ID never sends its
        samples. Replace it with the LIS Driver build.
      - The 3.0.20 DLL inside `RapidPoint.rar`.
      - `D:\LIS\...\Host_RAPID.dll`, v1.0.27.

## Build and release
- **"Build X"** means compile locally and stop. Run `.\tools\build.ps1` (or
  `-Projects HostForms_SA`, and so on). The output lands in `New Lab-I System v2\`. Don't
  tag or publish anything.
- **"Release" / "publish" / "release X for support"** means follow the
  `.claude/skills/release` skill:
  1. Build from a clean `main`.
  2. Commit the `.vbp` version bumps.
  3. Tag `v<Interface version>`, or `<Component>-v<version>` when Interface.exe
     isn't included.
  4. Push.
  5. Publish a GitHub release with the built files and support-facing notes.
- **Support staff download releases from
  https://github.com/strugz/Interface-Main-Application/releases**, not from
  Viber or chat. Always give the user the release link when you finish.
- **What a release contains.** Releases contain only `Interface.exe`,
  `HostForms_SA.dll`, `ComSet.dll` (and `HostForms_RE.dll` when asked).
  - Instrument plugins are released from LIS Driver.
  - Never attach databases, logs, `.reg` exports or `Passwords.txt`.

## Repository
- The whole folder is one git repo: private GitHub `strugz/Interface-Main-Application`, branch `main`.
- `.gitignore` keeps out databases, analyzer logs (`ComTransLog`, `HostSpooler`, `Received`), HL7 sample messages, passwords, registry exports, compiled binaries and archives. Never force-add them: they contain patient data or secrets.
- `.gitattributes` pins VB6 sources to CRLF and marks `.frx`/`.rpt`/`.res` as binary.
- The old per-folder repos (GitHub `strugz/IntefaceMainApplication` for New Lab-I System v2, Bitbucket `hostforms_re`) are retired. Their `.git` folders are in `D:\1 Project\Interface Main Application.git-backup\`.

## Rules
- **Before editing any VB6 file, follow `.claude/skills/vb6-safe-edit`.** Files are
  cp1252 with CRLF line endings. Do not touch designer blocks or `.frx` files. Keep
  the `Attribute` lines.
- `frmSettings2.frm` has `VB_Name = "frmSettings"`. Calls to `frmSettings` open that
  file.
- Components talk only through **late-bound COM** (`CreateObject("Lib.Class")`). Keep
  public signatures of `ComSet.ComSettings`, `HostForms_SA.ClassForms`,
  `HostForms_RE.ClassForms` and the plugin contracts backward-compatible. Installed
  sites mix versions.
- Settings are in the registry (`GetSetting`/`SaveSetting`, apps `Lab-I`, `ComSet`,
  `HostForms_SA`). Reuse the existing sections and keys.
- SQL is built by string concatenation. In code you touch, escape quotes, and pair
  `BeginTrans` with `RollbackTrans` in error handlers.
- **Secrets:** never print, copy or commit the values of `myKey`, the Jet/DSN
  passwords, `Passwords.txt`, `InterfaceRegistry.reg` or `Laboratory.dat`. Refer to
  them by file and line only.
- Do not read or modify `.mdb` databases, `AppData/Data`, `ComTransLog/`,
  `HostSpooler/` or `Received/` unless asked. They contain patient data.
- You cannot compile here unless VB6.EXE and all references are available. Say so
  instead of claiming a build passed.

## Project skills (`.claude/skills/`)
- `vb6-safe-edit`: file-format rules and traps for editing VB6 source.
- `vb6-build`: build order, `VB6.EXE /make`, regsvr32, external references.
- `instrument-plugin`: Host_/DTP_/VS_ plugin contracts and how to debug an analyzer.
- `graphify-vb6`: rebuild and query the knowledge graph.

## graphify
This project has a graphify knowledge graph at `graphify-out/`.

Rules:
- Before answering architecture or codebase questions, read
  `graphify-out/GRAPH_REPORT.md` for god nodes and community structure.
- For cross-module "how does X relate to Y" questions, prefer
  `graphify query "<question>"`, `graphify path "<A>" "<B>"` or
  `graphify explain "<concept>"` over grep. These traverse the graph's EXTRACTED and
  INFERRED edges instead of scanning files. Set `PYTHONIOENCODING=utf-8` first on
  Windows.
- graphify's own `/graphify .` cannot parse VB6. After modifying `.frm`, `.bas`,
  `.cls` or `.vbp` files, rebuild with:
  `& (Get-Content graphify-out\.graphify_python) tools\build_graph.py`
  (see the `graphify-vb6` skill).
