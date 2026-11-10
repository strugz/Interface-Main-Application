# Lab-I: Laboratory Instrument Interface

Lab-I is a Windows desktop system that connects clinical laboratory analyzers to a
laboratory database and the hospital information system. It was originally written
by Willie R. Fabroa for Marsman Drysdale Medical Products, Inc. and is built in
Visual Basic 6.

It does six things:
- **Receives** results from analyzers over RS-232. It supports ASTM, STKS, HL7-style
  and vendor formats through per-analyzer plugins.
- **Sends** worklists and orders to bidirectional analyzers. Orders come from HL7
  files or are entered manually.
- **Stores** results in Microsoft Access (Jet) or SQL Server. MySQL and Oracle
  connection strings also exist.
- **Reviews, verifies and prints** results. Values are flagged against reference
  ranges, and reports are printed with Crystal Reports 8.
- **Forwards** results to HIS and EMR systems through DTP and VerSa plugins, or as
  HL7 ORU files.
- **Backs up and archives** the database daily, and is licensed per machine.

For the architecture, data flow, databases, configuration and known issues, see
**[SYSTEM_MAP.md](SYSTEM_MAP.md)**.

## Repository layout

| Folder | Builds | Description |
|---|---|---|
| `New Lab-I System v2/` | `Interface.exe` | Main application. It is also the run folder: DLLs, `Plugins/`, `AppData/Reports`, `AppData/Data` |
| `ComSet/` | `ComSet.dll` | Serial (COM1–COM4) port configuration |
| `HostForms_SA v2/` | `HostForms_SA.dll` | Result entry, master data (exams, profiles, normal values, physicians), reports |
| `HostFomrs_RE/` | `HostForms_RE.dll` | Lab request entry for the SQL Server LIS |
| `HostForms_SA v2 - Copy/` | — | Old backup copy, not used |
| `Interface/` | — | Reference spreadsheets and a sample `Laboratory.mdb` |
| `tools/` | — | Knowledge-graph tooling (graphify for VB6) |
| `graphify-out/` | — | Generated knowledge graph and report |

## Requirements
- Windows. All components are 32-bit and are registered with
  `C:\Windows\SysWOW64\regsvr32.exe`.
- Visual Basic 6.0 SP6 (`C:\Program Files (x86)\Microsoft Visual Studio\VB98\VB6.EXE`).
- Crystal Reports 8 runtime (`craxdrt.dll`, `crviewer.dll`).
- ActiveX controls: MSCOMM32, MSCOMCTL, MSCOMCT2, MSHFLXGD, TABCTL32, COMDLG32,
  RICHTX32, MSWINSCK.
- ADO 2.8, DAO 3.6, Microsoft Scripting Runtime.
- Files from outside this repo that `HostForms_SA` needs:
  `..\..\LIS Main Application\vbGraph\` and
  `..\..\..\..\Laboratory Management System\Components\craxdrt.dll`.

## Build

Build the DLLs first, then the EXE:

```powershell
$vb6 = "C:\Program Files (x86)\Microsoft Visual Studio\VB98\VB6.EXE"
& $vb6 /make "ComSet\ComSet.vbp" /out build.log
& $vb6 /make "HostForms_SA v2\HostForms_SA.vbp" /out build.log
& $vb6 /make "HostFomrs_RE\HostForms_RE.vbp" /out build.log
& $vb6 /make "New Lab-I System v2\HostInterface.vbp" /out build.log
```

`ComSet.dll` and `HostForms_SA.dll` are written to `New Lab-I System v2\`. Register
the DLLs and the plugins (`Plugins\*.dll`) with regsvr32, run as administrator. For
details and troubleshooting, see `.claude/skills/vb6-build/SKILL.md`.

## Run and configure
1. Start `New Lab-I System v2\Interface.exe`. On first run, data and reports are
   copied to `%ALLUSERSPROFILE%\Laboratory Instrument Interface\`.
2. If no valid licence is found, **Help → License Manager** registers this machine.
   The licence is bound to the machine's hardware IDs.
3. Configure the program:
   - **Settings → Program Settings:** choose the instrument plugin (`HOST_*.dll`) and
     the DTP plugin.
   - **Settings → Com Settings:** set the port, baud rate, parity and handshake for
     each control (port).
   - **Settings → DB Settings:** choose Local (Access) or SQL Server.
4. Connect with **Connection → Receive Only** or **Send/Receive**.

All settings are stored in the registry under
`HKCU\Software\VB and VBA Program Settings\Lab-I` and `...\ComSet`.

## Troubleshooting a new analyzer
- Set the registry value `ComSet\Settings\LogTrans = 1` to log raw traffic to
  `ComTransLog\<InstrumentID>yyyymmdd.txt`.
- Messages that were not parsed stay in the `LabUnprocessed` table. Use
  **Tools → Plugin Test** to replay them through a plugin.
- Error 429 ("ActiveX can't create object") means a plugin or DLL is not registered.

## Knowledge graph (graphify)

`graphify-out/` contains a graph of the whole codebase:
- 1,689 nodes and 3,055 edges.
- It covers modules, procedures, calls, form references, COM ProgIDs and database
  tables.

To use it:
- **Browse:** open `graphify-out/graph.html` in a browser.
- **Summary:** read `graphify-out/GRAPH_REPORT.md`.
- **Query:**

  ```powershell
  $env:PYTHONIOENCODING = "utf-8"
  graphify query "how are instrument results saved"
  ```

- **Rebuild after code changes:**

  ```powershell
  & (Get-Content graphify-out\.graphify_python) tools\build_graph.py
  ```

graphify cannot parse VB6 on its own. `tools/graphify_vb6.py` is the parser for this
repo. To install graphify, run `uv tool install graphifyy`.

## Security notes
This codebase contains hardcoded encryption keys and database passwords, and a
`Passwords.txt` file in the main folder. It also builds SQL by string
concatenation. Treat the source as confidential. Do not publish it. Fix these issues
before exposing any component to untrusted input. See SYSTEM_MAP.md §8.
