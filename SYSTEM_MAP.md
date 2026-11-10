# System Map: Lab-I Laboratory Instrument Interface

Lab-I is a Visual Basic 6 desktop system from Marsman Drysdale Medical Products,
Inc. It connects clinical laboratory analyzers (hematology, chemistry, blood gas,
coagulation, urinalysis) over RS-232. It parses their result messages, stores the
results in Access or SQL Server, and lets lab staff review, verify and print
results. It can also forward results to a hospital information system (HIS).

> Line numbers are approximate (~). To navigate by symbol, use the graph in
> `graphify-out/` (see the `graphify-vb6` skill).

---

## 1. Components

```
                         ┌──────────────────────────────────────────────┐
  Analyzers (RS-232) ───▶│ Interface.exe  (New Lab-I System v2)         │
  COM1..COM4  MSComm     │  modMain.Sub Main → mdiMain (MDI shell)      │
                         │  frmReceiveData  ← serial I/O, framing        │
  HL7 files  ───────────▶│  tmHL7_Timer  ← Received\ folder (orders)     │
  (Received\)            │  frmProcess / SaveResult → DB                 │
                         └───┬──────────┬──────────┬──────────┬─────────┘
          CreateObject (late │ bound COM, regsvr32-registered)│
             ┌───────────────┘          │          │          └──────────────┐
             ▼                          ▼          ▼                         ▼
 ┌────────────────────┐  ┌──────────────────────┐ ┌───────────────────┐ ┌──────────────────┐
 │ ComSet.dll         │  │ HostForms_SA.dll     │ │ Plugins\HOST_*.dll│ │ Plugins\DTP_*.dll│
 │ Comset.ComSettings │  │ HostForms_SA.        │ │ HOST_X.MessageDef │ │ DTP_X.MessageDef │
 │ serial-port setup  │  │   ClassForms         │ │ per-analyzer      │ │ push results to  │
 │ (registry ComSet\) │  │ result entry, master │ │ message parser    │ │ HIS / other DB   │
 └────────────────────┘  │ data, Crystal print  │ └───────────────────┘ ├──────────────────┤
                         └──────────────────────┘                       │ Plugins\VS_*.dll │
                                                                        │ VS_X.mFunctions  │
                                                                        │ VerSa HIS export │
                                                                        └──────────────────┘
 ┌──────────────────────────────────────────────┐
 │ HostForms_RE.dll  (HostFomrs_RE)             │  Request entry for the SQL Server LIS
 │ HostForms_RE.ClassForms                      │  (LMS). Interface.exe does NOT
 │ patient search, request, STAT, sample check-in│  reference it; another host app does.
 └──────────────────────────────────────────────┘
```

| Folder | Project → output | Role |
|---|---|---|
| `New Lab-I System v2/` | `HostInterface.vbp` → **Interface.exe** (v3.1.x, Std EXE) | Main application: comms, parsing, storage, search, worklist, backup and licensing. Also the deployment folder: compiled DLLs, Plugins, Reports, AppData. |
| `ComSet/` | `ComSet.vbp` → **ComSet.dll** | COM-port settings dialog and property accessors. |
| `HostForms_SA v2/` | `HostForms_SA.vbp` → **HostForms_SA.dll** (v2.0.x) | Result entry and verification, master data (exams, profiles, normal values, physicians), Crystal report printing. |
| `HostFomrs_RE/` | `HostForms_RE.vbp` → **HostForms_RE.dll** | Lab request entry against SQL Server stored procedures. |
| `HostForms_SA v2 - Copy/` | — | Stale backup of SA (only frmClassifications differs). Do not edit. |
| `Interface/` | — | Reference data: `Laboratory.mdb`, AU test-method and conversion-factor sheets. |

The whole folder is one git repository (private GitHub `strugz/Interface-Main-Application`).
The older per-folder repositories are retired; see CLAUDE.md, "Repository".

---

## 2. Startup sequence (Interface.exe)

`modMain.Sub Main` (modMain.bas ~283) runs these steps in order:

1. **Command-line switches**
   - `-RESETEXPIRATION`
   - `-REPATCHDB`, which sets `UpdatedTable=0`. That makes the SA DLL re-run its schema patch.
   - `-SAFEMODE`, which does nothing.
2. **Read settings and load plugins.** Reads registry settings (ProfileID, Autoprocess,
   Autoprint), then calls `InitPlugin` (~416). That creates `HOST_<X>.MessageDef` as
   `MsgDef` and `DTP_<X>.MessageDef` as `ObjDTP`.
3. **`GetDBSettings` (~454)** builds the connection string for the selected server
   type: LOCAL DATABASE (Jet 4.0), MSSQL (sqloledb), MYSQL (ODBC 3.51) or ORACLE
   (MSDAORA).
4. **`CopyToVistaFolder` (~1555).** On first run, copies the .mdb files and Reports to
   `%ALLUSERSPROFILE%\Laboratory Instrument Interface\`.
5. **Licence check: `IsRegistered` (~2320).**
   - Decrypts `<DataDir>\%USERNAME%@%USERDOMAIN%\Laboratory.dat` (Blowfish).
   - Matches its CPU ID, volume serial and baseboard serial against this machine.
   - Warns when 30, 15 or ≤7 days are left.
   - If the licence is missing or expired, opens `frmLicenseManager`. That form calls
     the mdmpi.com.ph APIs `/api_registerhost` and `/api_gethostlicenses`.
6. **`OpenConnection` (~580).** For a local database only, it also:
   - runs `VerifyRequestDB` to extract the request .mdb files from resources;
   - runs `FixLink` to re-link the Jet tables;
   - runs `setResultReport`.
7. **`App.PrevInstance` check**, then `mdiMain.Show`.
8. **`mdiMain.MDIForm_Load` (~494):**
   - Reads `tbVersion`.
   - Creates `Comset.ComSettings`.
   - Calls `setHostForms` (modMain ~1321). This creates `HostForms_SA.ClassForms` as
     `HostSA` and injects the DB settings and the live ADO connection into it.
   - Auto-connects if `Settings\AutoConnect=1`.

`lab-i*.sys` files are **theme background images**, not licence files.

---

## 3. Result data flow

```
Analyzer ──RS-232──▶ mscMain(n)_OnComm (frmReceiveData ~550)
                      │  ACK/NAK ENQ/STX/ETX/ETB, or plugin.HandleInput (v2 FullHandshake)
                      ▼
                 HandleInput (~829): framing per MessageFormat
                 STKS | ASTM | ASTM2 | ASTM1394 | ILABBG | ASYNC | STXETX | ACTDIFF |
                 CXSYS | LXSYS | ILP | ILP2 | IQ200
                      │                                    ▲ ComTransLog\<Inst>yyyymmdd.txt
                      ▼                                      (if ComSet\Settings\LogTrans=1)
                 processData (~1179)
          ┌───────────┴──────────────────┐
          │ STKS + Autoprocess           │ other formats
          ▼                              ▼
     SaveResult (~1889)            INSERT raw → LabUnprocessed (LabBuffer.mdb)
          ▲                              │
          │                     frmProcess.processData (~319)
          └──── MsgDef.segment_OBX / segment_oBI ◀┘   (plugin parses raw text)
          │
          ▼
  tbLabInstrumentResults (header, ResultID = hex(InstID+date+specimen+name))
  tbLabInstrumentRValues (one row per exam value)
          │
          ├─▶ ObjDTP.ProcessDTP(obx, obi, raw)        → external HIS / DB
          ├─▶ HL7_ORU → HostSpooler\*.txt              (if HL7 AutoWriteResult)
          └─▶ LabTempPrint → HostSA.ResultAutoPrint    (if Autoprint)
```

**Middleware mode** (`Lab-I\Middleware\Enable=1`, branch `feature/middleware-mode`).
`processData` also calls `QueueForMiddleware`, which hands the result to
`modOutbox`. That module builds HL7 ORU messages and keeps a file spool
(`Outbox\<dest>\pending|sent|failed`). It delivers over MLLP (`clsMLLPClient`) or
as files, with retries. It replaces the inline `HL7_ORU` file write. See
`docs/MIDDLEWARE_MODE.md`.

**Orders (host query / bidirectional):**

```
HIS HL7 ORM files → Received\ → tmHL7_Timer (~1351) → SaveHL7toDB (modMain ~2004) → tbLabRequest
tbLabRequest (Status 0 held, 1 queued, 2 sent, 3 resulted)
   → ProcessHL7Request (~2072): build STKS (CRC16) or ASTM order frames
   → sent to the analyzer on query
```

You can also add orders manually with `frmWorkList` or `frmRequest`.

**Review and verification.** The flow runs:
1. `frmSearch` / `frmSearchResult`.
2. `HostSA.showResultEntry` opens `frmResultEntry` (in the SA DLL). That form:
   - flags values H/L against `tbLabClassExams`;
   - checks the differential count;
   - upserts the patient into `tbPatientMaster` / `tbPatientValues`.
3. Crystal report printing. The report template is a BLOB in `tbLabReportNames`.

---

## 4. Databases

**LOCAL DATABASE mode.** `AppData/Data/Laboratory.mdb` is the hub. It holds linked
tables (relinked by `FixLink`) that point at these files:

| File | Tables |
|---|---|
| Laboratory.mdb | tbcolabexam, tbLabProfiles, tbLabProfileExams, tbLabClassifications, tbLabClassExams, tbLabReportNames, tbVersion, tbHospitalInfo, tbPhysician, tbMedtech, tbLabPathologists, tbLabSpecimen, tbLabRemTemplate, tbSelectionType/Values, tbSequence |
| InstrumentResults.mdb | tbLabInstrumentResults, tbLabInstrumentQC |
| InstrumentRValues.mdb | tbLabInstrumentRValues |
| LabPatientData.mdb | tbPatientMaster |
| LabPatientValues.mdb | tbPatientValues |
| LabBuffer.mdb | LabUnprocessed, LabTempPrint |
| LabRequestBuffer.mdb | tbLabRequest |
| LabRequestForm.mdb | tbLabRequestForm |

**MSSQL mode.** The same table names live in one SQL Server database, accessed
through the single global ADO connection `myConnection`.

HostForms_RE uses SQL Server only, through stored procedures:
- `sp_PatientDemo_*`
- `sp_tbLabRequestMaster_*`, `sp_tbLabRequestValues_*`
- `sp_CheckInBySmpl_update`, `sp_LockRequest_*`
- `sp_tbQryExternalDB_*`

Its settings are in `tblabsettings` (category `RE_REQ`).

**Schema patching.** The `HostForms_SA.ClassForms.wConnection` setter runs
`ALTER TABLE` and `CREATE PROCEDURE` once. The registry key `UpdatedTable` guards it.

**Backups.** `frmDailyBackup` calls `DailyBackup` (modMain ~2398), which zips the
.mdb files with 7z into `DailyBackUp\` and keeps 15 archives. You can archive and
compact from `frmDbMan` / `frmArchive`.

---

## 5. Configuration (Windows registry)

Everything is stored with `GetSetting`/`SaveSetting` under
`HKCU\Software\VB and VBA Program Settings\`.

| App | Sections (examples) |
|---|---|
| `Lab-I` / `LAB-i` | Settings (PlugIN, AutoConnect, Autoprocess, Autoprint), DBSettings, LOCAL/MSSQL/MYSQL/ORACLE, HL7, HL7 Settings, DaTran, VerSa, OBISettings (FLD_n), Registration, Database, DataLocation, AutoArchive, FormColors |
| `ComSet` | `Control 1`..`Control 4` (ComPort, BaudRate, Parity, DataBits, StopBits, HandShake, buffers, `Plug-Ins`, FullHandshake, Enabled), `Settings` (LogTrans), legacy `PortSettings` |
| `HostForms_SA` / `HostForms` | UpdatedTable, DiffCount, ByPassDiff, form settings |

`ComSet.ini` and `ComSet.mdb` are empty and unused. The DB credentials in the
registry are Blowfish-encrypted with a hardcoded key.

**OBI (patient lookup) settings.** `frmOBISetiings` configures an external patient
lookup, which `frmPIDsearch` in the SA DLL uses. The source can be MSSQL, MySQL,
Informix or a REST API (JSON, sent as raw HTTP over Winsock).

---

## 6. Plugin contracts

See `.claude/skills/instrument-plugin/SKILL.md` for the full list.

Related locations outside this folder:
- **Plugin source:** `D:\1 Project\LIS Driver`, under `drivers/Host_<Name>/`.
- **Contract:** `LIS Driver\docs\host-plugin-interface.md`.
- **Vendor protocol manuals:** `Documents\VuexJaysWayFile\VuexJaysWayFile\Project\Machine Specification\`.

In short:

**`HOST_X.MessageDef`**
- Properties: `InstrumentID`, `MessageFormat`, `SectionCode`, `Version`.
- Methods: `segment_OBX`, `segment_oBI`, `textDisplay`, `showsettings`, `HL7_*`.
- In v2 and later: `HandleInput`, `HasQuery`, `CurrentSampleID*`, `HasResult`, `RawData`.

**`DTP_X.MessageDef`**
- `ProcessDTP`, `UploadResults`, `ReInit`, `ErrorNum`, `ErrorDesc`.

**`VS_X.mFunctions`**
- `Settings`, `ProcessVSP(header, details)`.

**Other late-bound COM servers** (not in this repo): `SearchCodes.ClassSearch`
(used by RE) and `BCSET.Barcode` (barcode printing in RE).

---

## 7. Main UI (mdiMain menus)

| Menu | Items |
|---|---|
| File | Search, Statistics, Requests, Worklist, QC |
| Connection | Receive only, Send/Receive, Send ID, Disconnect |
| Maintenance | `HostSA.show*` master-data screens |
| Reports | Worklist, Logbook, Other, Reprint |
| Settings | Program Setup (frmControlPanel), Program Settings (`frmSettings2.frm`, VB_Name **frmSettings**), Form, Com, DB and Init DTP settings |
| Tools | Barcode, Clear Buffer, Archive, DB Maintenance (frmDbMan), Plugin Test, Statistics |
| Utilities | Console (frmReceiveData commands: OPENDB, CLOSEDB, CONFIG, LOOPBACK, SAVELOG, SEND, …) |
| Help | License Manager |

---

## 8. Known risks and tech debt (prioritised)

1. **SQL injection everywhere.** Instrument, HL7 and user data are concatenated
   into SQL with no parameters. Examples: frmReceiveData ~1920-2041, frmProcess
   ~165-296, modMain ~1932/2050/2142, and the RE ClassForms.
2. **Hardcoded secrets.** These include the Blowfish key (modMain.bas ~10,
   modClassMain.bas ~4), the Jet DB passwords (modMain, SA ClassForms ~186/222) and
   the DSN password. `Passwords.txt` sits in the source folder. Values are
   deliberately not reproduced here.
3. ~~**Data loss: monthly auto-archive.**~~ *Fixed 2026-11-10.* The monthly
   auto-archive used to delete all `LabUnprocessed` rows. Now `ArchiveUnprocessed`
   first writes them to `UnprocessedArchive\<Inst>_<timestamp>.txt` in the data
   folder, then deletes only the rows it wrote.
4. ~~**Data loss: transaction leak.**~~ *Fixed 2026-11-10.* Both `SaveResult` copies
   now roll back on every error, including error 9. When a save fails, the raw
   message stays in, or is put back into, `LabUnprocessed` for reprocessing.
   `ERRCONV` rows are still deleted, as before.
5. **Per-port state.** All four ports share one global `MsgDef` and one set of
   handshake flags. Only Control 1 FullHandshake is honoured.
6. **COM10+ truncated.** `ComSettings.ComPort` uses `Mid(...,4,1)`.
7. **OS policy changes.** `frmSettings2` changes OS policy (DisableTaskMgr,
   DisableRegistryTools, USBSTOR).
8. **Blocking busy-waits.** `DoEvents` busy-waits block the UI and serial events, in
   frmReceiveData and the licence/REST calls.
9. **Debug leftover.** `frmLicenseManager` ~367 forces `RemoteHost="localhost"`.
10. **Dead code and copies:**
    - Nine orphan files in `New Lab-I System v2` that are not in the .vbp.
    - `HostForms_RE/frmRequest.frm`, which is fully commented out.
    - `*_bak`, `*_` and `- Copy` forms.
    - The `.gitignore` in `New Lab-I System v2` still contains unresolved
      merge-conflict markers.
11. **External build dependencies.** The build needs the Crystal Reports 8 runtime
    registered on the build PC. (vbGraph is now in the repo under `HostForms_SA v2/vbGraph/`.)
