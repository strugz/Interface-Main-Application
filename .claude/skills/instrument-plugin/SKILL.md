---
name: instrument-plugin
description: How Lab-I (Interface.exe) loads and talks to instrument plugins (Host_*.dll), data-transfer plugins (DTP_*.dll) and VerSa HIS plugins (VS_*.dll). Use when adding/debugging an analyzer interface, a new message format, a plugin ProgID/registration problem (error 429), or result parsing.
---

# Lab-I plugin contract

Plugins are **late-bound ActiveX DLLs** created with `CreateObject`. This repo holds
only the compiled DLLs, in `ComSet/Plugins/` and `New Lab-I System v2/Plugins/`.

- **Driver source:** `D:\1 Project\LIS Driver` (`drivers/Host_<Name>/`). It is a
  separate git repo with its own CLAUDE.md, `build.ps1` and `test.ps1`.
- **Plugin contract, written from the driver side:**
  `LIS Driver\docs\host-plugin-interface.md`.
- **Host_RAPID:** the working source is `LIS Driver\drivers\Host_RAPID_RIGHT`. Before
  touching RAPID, read the Host_RAPID note in this repo's CLAUDE.md.
- **Vendor protocol manuals:**
  `C:\Users\JayBryanCAbaoag\Documents\VuexJaysWayFile\VuexJaysWayFile\Project\Machine Specification\`. Every plugin must be registered with
`regsvr32` (32-bit: `C:\Windows\SysWOW64\regsvr32.exe`). If it is not registered,
`CreateObject` raises **error 429**. `InitPlugin` swallows that error silently
(modMain.bas ~446).

## ProgID rule
ProgID = uppercased DLL filename, with `.DLL` replaced by the class name:
- `Host_ASTM.dll` -> `HOST_ASTM.MessageDef` (global `MsgDef`)
- `DTP_SLMC.dll` -> `DTP_SLMC.MessageDef` (global `ObjDTP`)
- `VS_SQLCON.dll` -> `VS_SQLCON.mFunctions`

The plugin selection UI lists files that match `HOST_*.dll`, `DTP_*.dll` and `VS_*.dll`
(frmSettings2.frm ~1910). The chosen name is stored in the registry.

## Host plugin (`HOST_X.MessageDef`)
Created at modMain.bas `InitPlugin` (~431), mdiMain.frm (~807), frmControlPanel
(~249) and frmSettings2 (~1768).

**Properties:**
- `InstrumentID`, `MessageFormat`, `SectionCode`, `Version`, `CompanyName`, `Comments`.
- `MessageFormat` selects the framing in `frmReceiveData.HandleInput`. Supported
  values: STKS, ASTM, ASTM2, ILABBG, ASYNC, STXETX, ASTM1394, ACTDIFF, CXSYS, LXSYS,
  ILP, ILP2, and IQ200.

**Methods:**
- `showsettings`
- `textDisplay(raw)`
- `segment_OBX(raw)` returns result rows: pipe-separated, rows separated by `<SP>`.
- `segment_oBI(raw)` returns patient/order info rows. It returns `"ERRCONV"` on failure.
- `Flags(2, msg)` is used by the ASYNC format.
- `HL7_MSH/PID/ORC/OBR/OBX(...)`

**Plugin v2 and later, with "FullHandshake" on.** The plugin drives the protocol
itself through:
- `HandleInput(buf)`: returns the bytes to send back.
- `HasQuery`, `CurrentSampleIDCount`, `CurrentSampleID`, `HasResult`
- `RawData([reset])`

See frmReceiveData.frm ~645-678.

**Parsed field positions used by `SaveResult`:**
- OBX: (1) sort, (3) exam code, (5) result, (6) unit, (14) date.
- OBI/ResultOtherInfo: (2) name, (3) sample number, (6) specimen/accession number,
  (7)-(14) demographics, (17) operator, (18) comments.

## DTP plugin (`DTP_X.MessageDef`)
- `ProcessDTP(obx, obi[, raw])`, `UploadResults(header, details)`, `ReInit`,
  `showsettings`
- `ErrorNum`, `ErrorDesc`

It is called after each result is saved, and from frmSearchResult when results are
uploaded.

## VerSa plugin (`VS_X.mFunctions`)
- `Settings`
- `ProcessVSP(header, details)`

frmSearchResult (~1829) feeds it from the stored procedure `vs_ValuesSelect`.

## Debugging a new analyzer
1. Turn on transmission logging: set the registry value `ComSet\Settings\LogTrans=1`.
   Raw traffic then goes to `ComTransLog\<InstID>yyyymmdd.txt`.
2. Unparsed messages are stored raw in the `LabUnprocessed` table (LabBuffer.mdb).
   Use **Tools > Plugin Test** (frmPlugInTest) to replay them through a plugin.
3. Serial settings (port, baud, parity, handshake) come from ComSet
   (`Comset.ComSettings.showsettings`), in registry sections `ComSet\Control 1..4`.
4. Known limitation: a single global `MsgDef` is shared by all four COM ports. Only
   "Control 1" FullHandshake is consulted.
