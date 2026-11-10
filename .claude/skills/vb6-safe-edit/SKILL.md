---
name: vb6-safe-edit
description: Rules for safely reading and editing the VB6 source in this repo (.frm, .bas, .cls, .vbp, .ctl). Use before changing any Visual Basic 6 file here - encoding, CRLF, designer blocks, .frx companions, VB_Name traps, and which files are dead copies.
---

# Editing VB6 source safely

The VB6 IDE is unforgiving about file format. A file that "looks fine" in a text
editor can fail to load or silently lose controls. Follow these rules for every edit.

## File format
- **Encoding is Windows-1252 (ANSI), not UTF-8.** Never write a BOM. Avoid
  non-ASCII characters; if you must, make sure they survive as cp1252.
- **Line endings are CRLF.** Keep them. LF-only lines make the IDE merge lines or
  refuse to load forms. After editing with a tool that might normalise endings,
  check with `file <path>` (should say "CRLF line terminators").
- Keep `Attribute VB_Name = "..."` and the other `Attribute` lines exactly as they
  are. They must stay at the top of the code section.

## Forms (.frm) and user controls (.ctl)
- A `.frm` has two parts: the **designer block** (`VERSION 5.00`, `Begin VB.Form ... End`)
  and then the code after the `Attribute VB_*` lines. Edit only the code part unless
  the task is explicitly a UI change.
- Binary properties (images, list items, some grid settings) live in the matching
  **`.frx`** file, referenced as `"frmX.frx":0000` offsets. Never hand-edit `.frx`.
  Never reorder or delete designer properties that reference `.frx` offsets.
- Adding a control by hand is risky. Prefer to describe the change for the developer
  to make in the VB6 IDE.

## Naming traps in this repo
- `New Lab-I System v2/frmSettings2.frm` declares `VB_Name = "frmSettings"`. Code
  that calls `frmSettings` opens **frmSettings2.frm**, not `frmSettings.frm`.
- The authoritative file list is the `.vbp`. Files that sit in the folder but are
  not listed in the `.vbp` are dead and are not compiled. In New Lab-I System v2 these
  are: frmSettings.frm, frmSplash, frmRptViewer, frmProgramSettings, frmLabExam,
  frmLabProfileExams, frmDbCopy, frmDBdataCopy, modCRC.bas. Check the `.vbp` before
  editing.
- `HostForms_SA v2 - Copy/`, the `*.bak` files, `frmLabExam_bak.frm`,
  `frmResultEntry_.frm` and the `*-L-WRFABROA.vbp` / `*-Luis-PC.vbp` variants are
  backups or developer-machine copies. Do not edit them.
- The folder `HostFomrs_RE` is misspelled on purpose (historical). Do not rename it,
  because other paths depend on it.

## Code conventions
- Keep `Option Explicit`. Declare every new variable with an explicit type.
- Error handling follows the `On Error GoTo ErrHandler` pattern. Any procedure that
  calls `BeginTrans` must call `RollbackTrans` in its error path.
- SQL is built by string concatenation. For any new or touched SQL, at least double
  single quotes (`Replace(x, "'", "''")`) on values that come from instruments,
  HL7 files or users.
- Settings live in the registry via `GetSetting`/`SaveSetting` (app names `Lab-I`,
  `ComSet`, `HostForms_SA`, `HostForms`). Reuse the existing section and key names.

## Secrets
Never print or copy the values of `myKey`, the Jet database passwords or
`Passwords.txt`. Refer to them by file and line only.

## After editing
- Update the graph so it stays accurate: run the `graphify-vb6` skill.
- Compiling needs the VB6 IDE (see the `vb6-build` skill). You cannot verify a build
  from here unless VB6.EXE is available, so say so when you report.
