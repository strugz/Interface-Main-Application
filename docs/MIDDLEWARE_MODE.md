# Middleware Mode: Design

Status: **Phase 1 implemented on branch `feature/middleware-mode`. It has not been
compiled or tested on the build PC yet.** Date: 2026-11-10.

## Phase 1 as built

| Piece | Where |
|---|---|
| Switch: `Lab-I\Middleware\Enable = 1` forces Autoprocess=1 and Autoprint=0 | `modMain.InitPlugin` |
| Menus and toolbar buttons hidden; caption shows `[MIDDLEWARE]` | `mdiMain.ApplyMiddlewareMenus` |
| Results queued after each message (instead of the old inline HL7 file write) | `frmReceiveData.processData` → `QueueForMiddleware` → `modOutbox.OutboxEnqueueResult` |
| Generic HL7 ORU^R01 builder from the OBX/ORI contract fields | `modOutbox.MW_BuildORU` |
| Spool, retry with backoff, failed folder, prune of `sent\` | `modOutbox` (`OutboxPump`, run every second from `tmrStatus`) |
| HL7 over TCP (MLLP) with ACK check | `clsMLLPClient`, bound to the unused `frmReceiveData.Winsock1` (no designer change) |
| HL7 file drop: temp name, then rename | `modOutbox.PumpFile` |
| Status: status-bar panel 7, plus console commands `OUTBOX` and `RESEND [HL7TCP\|HL7FILE]` | `modOutbox`, `frmReceiveData.consoleCmd` |

**Differences from the design below:**
- **Search is hidden in phase 1.** The search screen also edits, deletes, merges,
  prints and uploads results, so it cannot be made read-only cheaply. The
  read-only audit view comes with the Outbox screen in phase 2.
- **Maintenance stays partly visible:** Exams, Normal Values, Classifications and
  Profiles are kept. The people master data (hospital, pathologist, physician,
  med. tech.), flag lists, selection lists and report formats are hidden.
- **Test code mapping uses the existing columns:** `tbcolabexam.HisExamCode` →
  OBX-3 code, `DisplayName` → OBX-3 text, `ExamNotes` → OBX-7 reference range
  (written as `male|female`, or a single range). Unmapped tests are sent with the
  analyzer code when `SendUnmapped = 1`, the default. Set it to 0 to drop them.
- **Database not connected:** the result is not forwarded, because there is no
  code map. The raw message is written to `<Outbox>\unqueued\`. Messages that hit
  a plugin or build error go there too.
- **DTP** still runs inline, as today. It moves into the outbox in phase 2.

## Phase 2 as built: Middleware Settings + Outbox screen

`frmMiddleware.frm` opens from **Settings → Middleware Settings...** or from
**Program Setup → 17 - Middleware (LIS/HIS)**. It works in both modes, so a site
can be set up before switching.

**Settings tab:**
- **Mode:** shows the current mode. A **Switch to … Mode** button runs the same
  restart switch as the menu (`mdiMain.SwitchModeWithRestart`).
- **HL7 over TCP:** on/off, host, port, ACK timeout and keep-alive.
  **Test Connection** opens a TCP connection only. No HL7 is sent, so no test
  result can end up in the LIS.
- **HL7 files:** on/off, folder, extension and CR+LF. The folder is the shared
  `HL7 Settings\FWDIR`.
- **HL7 message:** sending and receiving application (the shared `HL7 Settings`
  keys), the two facility names, and whether to send unmapped tests.
- **Delivery:** first retry delay, maximum attempts, days to keep sent messages,
  and the Outbox folder.
- **Validation:** the port must be 1–65535 and the numbers within range. It warns
  if the file folder can't be reached or no destination is on.
- **Unsaved edits** are offered for saving on Close, on the X button, or before
  switching mode.
- Saved values apply to the next message without a restart. Only the mode change
  restarts Lab-I.

**Outbox tab:**
- Pick a destination (TCP or files) and a folder (Pending, Failed or Sent).
- Counts refresh every 5 s. The list shows the newest 500 messages: time, sample
  number, control ID, tries, status (waiting / retry at … / FAILED / Sent) and the
  last result or error.
- **Resend Selected** puts failed or sent messages back in the queue. For sent
  messages it warns that the LIS may get duplicates. On Pending the button becomes
  **Retry Now**.
- **View** (or double-click) shows the exact HL7 message that was built.
- It reads the spool files only. No database access.

Like `mnuSwitchMode`, the new form and the `mnuMiddleware` menu block were written
by hand. Open both in the VB6 IDE once on the build PC and save from the IDE, so
VB normalises the designer blocks.

### Switching modes (Settings menu)

**Main mode** is the normal Lab-I and the default. **Settings → Switch to Middleware
Mode...** changes to middleware. In middleware mode the same menu item reads
**Switch to Main Mode...**.

1. If the instrument is connected, Lab-I asks to disconnect first.
2. It asks for confirmation.
3. It saves the mode (`modMain.SetMiddlewareMode`):
   - **To Middleware:** the current Main-mode `Autoprocess` / `Autoprint` values are
     saved as `Middleware\SavedAutoprocess` / `SavedAutoprint`, then `Enable = 1`.
   - **Back to Main:** `Enable = 0`, and `Autoprocess` / `Autoprint` are put back
     from the saved values. The saved keys are then deleted.
4. Lab-I restarts by itself, so menus and settings are applied cleanly.

The destination settings (`HL7Tcp`, host, port, …) are kept when you switch back
to Main mode, so switching to Middleware again needs no new setup. Messages left
in the Outbox are sent the next time Middleware mode is on.

The menu item is a new `VB.Menu` block in the `mdiMain.frm` designer: `mnuSep06`
and `mnuSwitchMode`, after "Initialize DTP...". It is plain text with no `.frx`
data. Open the form once in the VB6 IDE on the build PC to confirm it loads.

### Registry: `HKCU\...\Lab-I\Middleware`

| Key | Default | Meaning |
|---|---|---|
| `Enable` | 0 | 1 = middleware mode |
| `HL7Tcp` | 0 | 1 = send over MLLP |
| `HL7TcpHost` / `HL7TcpPort` | — | LIS listener |
| `AckTimeout` | 10 | Seconds to wait for the connection and for the ACK |
| `KeepAlive` | 1 | Keep the socket open between messages |
| `HL7File` | 0 | 1 = write files to `HL7 Settings\FWDIR` |
| `HL7FileExt` | hl7 | File extension (file name = MSH-10 control ID) |
| `HL7FileCRLF` | 1 | Segment end in files: 1 = CRLF (as the old writer did), 0 = CR |
| `RetrySeconds` | 30 | First retry delay. It doubles on each retry, up to a 600 s cap |
| `MaxAttempts` | 10 | After this many tries the message moves to `failed\` |
| `KeepSentDays` | 90 | `sent\` files older than this are deleted (once a day) |
| `OutboxDir` | data folder `\Outbox` (LOCAL), else `App.Path\Outbox` | Spool root |
| `SendUnmapped` | 1 | Send tests that have no `HisExamCode` under the analyzer code |
| `SendingFacility` / `ReceivingFacility` | empty | MSH-4 / MSH-6 |

MSH-3 and MSH-5 come from the existing `HL7 Settings` keys "Sending Application"
and "Receiving Application". MSH-12 comes from "HL7 Version" (default 2.5.1).

### Message layout

- `MSH|^~\&|<SendApp>|<SendFac>|<RecvApp>|<RecvFac>|<now>||ORU^R01|<yymmddhhnnss+seq>|P|<ver>`
- `PID|1||<ORI(7) patient ID, else ORI(3) sample no.>||<Last>^<First>||<DOB>|<Sex>`
- `OBR|1|<ORI(6) seq/accession>|<ORI(3) sample no.>|<Section>^<InstrumentID>^L|||<obs date>||||||||<ORI(12) specimen>|<ORI(14) doctor>||||||<now>||<Section>|F`
- `OBX|n|NM or ST|<HisExamCode>^<DisplayName>^L||<value>|<unit>|<range>|<flag>|||F|||<date>||||<InstrumentID>`

Segments end with CR. Special characters are escaped (`\F\ \S\ \R\ \E\ \T\`).
For MLLP, MSA-1 = AA or CA counts as accepted. MSA-2 must match MSH-10 when the
LIS sends it.

### How to test (build PC)

1. Build with `.\tools\build.ps1 -Projects Interface`.
2. Set `Enable=1`, `HL7Tcp=1`, `HL7TcpHost=127.0.0.1` and `HL7TcpPort=6661`.
   Point it at any MLLP listener, for example the HAPI TestPanel or 7edit.
3. Connect, then send a sample from the analyzer or a simulator. Check that the
   status bar shows `Outbox: TCP 0 pend/0 fail` and that the listener got the
   ORU.
4. Stop the listener and send again. You should see the retry lines, then the
   message in `failed\` after `MaxAttempts`. Restart the listener, type `RESEND`,
   and check that it is delivered.
5. Set `HL7File=1`. The same results should appear in `FWDIR` as `.hl7` files.

---

# Original design

## Goal

Let Interface.exe run as **middleware only**. It receives results from the
analyzer and delivers them to the customer's LIS/HIS. In this mode Lab-I does
not verify results, print patient reports or keep master data.

```
Analyzer ──RS-232──▶ Interface.exe (Middleware mode) ──▶ LIS / HIS
                     Host_ plugin parses       store-and-forward outbox
                                               ├─ HL7 ORU over TCP (MLLP)
                                               ├─ HL7 ORU files to a folder
                                               ├─ DTP_ plugin (DB / API)
                                               └─ ASTM E1394 to LIS
```

Decisions already made:

| Question | Decision |
|---|---|
| Shape | A **mode inside Interface.exe**, not a new app. |
| Outputs | **All four**: HL7 MLLP, HL7 files, DTP plugin, ASTM. Each can be turned on per site. **HL7 over TCP comes first** (phase 1). |
| Visibility | Reporting and verification are hidden; an **Outbox monitor screen** is added (phase 2). |
| Outbox storage | **File spool** with pending / sent / failed folders. |
| Orders | **Results only.** The LIS → analyzer order path is out of scope. The existing HL7 ORM path is left as it is. |
| Delivery | **Store-and-forward.** Nothing is dropped. Every message is retried until the destination accepts it. |

## What changes for the user

When `Lab-I\Middleware\Enable = 1`:

- **Hidden:** Maintenance (master data), Reports, Reprint, Worklist, Requests,
  QC, result entry/verification (`HostSA.showResultEntry`) and Autoprint
  (`tmrAutoPrint`, `LabTempPrint`).
- **Forced on:** Autoprocess. Every result is parsed as soon as it arrives.
- **Kept:** the connection menu, ComSet, plugin settings, Search (read-only, as an
  audit view), the console, and licensing.
- **New: Outbox monitor screen** (File → Outbox, new form `frmOutbox`):
  - It has one tab per enabled destination, with counts for pending, sent and
    failed.
  - Its grid lists each message: time, InstrumentID, sample/accession number,
    attempts, last error and status.
  - **Resend** moves the selected `failed\` (or `sent\`) entries back to
    `pending\`. **View** shows the payload that was built.
  - It refreshes from the spool folders on a timer. It reads files only; it never
    touches the result tables.
  - A one-line Outbox status also shows in `frmReceiveData`. The console gets
    `OUTBOX` and `RESEND` commands as a fallback.

`HostForms_SA.dll` is still loaded, because `setHostForms` injects the DB
connection and runs the schema patch. Only its screens are hidden. No public
COM signature changes, so mixed-version sites keep working.

## Result flow in middleware mode

```
processData (frmReceiveData ~1180)
   │  raw message tMess
   ├─▶ SaveResult / LabUnprocessed     (unchanged: local audit copy)
   └─▶ Outbox.Enqueue tMess            (new: one entry per enabled destination)
            │
            ▼  tmrOutbox (new timer, every few seconds)
       for each pending entry:
         build payload (HL7 ORU / ASTM / DTP call)
         send → wait for ACK
           accepted → move to sent\
           rejected / timeout → retry with backoff; after N tries mark failed\
                                (it stays retryable from the console)
```

Today the DTP and HL7-file calls run inline in `processData`. They only run when
Autoprocess is on, and a failure is only shown on screen, so the result is lost.
In middleware mode those inline calls are replaced by `Outbox.Enqueue`.
**Outside middleware mode the current behaviour is unchanged.**

### Outbox storage

The outbox is a **file spool, not a database table**:
`<DataDir>\Outbox\<Destination>\{pending,sent,failed}\<yyyymmddhhnnss>_<seq>.msg`.

- It works the same for LOCAL, MSSQL, MySQL and Oracle installs, with no schema patch.
- It survives a database outage. Each file holds the raw analyzer message plus a
  small header (InstrumentID, attempt count, last error).
- `sent\` is pruned by the existing monthly auto-archive. `failed\` is never
  pruned automatically.
- The folder is in the data folder, which `.gitignore` already excludes. It holds
  patient data.

## Payload builders

### HL7 ORU^R01 (used by both MLLP and file output)

`HL7_ORU` exists, but it is not good enough for middleware:

1. It only handles `MessageFormat` STKS and ASTM. Every other plugin returns `""`.
2. It takes PID/ORC/OBR from a matching order in `tbLabRequest`. Results-only
   sites have no orders.
3. MSH-3 and MSH-5 are swapped: the "Receiving Application" setting is written to
   the sending-application field.
4. MSH-10 (control ID) only changes once per second, so it is not unique.
5. Segments are joined with CRLF, but HL7/MLLP requires CR.

Plan: add a new **`HL7_ORU_Generic(obx, obi)`**. It builds the message from the
plugin contract fields (`segment_OBX` / `segment_oBI`), so it works with **every**
Host_ plugin:

| HL7 field | Source |
|---|---|
| PID-3 patient ID / OBR-3 filler no. | ORI(3) sample number, ORI(6) accession |
| PID-5 name | ORI(2) |
| OBX-3 test code | OBX(3), mapped through a per-destination code map (see open items) |
| OBX-5 / OBX-6 | OBX(5) result, OBX(6) unit |
| OBX-7 / OBX-8 | reference range and H/L flag from `tbLabClassExams`, when present |
| OBX-14 / OBR-7 | OBX(14) result date |
| MSH-10 | unique: `yyyymmddhhnnss` + sequence |

The existing `HL7_ORU` stays untouched, because current sites use it.

### HL7 over TCP (MLLP)

- Add a new class `clsMLLPClient`. It creates its Winsock control at runtime with
  `Controls.Add "MSWinsock.Winsock"`, so no form designer block is edited. The
  Winsock control is already referenced: `mdiMain.wsBroadcast`.
- Framing: `<VT>` message `<FS><CR>`. Wait for `MSA|AA` or `MSA|CA` within
  `AckTimeout`. `AE`/`AR`/timeout counts as a failure and is retried.
- Settings (`Lab-I\Middleware`): `HL7TcpHost`, `HL7TcpPort`, `AckTimeout`
  (default 10 s), `KeepAlive`.

### HL7 files

These reuse `FWDIR` from `HL7 Settings`. The change is to write a temp name and
then rename it, so the LIS never picks up a half-written file, and to use the
generic builder.

### DTP plugin

The outbox calls `ObjDTP.ProcessDTP obx, obi, raw` and treats
`ObjDTP.ErrorNum <> 0` as a failure that is retried. This needs no contract
change. Existing DTP plugins (DTP_OPENMRS, DTP_SLMC, …) work as they are, but
they must tolerate being called again for the same result. That has to be
checked per plugin.

### ASTM E1394 to LIS

This is the largest piece. It needs a sender state machine (ENQ / frames with
checksum / ETB/ETX / EOT, ACK/NAK per frame) over TCP or a spare COM port.
Records are built as H / P / O / R / L from the contract fields. It is phase 3,
and only for a site that actually needs it, because the target LIS's ASTM
dialect decides the details.

## Phases

| Phase | Scope | Touches |
|---|---|---|
| **1** | Middleware switch + hidden menus + forced Autoprocess; outbox spool + retry timer; `HL7_ORU_Generic`; **HL7 over TCP (MLLP)** with ACK handling; HL7 files through the outbox; status line + console `OUTBOX` / `RESEND` | `modMain.bas`, `frmReceiveData.frm` (code only), `mdiMain.frm` (code only), new `modOutbox.bas`, new `clsMLLPClient.cls` |
| **2** | **Outbox monitor screen** (`frmOutbox`); DTP output through the outbox | new `frmOutbox.frm`, `mdiMain.frm` (one new menu item) |
| **3** | ASTM sender | new `clsASTMSender.cls` |
| **4** | Settings screen for the Middleware section (until then, registry/console only) | new `frmMiddlewareSettings.frm` |

Phase 2 adds a File → Outbox menu item. That is the only designer change to an
existing form (`mdiMain.frm`). It has to be added in the VB6 IDE on the build PC,
not by hand-editing text. Until then, `frmOutbox` can be opened from the console
(`OUTBOX SHOW`).

Each phase needs a VB6 build on the build PC. These cannot be compiled from this
machine.

## Open items for the customer/site

- **Test code mapping.** Which LIS code does each analyzer test go to? Proposed: a
  `Outbox\<Destination>\codemap.txt` (analyzer code → LIS code), or a column on
  `tbcolabexam`. This needs a decision before go-live at the first site.
- **Patient identity.** Most results-only LIS match on sample/accession number
  (ORI(3)/ORI(6)), not on name. Confirm per LIS.
- **Multiple analyzers.** All four ports share one plugin (`MsgDef`), as noted in
  SYSTEM_MAP risk #5. A site with several analyzer types needs one Interface.exe
  instance per analyzer, or that limitation fixed first.
- **QC results.** Forward them, or keep them local? The default is to keep them
  local.
