# Graph Report - Interface Main Application  (2026-11-10)

## Corpus Check
- 104 files · ~0 words
- Verdict: corpus is large enough that graph structure adds value.

## Summary
- 1888 nodes · 3523 edges · 97 communities (94 shown, 3 thin omitted)
- Extraction: 93% EXTRACTED · 7% INFERRED · 0% AMBIGUOUS · INFERRED: 234 edges (avg confidence: 0.85)
- Token cost: 0 input · 0 output

## Community Hubs (Navigation)
- SA: Result Entry
- Interface: Outbox
- RE: Edit Request
- RE: Wards Rooms
- SA: Normal Values
- ComSet: Com Settings
- Interface: Base64
- SA: JSON
- Interface: Receive Data
- Interface: Request
- SA: Lab Exam
- Interface: Main
- SA: Class Forms
- SA: Graph
- Interface: Settings
- SA: Lab Profile Exams
- Interface: OBISetiings
- Interface: Registration
- SA: Functions
- Interface: Main (2)
- SA: Report Formats
- Interface: Main (3)
- Interface: Search Result
- RE: Search Request
- SA: Result Entry (2)
- Interface: Search
- RE: Blowfish
- RE: Request Setting
- SA: Points
- Interface: Blowfish
- SA: PIDsearch
- SA: Result Entry (3)
- Interface: Work List
- RE: Selection
- SA: Lab Profiles
- SA: Datasets
- RE: Class Main
- Shared: Data Base Copy
- Interface: Main (4)
- RE: Class Forms
- RE: Exam Selection
- SA: Dataset
- Interface: MLLPClient
- Interface: DBsettings
- RE: Edit Request (2)
- SA: Input Box
- RE: Del Req
- SA: Logbook
- SA: Print Copies
- SA: Selection Lists
- SA: Rpt Export
- Interface: License Manager
- Interface: Check
- SA: Pathologist
- SA: Result Entry (4)
- SA: Search Code
- RE: Summary Request
- SA: Form Settings
- Interface: Search Result (2)
- SA: Rpt Viewer
- SA: Graph (2)
- Interface: Input Box
- Interface: Process
- Interface: Main (5)
- Interface: Settings (2)
- SA: Medtech
- SA: Result Entry (5)
- Interface: Registration (2)
- RE: Req Rates
- SA: Physician
- SA: Result Entry (6)
- Interface: Db Man
- Interface: Archive
- Interface: Data Base Copy
- Interface: Result Viewer
- RE: Sample In
- SA: Hospital
- SA: Rem Template
- SA: Lab Profile Exams (2)
- Interface: Activation
- Interface: Main (6)
- SA: Worklist Report
- SA: Specimen
- Interface: Class Forms
- Interface: Search (2)
- Interface: Search Result (3)
- SA: Add Profiles
- Interface: Control Panel
- Interface: Settings (3)
- Interface: Statistics
- Interface: Backup List
- Interface: Search Result (4)
- Interface: Search Result (5)
- Interface: Search Result (6)
- Interface: Settings (4)
- SA: Result Entry (7)
- Interface: Search Result (7)

## God Nodes (most connected - your core abstractions)
1. `frmResultEntry (frm)` - 188 edges
2. `frmSearchResult (frm)` - 75 edges
3. `frmEditRequest (frm)` - 69 edges
4. `mdiMain (frm)` - 65 edges
5. `modMain (bas)` - 64 edges
6. `frmSettings (frm)` - 61 edges
7. `frmReceiveData (frm)` - 54 edges
8. `frmNormalValues (frm)` - 51 edges
9. `modFunctions.highLight()` - 49 edges
10. `ClassForms (cls)` - 45 edges

## Surprising Connections (you probably didn't know these)
- `ClassForms.ShowResultViewer()` --references--> `frmResultViewer (frm)`  [EXTRACTED]
  New Lab-I System v2/ClassForms.cls → New Lab-I System v2/frmResultViewer.frm
- `ClassForms.showInputBox()` --references--> `frmInputBox (frm)`  [EXTRACTED]
  New Lab-I System v2/ClassForms.cls → New Lab-I System v2/frmInputBox.frm
- `modOutbox.CompleteSpool()` --calls--> `clsMLLPClient.Accepted()`  [EXTRACTED]
  New Lab-I System v2/modOutbox.bas → New Lab-I System v2/clsMLLPClient.cls
- `ClassForms.SearchRequestFormLeft()` --references--> `frmSearchRequest (frm)`  [EXTRACTED]
  HostFomrs_RE/ClassForms.cls → HostFomrs_RE/frmSearchRequest.frm
- `ClassForms.showSearchForm()` --references--> `frmSearchRequest (frm)`  [EXTRACTED]
  HostFomrs_RE/ClassForms.cls → HostFomrs_RE/frmSearchRequest.frm

## Import Cycles
- None detected.

## Communities (97 total, 3 thin omitted)

### Community 0 - "SA: Result Entry"
Cohesion: 0.02
Nodes (3): frmResultEntry (frm), frmResultEntry.cboResult_Change(), frmResultEntry.cboResult_Click()

### Community 1 - "Interface: Outbox"
Cohesion: 0.07
Nodes (71): frmMiddleware (frm), frmMiddleware.cboDest_Click(), frmMiddleware.cboFolder_Click(), frmMiddleware.cmdClose_Click(), frmMiddleware.cmdRefresh_Click(), frmMiddleware.cmdResend_Click(), frmMiddleware.cmdSave_Click(), frmMiddleware.cmdSwitch_Click() (+63 more)

### Community 2 - "RE: Edit Request"
Cohesion: 0.06
Nodes (37): ClassForms.printRequestForm(), frmEditRequest (frm), frmEditRequest.blnMandatoryClear(), frmEditRequest.CheckWithSTAT(), frmEditRequest.chkStat_Click(), frmEditRequest.cmbmsfgResult_KeyDown(), frmEditRequest.cmbmsfgResult_KeyPress(), frmEditRequest.DelExam() (+29 more)

### Community 3 - "RE: Wards Rooms"
Cohesion: 0.06
Nodes (26): frmWardsRooms (frm), frmWardsRooms.checkDefRate(), frmWardsRooms.cmdDeptClose_Click(), frmWardsRooms.cmdRoomClose_Click(), frmWardsRooms.cmdUpdateDept_Click(), frmWardsRooms.cmdUpdateRoom_Click(), frmWardsRooms.cmdUpdateWard_Click(), frmWardsRooms.cmdWardClose_Click() (+18 more)

### Community 4 - "SA: Normal Values"
Cohesion: 0.05
Nodes (26): frmNormalValues (frm), frmNormalValues.clearItems(), frmNormalValues.cmdAddSave_Click(), frmNormalValues.cmdCancel_Click(), frmNormalValues.cmdCopyFrom_Click(), frmNormalValues.cmdDeleteItems_Click(), frmNormalValues.displayExams(), frmNormalValues.msfgResult_DblClick() (+18 more)

### Community 5 - "ComSet: Com Settings"
Cohesion: 0.06
Nodes (26): ComSettings (cls), ComSettings.showSettings(), frmSelectPlugins (frm), frmSettings (frm), frmSettings2 (frm), frmSettings2.AssignValues(), frmSettings2.cboControl_Change(), frmSettings2.cboControl_Click() (+18 more)

### Community 6 - "Interface: Base64"
Cohesion: 0.05
Nodes (30): frmAbout (frm), frmAbout.cmdSysInfo_Click(), frmAbout.GetKeyValue(), frmAbout.StartSysInfo(), frmBarCode (frm), frmBarCode.cmdPrint_Click(), frmBarCode.cmdPrint_GotFocus(), frmCharmap (frm) (+22 more)

### Community 7 - "SA: JSON"
Cohesion: 0.10
Nodes (29): cJSONScript (cls), cJSONScript.Eval(), cJSONScript.ExecCommand(), cStringBuilder (cls), cStringBuilder.Append(), cStringBuilder.AppendByVal(), cStringBuilder.AppendNL(), cStringBuilder.Insert() (+21 more)

### Community 8 - "Interface: Receive Data"
Cohesion: 0.11
Nodes (32): frmReceiveData (frm), frmReceiveData.ArchiveUnprocessed(), frmReceiveData.checkComStatus(), frmReceiveData.Command1_Click(), frmReceiveData.Command2_Click(), frmReceiveData.Command3_Click(), frmReceiveData.Command4_Click(), frmReceiveData.consoleCmd() (+24 more)

### Community 9 - "Interface: Request"
Cohesion: 0.06
Nodes (4): frmRequest (frm), frmRequest.cmdAdd_Click(), frmRequest.Form_Load(), mdiMain.mnuRequest_Click()

### Community 10 - "SA: Lab Exam"
Cohesion: 0.09
Nodes (23): ClassForms.showLabExams(), frmLabExam (frm), frmLabExam.cboResultType_Click(), frmLabExam.ClearValues(), frmLabExam.cmdAddSave_Click(), frmLabExam.cmdCancel_Click(), frmLabExam.cmdDown_Click(), frmLabExam.cmdGenerateSort_Click() (+15 more)

### Community 11 - "Interface: Main"
Cohesion: 0.09
Nodes (18): mdiMain (frm), mdiMain.mnuBackUp_Click(), mdiMain.mnuComSet_Click(), mdiMain.mnuControlPanel_Click(), mdiMain.mnuDisconnect_Click(), mdiMain.mnuDoctors_Click(), mdiMain.mnuExit_Click(), mdiMain.mnuLabExam_Click() (+10 more)

### Community 12 - "SA: Class Forms"
Cohesion: 0.08
Nodes (12): ClassForms (cls), ClassForms.GetSearchRecordSet(), ClassForms.SearchCode(), ClassForms.showClassifications(), ClassForms.showNormalValues(), ClassForms.showPIDsearch(), ClassForms.showResultEntry(), ClassForms.showSearchCode() (+4 more)

### Community 13 - "SA: Graph"
Cohesion: 0.13
Nodes (29): Graph (ctl), Graph.AddDefaultPoint(), Graph.AddDefaultPoints(), Graph.AxisColor(), Graph.BackColor(), Graph.BarWidth(), Graph.DrawBar(), Graph.DrawGraph() (+21 more)

### Community 14 - "Interface: Settings"
Cohesion: 0.07
Nodes (7): frmSettings (frm), frmSettings.cmdSet_Click(), frmSettings.cmdSetPlugins_Click(), frmSettings.List1_Click(), frmSettings.LoadFileToTB(), frmSettings.txtDetails_DblClick(), mdiMain.mnuProgSettings_Click()

### Community 15 - "SA: Lab Profile Exams"
Cohesion: 0.12
Nodes (20): ClassForms.showProfileExams(), frmLabProfileExams (frm), frmLabProfileExams.cmdAddExam_Click(), frmLabProfileExams.cmdBack_Click(), frmLabProfileExams.cmdDown_Click(), frmLabProfileExams.cmdForward_Click(), frmLabProfileExams.cmdSet_Click(), frmLabProfileExams.cmdUp_Click() (+12 more)

### Community 16 - "Interface: OBISetiings"
Cohesion: 0.11
Nodes (19): frmOBISetiings (frm), frmOBISetiings.cboServerType_Change(), frmOBISetiings.cboServerType_Click(), frmOBISetiings.cmdConnect_Click(), frmOBISetiings.cmdTestApi_Click(), frmOBISetiings.FillCboList(), frmOBISetiings.findCboIndex(), frmOBISetiings.findcboType() (+11 more)

### Community 17 - "Interface: Registration"
Cohesion: 0.12
Nodes (17): frmRegistration (frm), frmRegistration.cmdNewCode_Click(), frmRegistration.EnControls(), frmRegistration.Form_Load(), frmRegistration.ISHEX(), frmRegistration.OpenRegistrationServer(), frmRegistration.ProcessResponse(), frmRegistration.RetrieveRegistration() (+9 more)

### Community 18 - "SA: Functions"
Cohesion: 0.15
Nodes (21): ClassForms.PrintBarCode(), ClassForms.ResultAutoPrint(), ClassForms.ResultPrint(), ClassForms.showResultAutoPrint(), ClassForms.showResultPrintPreview(), frmClassifications.cmdPrint_Click(), frmLabProfiles.cmdPrint_Click(), frmLogbook.cmdPrint_Click() (+13 more)

### Community 19 - "Interface: Main (2)"
Cohesion: 0.10
Nodes (17): frmDailyBackup.Form_Activate(), frmSettings.cmdSave_MouseUp(), modMain (bas), modMain.CLargeInt(), modMain.DailyBackup(), modMain.EnumRegistryKeys(), modMain.EnumRegistryValuesEx(), modMain.FilenameToDate() (+9 more)

### Community 20 - "SA: Report Formats"
Cohesion: 0.15
Nodes (17): ClassForms.showReportFormats(), frmReportFormats (frm), frmReportFormats.assignValues(), frmReportFormats.ClearValues(), frmReportFormats.cmdBrowse_Click(), frmReportFormats.cmdCancel_Click(), frmReportFormats.cmdSave_Click(), frmReportFormats.cmdUpdate_Click() (+9 more)

### Community 21 - "Interface: Main (3)"
Cohesion: 0.14
Nodes (22): frmLabProfiles.cmdSave_Click(), frmPlugInTest (frm), frmPlugInTest.cboFormatting_Click(), frmPlugInTest.cmdNextRecord_Click(), frmPlugInTest.cmdPreviousRecord_Click(), frmPlugInTest.fileHosts_Click(), frmPlugInTest.fillcboFormatting(), frmPlugInTest.Form_Load() (+14 more)

### Community 22 - "Interface: Search Result"
Cohesion: 0.12
Nodes (16): frmSearchResult (frm), frmSearchResult.cmdCancel_Click(), frmSearchResult.HighlightOver(), frmSearchResult.Image5_Click(), frmSearchResult.Image6_Click(), frmSearchResult.lblArchive_Click(), frmSearchResult.lblArchive_MouseMove(), frmSearchResult.lblCancel_Click() (+8 more)

### Community 23 - "RE: Search Request"
Cohesion: 0.17
Nodes (17): ClassForms.SearchRequestFormTop(), frmSearchRequest (frm), frmSearchRequest.cmdCancelSearchHIS_Click(), frmSearchRequest.cmdNew_Click(), frmSearchRequest.cmdSearch_Click(), frmSearchRequest.delPrevExtDB(), frmSearchRequest.Form_Load(), frmSearchRequest.initGrid() (+9 more)

### Community 24 - "SA: Result Entry (2)"
Cohesion: 0.10
Nodes (23): frmResultEntry.ConvertToNumeric(), frmResultEntry.DisplayResultForm(), frmResultEntry.FillCBOResult(), frmResultEntry.findCboIndex(), frmResultEntry.FormulaValue(), frmResultEntry.getResultValue(), frmResultEntry.hideBlankInstrumentResults(), frmResultEntry.lblClassName_Change() (+15 more)

### Community 25 - "Interface: Search"
Cohesion: 0.11
Nodes (10): frmSearch (frm), frmSearch.cmdClear_Click(), frmSearch.Form_Load(), frmSearch.Form_Resize(), frmSearch.Form_Unload(), frmWall (frm), frmWall.changeSize(), frmWall.Form_Load() (+2 more)

### Community 26 - "RE: Blowfish"
Cohesion: 0.22
Nodes (20): clsBlowfish (cls), clsBlowfish.Append(), clsBlowfish.DecryptBlock(), clsBlowfish.DecryptByte(), clsBlowfish.DecryptFile(), clsBlowfish.DecryptString(), clsBlowfish.DeHex(), clsBlowfish.EncryptBlock() (+12 more)

### Community 27 - "RE: Request Setting"
Cohesion: 0.12
Nodes (12): frmRequestSetting (frm), frmRequestSetting.chkEditSort_Click(), frmRequestSetting.cmdDeptWardRoom_Click(), frmRequestSetting.cmdSave_Click(), frmRequestSetting.msfgRequestField_Click(), frmRequestSetting.msfgRequestField_LeaveCell(), frmRequestSetting.msfgRequestField_Scroll(), frmRequestSetting.txtLabel_KeyDown() (+4 more)

### Community 28 - "SA: Points"
Cohesion: 0.14
Nodes (17): modFunctions.PtrObj(), Dataset.Datasets(), Dataset.Points(), Point (cls), Point.Points(), Point.State(), Point.StateChanged(), Point.SuperState() (+9 more)

### Community 29 - "Interface: Blowfish"
Cohesion: 0.22
Nodes (20): clsBlowfish (cls), clsBlowfish.Append(), clsBlowfish.DecryptBlock(), clsBlowfish.DecryptByte(), clsBlowfish.DecryptFile(), clsBlowfish.DecryptString(), clsBlowfish.DeHex(), clsBlowfish.EncryptBlock() (+12 more)

### Community 30 - "SA: PIDsearch"
Cohesion: 0.15
Nodes (12): frmPIDsearch (frm), frmPIDsearch.assignValues(), frmPIDsearch.ClearValues(), frmPIDsearch.cmdSelect_Click(), frmPIDsearch.Form_Load(), frmPIDsearch.GenerateTestPacket(), frmPIDsearch.GetError(), frmPIDsearch.GetHTTPStatusCodeDesc() (+4 more)

### Community 31 - "SA: Result Entry (3)"
Cohesion: 0.18
Nodes (21): frmResultEntry.txtAge_GotFocus(), frmResultEntry.txtBirthDate_GotFocus(), frmResultEntry.txtCivilStatus_GotFocus(), frmResultEntry.txtClassID_GotFocus(), frmResultEntry.txtEntry_GotFocus(), frmResultEntry.txtFirstName_GotFocus(), frmResultEntry.txtLastName_GotFocus(), frmResultEntry.txtMedtech_GotFocus() (+13 more)

### Community 32 - "Interface: Work List"
Cohesion: 0.18
Nodes (18): frmWorkList (frm), frmWorkList.cboFilter_Click(), frmWorkList.cmdPos_Click(), frmWorkList.cmdRefresh_Click(), frmWorkList.cmdRemoveFromWorklist_Click(), frmWorkList.cmdSend_Click(), frmWorkList.dtpDate_Click(), frmWorkList.fillCboFilter() (+10 more)

### Community 33 - "RE: Selection"
Cohesion: 0.13
Nodes (9): frmSelection (frm), frmSelection.cmdSelect_Click(), frmSelection.Form_Load(), frmSelection.msfgResult_DblClick(), frmSelection.msfgResult_KeyPress(), frmSelection.setGrid(), frmSelection.txtFilter_Change(), frmSelection.txtFilter_KeyDown() (+1 more)

### Community 34 - "SA: Lab Profiles"
Cohesion: 0.12
Nodes (12): ClassForms.showLabProfiles(), frmLabProfiles (frm), frmLabProfiles.cmdTemplate_Click(), frmLabProfiles.Form_Load(), frmLabProfiles.setFlexGrid(), frmTemplate (frm), frmTemplate.cmdSelect_Click(), frmTemplate.Form_Load() (+4 more)

### Community 35 - "SA: Datasets"
Cohesion: 0.13
Nodes (15): Datasets (cls), Datasets.Add(), Datasets.Class_Initialize(), Datasets.Clear(), Datasets.Load(), Datasets.Save(), Datasets.StateChanged(), Datasets.SuperState() (+7 more)

### Community 36 - "RE: Class Main"
Cohesion: 0.16
Nodes (18): ClassForms.getDispTestReqSum(), ClassForms.wConnection(), frmEditRequest.Form_Load(), frmEditRequest.ReqIsLocked(), frmEditRequest.ReqLock(), frmEditRequest.setmsfgResultDemo(), frmEditRequest.setmsfgResultReq(), frmEditRequest.sizeFix() (+10 more)

### Community 37 - "Shared: Data Base Copy"
Cohesion: 0.12
Nodes (6): frmDataBaseCopy.GenerateInsertResultID(), frmDataBaseCopy.StartCopy(), frmDbMan.cmdArchive_Click(), frmDbMan.cmdGenerateResultID_Click(), modMain.GenerateResultID(), modMain.LinkTable()

### Community 38 - "Interface: Main (4)"
Cohesion: 0.22
Nodes (18): frmDbMan.cmdChangePassword_Click(), frmDbMan.cmdCompactDB_MouseUp(), frmDbMan.cmdRestoreDB_Click(), frmSettings.cmdClose_Click(), mdiMain.mnuConnServ_Click(), mdiMain.mnuConSet_Click(), mdiMain.mnuReportFormats_Click(), modMain.BackupDatabase() (+10 more)

### Community 39 - "RE: Class Forms"
Cohesion: 0.12
Nodes (10): ClassForms (cls), ClassForms.CheckInSample(), ClassForms.Class_Initialize(), ClassForms.EditRequest(), ClassForms.printBarCode(), ClassForms.SearchRequestFormLeft(), ClassForms.showRequestSettings(), ClassForms.showSearchForm() (+2 more)

### Community 40 - "RE: Exam Selection"
Cohesion: 0.16
Nodes (11): frmExamSelection (frm), frmExamSelection.cmbSpecimen_Click(), frmExamSelection.cmdSelect_Click(), frmExamSelection.Form_Load(), frmExamSelection.msfgResult_DblClick(), frmExamSelection.msfgResult_KeyPress(), frmExamSelection.setGrid(), frmExamSelection.setListCmbSelection() (+3 more)

### Community 41 - "SA: Dataset"
Cohesion: 0.19
Nodes (16): Dataset (cls), Dataset.BarColor(), Dataset.CapColor(), Dataset.Class_Initialize(), Dataset.InitProperties(), Dataset.LineColor(), Dataset.mobjPoints_Changed(), Dataset.PointColor() (+8 more)

### Community 42 - "Interface: MLLPClient"
Cohesion: 0.19
Nodes (14): clsMLLPClient (cls), clsMLLPClient.Accepted(), clsMLLPClient.CheckTimeout(), clsMLLPClient.Class_Terminate(), clsMLLPClient.Finish(), clsMLLPClient.LastError(), clsMLLPClient.ParseAck(), clsMLLPClient.SendFrame() (+6 more)

### Community 43 - "Interface: DBsettings"
Cohesion: 0.13
Nodes (6): frmDBsettings (frm), frmDBsettings.cmdSave_Click(), frmDBsettings.fillcboServerType(), frmDBsettings.Form_Load(), frmDBsettings.setCBOServer(), frmDBsettings.txtUsername_GotFocus()

### Community 44 - "RE: Edit Request (2)"
Cohesion: 0.23
Nodes (16): frmEditRequest.AddExamDet(), frmEditRequest.blnExamOccured(), frmEditRequest.cmdExamMenu_Click(), frmEditRequest.RemoveFromTestToDel(), frmEditRequest.txtmsfgExam_GotFocus(), frmEditRequest.txtmsfgExam_KeyDown(), frmEditRequest.txtmsfgExam_KeyPress(), frmEditRequest.txtmsfgResult_GotFocus() (+8 more)

### Community 45 - "SA: Input Box"
Cohesion: 0.12
Nodes (3): ClassForms.showInputBox(), frmInputBox (frm), frmInputBox.txtValue_GotFocus()

### Community 46 - "RE: Del Req"
Cohesion: 0.21
Nodes (12): ClassForms.DelReq(), frmDelReq (frm), frmDelReq.AddExamDet(), frmDelReq.blnExamOccured(), frmDelReq.DelExamDet(), frmDelReq.msfgResult_Click(), frmDelReq.msfgResult_EnterCell(), frmDelReq.msfgResult_LeaveCell() (+4 more)

### Community 47 - "SA: Logbook"
Cohesion: 0.13
Nodes (11): ClassForms.showLogbook(), ClassForms.showOtherReports(), frmInputBox.Form_Load(), frmLabProfileExams.Form_Load(), frmLogbook (frm), frmLogbook.Form_Load(), frmNormalValues.Form_Load(), frmOtherReports (frm) (+3 more)

### Community 48 - "SA: Print Copies"
Cohesion: 0.13
Nodes (7): ClassForms.showPrintCopies(), frmPrintCopies (frm), frmPrintCopies.txtCopy1_Change(), frmPrintCopies.txtCopy1_GotFocus(), frmPrintCopies.txtCopy2_Change(), frmPrintCopies.txtCopy2_GotFocus(), frmPrintCopies.txtCopy2Tag_GotFocus()

### Community 49 - "SA: Selection Lists"
Cohesion: 0.23
Nodes (10): ClassForms.showSelectionLists(), frmSelectionLists (frm), frmSelectionLists.cmdAdd_Click(), frmSelectionLists.cmdAddCategories_Click(), frmSelectionLists.cmdDelete_Click(), frmSelectionLists.cmdDeleteCat_Click(), frmSelectionLists.filllstType(), frmSelectionLists.findValues() (+2 more)

### Community 50 - "SA: Rpt Export"
Cohesion: 0.20
Nodes (12): frmRptExport (frm), frmRptExport.BrowseForFolder(), frmRptExport.cboFileType_Change(), frmRptExport.cboFileType_Click(), frmRptExport.cmdBrowse_Click(), frmRptExport.fillcboExport(), frmRptExport.findCboIndex(), frmRptExport.Form_Load() (+4 more)

### Community 51 - "Interface: License Manager"
Cohesion: 0.21
Nodes (12): frmLicenseManager (frm), frmLicenseManager.cmdGetLicense_Click(), frmLicenseManager.cmdRegister_Click(), frmLicenseManager.DeletePayload(), frmLicenseManager.GetLicences(), frmLicenseManager.LoadDetails(), frmLicenseManager.OpenRegistrationServer(), frmLicenseManager.ProcessResponse() (+4 more)

### Community 52 - "Interface: Check"
Cohesion: 0.25
Nodes (15): modCheck (bas), modCheck.ASTMChecksum(), modCheck.bin2dec(), modCheck.CheckSum8(), modCheck.CRC16_CCITT(), modCheck.CRC16_CCITT_M(), modCheck.dec2bin(), modCheck.isValidCS() (+7 more)

### Community 53 - "SA: Pathologist"
Cohesion: 0.22
Nodes (9): ClassForms.showPathologist(), frmPathologist (frm), frmPathologist.cmdAdd_Click(), frmPathologist.cmdDelete_Click(), frmPathologist.Form_Load(), frmPathologist.msfgResult_KeyDown(), frmPathologist.setFlexGrid(), frmResultEntry.Label1_Click() (+1 more)

### Community 54 - "SA: Result Entry (4)"
Cohesion: 0.16
Nodes (14): frmResultEntry.cmdSave_Click(), frmResultEntry.fillLists(), frmResultEntry.GetLicenses(), frmResultEntry.getTotalDiff(), frmResultEntry.hasDiff(), frmResultEntry.lblDiff_Change(), frmResultEntry.replaceStr(), frmResultEntry.returnDate() (+6 more)

### Community 55 - "SA: Search Code"
Cohesion: 0.16
Nodes (6): frmSearchCode (frm), frmSearchCode.cmdSelect_Click(), frmSearchCode.Form_Load(), frmSearchCode.msfgResult_DblClick(), frmSearchCode.setFlexGrid(), frmSearchCode.txtFilter_GotFocus()

### Community 56 - "RE: Summary Request"
Cohesion: 0.17
Nodes (7): ClassForms.showReqSummary(), frmRequest (frm), frmSummaryRequest (frm), frmSummaryRequest.Form_Load(), frmSummaryRequest.setMsfgSummary(), modClassMain.showSummaryRequest(), HostForms_RE (HostForms_RE.dll, OleDll)

### Community 57 - "SA: Form Settings"
Cohesion: 0.18
Nodes (5): ClassForms.showFormSettings(), frmFormSettings (frm), frmFormSettings.fillCboSelection(), frmFormSettings.findCboIndex(), frmFormSettings.Form_Load()

### Community 58 - "Interface: Search Result (2)"
Cohesion: 0.21
Nodes (11): frmResultEntry.getUpdatedValues(), frmSearchResult.ArchiveResults(), frmSearchResult.cmdDelete_Click(), frmSearchResult.DeleteResults(), frmSearchResult.Image4_Click(), frmSearchResult.Image7_Click(), frmSearchResult.lblDelete_Click(), frmSearchResult.lblUpload_Click() (+3 more)

### Community 59 - "SA: Rpt Viewer"
Cohesion: 0.15
Nodes (4): frmRptViewer (frm), frmRptViewer.CRViewer1_ExportButtonClicked(), frmRptViewer.CRViewer1_PrintButtonClicked(), modFunctions.ShowExportReport()

### Community 60 - "SA: Graph (2)"
Cohesion: 0.15
Nodes (13): Graph.Appearance(), Graph.BorderStyle(), Graph.ControlState(), Graph.DrawControl(), Graph.FileState(), Graph.GraphState(), Graph.InitProperties(), Graph.SaveSettings() (+5 more)

### Community 62 - "Interface: Process"
Cohesion: 0.24
Nodes (11): frmProcess (frm), frmProcess.Form_Load(), frmProcess.processData(), frmProcess.SaveResult(), frmProcess.Timer1_Timer(), frmReceiveData.processData(), frmReceiveData.QueueForMiddleware(), frmReceiveData.SaveResult() (+3 more)

### Community 63 - "Interface: Main (5)"
Cohesion: 0.19
Nodes (12): frmReceiveData.Form_Unload(), mdiMain.ApplyMiddlewareMenus(), mdiMain.MDIForm_Load(), mdiMain.MDIForm_Resize(), mdiMain.mnuConnect_Click(), mdiMain.mnuSendID_Click(), mdiMain.mnuSendReceive_Click(), mdiMain.tbrMain_ButtonMenuClick() (+4 more)

### Community 64 - "Interface: Settings (2)"
Cohesion: 0.15
Nodes (13): frmSettings.File1_Click(), frmSettings.File2_Click(), frmSettings.fillcboExport(), frmSettings.fillcboScheme(), frmSettings.findcboType(), frmSettings.Form_Load(), frmSettings.initCdgExportLocation(), frmSettings.ListSubDirs() (+5 more)

### Community 65 - "SA: Medtech"
Cohesion: 0.27
Nodes (7): ClassForms.showMedTech(), frmMedtech (frm), frmMedtech.cmdAdd_Click(), frmMedtech.cmdDelete_Click(), frmMedtech.Form_Load(), frmMedtech.msfgResult_KeyDown(), frmMedtech.setFlexGrid()

### Community 66 - "SA: Result Entry (5)"
Cohesion: 0.21
Nodes (12): frmResultEntry.ClearValues(), frmResultEntry.cmdNext_Click(), frmResultEntry.cmdPrevious_Click(), frmResultEntry.Form_Load(), frmResultEntry.getIResultDate(), frmResultEntry.InitForm(), frmResultEntry.msfgResult_EnterCell(), frmResultEntry.OpenPrevious() (+4 more)

### Community 67 - "Interface: Registration (2)"
Cohesion: 0.17
Nodes (12): frmDBsettings.txtDatabase_GotFocus(), frmDBsettings.txtPassword_GotFocus(), frmDBsettings.txtSchema_GotFocus(), frmDBsettings.txtServerName_GotFocus(), frmRegistration.txtCode1_GotFocus(), frmRegistration.txtCode2_GotFocus(), frmRegistration.txtCode3_GotFocus(), frmRegistration.txtCode4_GotFocus() (+4 more)

### Community 68 - "RE: Req Rates"
Cohesion: 0.24
Nodes (7): ClassForms.showReqRates(), frmReqRates (frm), frmReqRates.cmdSaveDiscount_Click(), frmReqRates.Form_Load(), frmReqRates.setmsfgRates(), frmReqRates.setmsfgReqDiscounts(), frmReqRates.txtDiscountName_GotFocus()

### Community 69 - "SA: Physician"
Cohesion: 0.31
Nodes (7): ClassForms.showPhysician(), frmPhysician (frm), frmPhysician.cmdAdd_Click(), frmPhysician.cmdDelete_Click(), frmPhysician.Form_Load(), frmPhysician.msfgResult_KeyDown(), frmPhysician.setFlexGrid()

### Community 70 - "SA: Result Entry (6)"
Cohesion: 0.20
Nodes (11): frmResultEntry.AssignPIDValues(), frmResultEntry.cmdPIDSearch_Click(), frmResultEntry.Form_Activate(), frmResultEntry.GetIdFromName(), frmResultEntry.getNVclass(), frmResultEntry.setValues(), frmResultEntry.txtBirthDate_Change(), frmResultEntry.txtClassID_KeyDown() (+3 more)

### Community 71 - "Interface: Db Man"
Cohesion: 0.20
Nodes (4): frmDbMan (frm), frmDbMan.cmdCopy_Click(), frmDbMan.cmdFixLink_Click(), frmDbMan.showStatus()

### Community 72 - "Interface: Archive"
Cohesion: 0.20
Nodes (8): ClassForms.showArchive(), frmActivation.txtSerial_Change(), frmArchive (frm), frmArchive.Form_Load(), frmInputBox.Form_Load(), frmReceiveData.Form_Activate(), frmSearch.Form_Activate(), modMain.setFormColors()

### Community 73 - "Interface: Data Base Copy"
Cohesion: 0.22
Nodes (6): frmDataBaseCopy (frm), frmDataBaseCopy.cmdStart_Click(), frmDataBaseCopy.Dir1_Change(), frmDataBaseCopy.Dir1_Click(), frmDataBaseCopy.findCol(), frmDataBaseCopy.Form_Load()

### Community 74 - "Interface: Result Viewer"
Cohesion: 0.24
Nodes (6): frmResultViewer (frm), frmResultViewer.cmdSave_Click(), frmResultViewer.Form_Load(), frmResultViewer.setmsfgResult(), frmResultViewer.txtSampleNum_GotFocus(), frmResultViewer.UpdateData()

### Community 75 - "RE: Sample In"
Cohesion: 0.25
Nodes (3): frmSampleIn (frm), frmSampleIn.Form_Load(), frmSampleIn.setmsfgSamples()

### Community 76 - "SA: Hospital"
Cohesion: 0.28
Nodes (5): ClassForms.showHospitalInfo(), frmHospital (frm), frmHospital.cmdSave_Click(), frmHospital.Form_Load(), frmHospital.replaceStr()

### Community 77 - "SA: Rem Template"
Cohesion: 0.36
Nodes (6): ClassForms.showRemarksTemplate(), frmRemTemplate (frm), frmRemTemplate.cmdAdd_Click(), frmRemTemplate.Form_Load(), frmRemTemplate.msfgResult_KeyDown(), frmRemTemplate.setFlexGrid()

### Community 78 - "SA: Lab Profile Exams (2)"
Cohesion: 0.22
Nodes (6): ClassForms.wConnection(), frmLabProfileExams.cmdCopy_Click(), frmLabProfileExams.cmdSave_Click(), frmResultEntry.setAdditionalProfiles(), frmDataBaseCopy.cmdCheck_Click(), frmDbMan.Form_Load()

### Community 79 - "Interface: Activation"
Cohesion: 0.25
Nodes (4): frmActivation (frm), frmActivation.cmdEnter_Click(), frmActivation.saveData(), frmActivation.txtSerial_GotFocus()

### Community 80 - "Interface: Main (6)"
Cohesion: 0.31
Nodes (8): frmLicenseManager.Form_Load(), frmRegistration.cmdRegister_Click(), mdiMain.mnuactivation_Click(), modMain.GetCpuID(), modMain.GetSAVersion(), modMain.GetSerialNumber(), modMain.IsRegistered(), modMain.SystemSerialNumber()

### Community 81 - "SA: Worklist Report"
Cohesion: 0.25
Nodes (3): ClassForms.showWorklistReport(), frmWorklistReport (frm), frmWorklistReport.Form_Load()

### Community 82 - "SA: Specimen"
Cohesion: 0.43
Nodes (5): frmSpecimen (frm), frmSpecimen.cmdAdd_Click(), frmSpecimen.Form_Load(), frmSpecimen.msfgResult_KeyDown(), frmSpecimen.setFlexGrid()

### Community 83 - "Interface: Class Forms"
Cohesion: 0.29
Nodes (7): ClassForms (cls), ClassForms.showInputBox(), ClassForms.ShowResultViewer(), frmDbMan.cmdClearBuffer_Click(), frmSettings.cmdTestPlugin_Click(), mdiMain.mnuClearBuffer_Click(), mdiMain.mnuPlugTest_Click()

### Community 84 - "Interface: Search (2)"
Cohesion: 0.25
Nodes (7): frmSearch.cboStatus_Click(), frmSearch.cmdSearch_Click(), frmSearch.getSearchString(), frmSearchResult.msfgResult_Click(), frmSearchResult.msfgResult_MouseDown(), frmSearchResult.wTrans(), mdiMain.mnuStat_Click()

### Community 85 - "Interface: Search Result (3)"
Cohesion: 0.25
Nodes (8): frmSearchResult.cmdPreview_Click(), frmSearchResult.Form_Load(), frmSearchResult.Form_Resize(), frmSearchResult.Image1_Click(), frmSearchResult.lblPreview_Click(), frmSearchResult.msfgResult_Scroll(), frmSearchResult.OpenLocalConnection(), frmSearchResult.setIcons()

### Community 86 - "SA: Add Profiles"
Cohesion: 0.33
Nodes (4): ClassForms.showAddProfile(), frmAddProfiles (frm), frmAddProfiles.Form_Load(), frmAddProfiles.setListView()

### Community 87 - "Interface: Control Panel"
Cohesion: 0.33
Nodes (6): frmControlPanel (frm), frmControlPanel.Form_Load(), frmControlPanel.lvwItems_DblClick(), frmControlPanel.lvwItems_KeyDown(), frmSettings.cmdSave_Click(), frmSettings.saveSettings()

### Community 88 - "Interface: Settings (3)"
Cohesion: 0.29
Nodes (7): frmSettings.Check1_Click(), frmSettings.chkRegistry_Click(), frmSettings.chkTaskM_Click(), frmSettings.chkUsbStor_Click(), frmSettings.cmdRegistry_Click(), frmSettings.cmdTaskMAn_Click(), frmSettings.SaveValue()

### Community 89 - "Interface: Statistics"
Cohesion: 0.43
Nodes (5): frmStatistics (frm), frmStatistics.countData(), frmStatistics.Form_Load(), frmStatistics.getAverage(), frmStatistics.setmsfgResult()

### Community 90 - "Interface: Backup List"
Cohesion: 0.40
Nodes (4): frmBackupList (frm), frmBackupList.Form_Load(), frmBackupList.ListSubDirs(), modMain.showBackupList()

### Community 91 - "Interface: Search Result (4)"
Cohesion: 0.40
Nodes (5): frmSearchResult.CanMerge(), frmSearchResult.GetResultID(), frmSearchResult.Image8_Click(), frmSearchResult.lblMerge_Click(), frmSearchResult.lblMerge_MouseMove()

### Community 92 - "Interface: Search Result (5)"
Cohesion: 0.40
Nodes (5): frmSearchResult.cmdOpen_Click(), frmSearchResult.Image2_Click(), frmSearchResult.lblEdit_Click(), frmSearchResult.msfgResult_DblClick(), frmSearchResult.setSelectedItems()

### Community 93 - "Interface: Search Result (6)"
Cohesion: 0.40
Nodes (5): frmSearchResult.cmdPrint_Click(), frmSearchResult.Image3_Click(), frmSearchResult.lblPrint_Click(), frmSearchResult.msfgResult_SelChange(), frmSearchResult.PrintResults()

### Community 94 - "Interface: Settings (4)"
Cohesion: 0.50
Nodes (4): frmSettings.chkPlugin_Click(), frmSettings.clearDetails(), frmSettings.fileList_Click(), frmSettings.setDetails()

## Knowledge Gaps
- **4 isolated node(s):** `frmSelectInstrument (frm)`, `frmRequest (frm)`, `modConstants (bas)`, `modTypes (bas)`
  These have ≤1 connection - possible missing edges. (Counts symbols only; 637 node(s) total have ≤1 connection when file, concept and rationale nodes are included.)
- **3 thin communities (<3 nodes) omitted from report** — run `graphify query` to explore isolated nodes.

## Suggested Questions
_Questions this graph is uniquely positioned to answer:_

- **Why does `mdiMain (frm)` connect `Interface: Main` to `Interface: Work List`, `Interface: Base64`, `Interface: Main (4)`, `Interface: Request`, `Interface: Settings`, `Interface: Main (6)`, `Interface: Class Forms`, `Interface: Search (2)`, `Interface: Process`, `Interface: Main (5)`?**
  _High betweenness centrality (0.020) - this node is a cross-community bridge._
- **Why does `mdiMain.MDIForm_Load()` connect `Interface: Main (5)` to `SA: Lab Profile Exams (2)`, `Interface: Class Forms`, `Interface: Main (4)`?**
  _High betweenness centrality (0.011) - this node is a cross-community bridge._
- **Why does `modMain.setHostForms()` connect `Interface: Main (4)` to `Interface: Main (6)`?**
  _High betweenness centrality (0.011) - this node is a cross-community bridge._
- **What connects `frmSelectInstrument (frm)`, `frmRequest (frm)`, `modConstants (bas)` to the rest of the system?**
  _4 weakly-connected nodes found - possible documentation gaps or missing edges._
- **Should `SA: Result Entry` be split into smaller, more focused modules?**
  _Cohesion score 0.02 - nodes in this community are weakly interconnected._
- **Should `Interface: Outbox` be split into smaller, more focused modules?**
  _Cohesion score 0.07207207207207207 - nodes in this community are weakly interconnected._
- **Should `RE: Edit Request` be split into smaller, more focused modules?**
  _Cohesion score 0.06464646464646465 - nodes in this community are weakly interconnected._