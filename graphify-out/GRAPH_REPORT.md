# Graph Report - Interface Main Application  (2026-11-10)

## Corpus Check
- 93 files · ~0 words
- Verdict: corpus is large enough that graph structure adds value.

## Summary
- 1690 nodes · 3058 edges · 90 communities (86 shown, 4 thin omitted)
- Extraction: 92% EXTRACTED · 8% INFERRED · 0% AMBIGUOUS · INFERRED: 231 edges (avg confidence: 0.85)
- Token cost: 0 input · 0 output

## Community Hubs (Navigation)
- SA: Result Entry
- SA: Rpt Export
- SA: Lab Profile Exams
- RE: Edit Request
- RE: Wards Rooms
- SA: Normal Values
- ComSet: Com Settings
- SA: JSON
- Interface: Receive Data
- Interface: Request
- SA: Lab Exam
- Interface: Main
- SA: Form Settings
- Interface: OBISetiings
- Interface: Settings
- Interface: Search Result
- SA: Report Formats
- Interface: Db Man
- Interface: Registration
- Interface: Main (2)
- RE: Search Request
- Interface: Search
- RE: Blowfish
- RE: Request Setting
- SA: PIDsearch
- Interface: Blowfish
- SA: Result Entry (2)
- SA: Result Entry (3)
- RE: Selection
- RE: Class Main
- SA: Lab Profiles
- Interface: Main (3)
- Interface: Work List
- RE: Class Forms
- RE: Exam Selection
- SA: Worklist Report
- Interface: DBsettings
- RE: Edit Request (2)
- SA: Input Box
- SA: Result Entry (4)
- Interface: Plug In Test
- Shared: Logbook
- RE: Del Req
- SA: Class Forms
- SA: Print Copies
- Interface: License Manager
- Interface: Check
- SA: Rem Template
- SA: Pathologist
- SA: Search Code
- RE: Summary Request
- SA: Hospital
- Interface: Input Box
- Interface: Settings (2)
- SA: Medtech
- SA: Result Entry (5)
- Interface: Control Panel
- Interface: Registration (2)
- RE: Req Rates
- SA: Physician
- Interface: Search Result (2)
- Interface: Main (4)
- Interface: Main (5)
- Interface: Activation
- Interface: Warning
- Interface: Result Viewer
- Interface: Search Result (3)
- RE: Sample In
- SA: Classifications
- SA: Result Entry (6)
- Interface: Base64
- SA: Specimen
- Interface: Class Forms
- Interface: Main (6)
- SA: Add Profiles
- Interface: Settings (3)
- Interface: Statistics
- Interface: Wall
- Interface: About
- Interface: Backup List
- Interface: Bar Code
- Interface: Charmap
- Interface: Archive
- Interface: HL7Settings
- Interface: Search Result (4)
- Interface: Search Result (5)
- Interface: Search Result (6)
- Interface: Create DNS
- Interface: Settings (4)
- Interface: Search Result (7)

## God Nodes (most connected - your core abstractions)
1. `frmResultEntry (frm)` - 188 edges
2. `frmSearchResult (frm)` - 75 edges
3. `frmEditRequest (frm)` - 69 edges
4. `modMain (bas)` - 63 edges
5. `frmSettings (frm)` - 61 edges
6. `mdiMain (frm)` - 59 edges
7. `frmReceiveData (frm)` - 52 edges
8. `frmNormalValues (frm)` - 51 edges
9. `modFunctions.highLight()` - 49 edges
10. `ClassForms (cls)` - 45 edges

## Surprising Connections (you probably didn't know these)
- `ClassForms.ShowResultViewer()` --references--> `frmResultViewer (frm)`  [EXTRACTED]
  New Lab-I System v2/ClassForms.cls → New Lab-I System v2/frmResultViewer.frm
- `ClassForms.showInputBox()` --references--> `frmInputBox (frm)`  [EXTRACTED]
  New Lab-I System v2/ClassForms.cls → New Lab-I System v2/frmInputBox.frm
- `ClassForms.SearchRequestFormLeft()` --references--> `frmSearchRequest (frm)`  [EXTRACTED]
  HostFomrs_RE/ClassForms.cls → HostFomrs_RE/frmSearchRequest.frm
- `ClassForms.showSearchForm()` --references--> `frmSearchRequest (frm)`  [EXTRACTED]
  HostFomrs_RE/ClassForms.cls → HostFomrs_RE/frmSearchRequest.frm
- `ClassForms.showRequestSettings()` --references--> `frmRequestSetting (frm)`  [EXTRACTED]
  HostFomrs_RE/ClassForms.cls → HostFomrs_RE/frmRequestSetting.frm

## Import Cycles
- None detected.

## Communities (90 total, 4 thin omitted)

### Community 0 - "SA: Result Entry"
Cohesion: 0.02
Nodes (5): frmResultEntry (frm), frmResultEntry.cboResult_Change(), frmResultEntry.cboResult_Click(), frmResultEntry.Form_Resize(), frmResultEntry.setPosDemo()

### Community 1 - "SA: Rpt Export"
Cohesion: 0.06
Nodes (39): ClassForms.PrintBarCode(), ClassForms.ResultAutoPrint(), ClassForms.ResultPrint(), ClassForms.showResultAutoPrint(), ClassForms.showResultPrintPreview(), frmClassifications.cmdPrint_Click(), frmLabProfiles.cmdPrint_Click(), frmLogbook.cmdPrint_Click() (+31 more)

### Community 2 - "SA: Lab Profile Exams"
Cohesion: 0.05
Nodes (32): ClassForms.showProfileExams(), frmLabProfileExams (frm), frmLabProfileExams.cmdAddExam_Click(), frmLabProfileExams.cmdBack_Click(), frmLabProfileExams.cmdCopy_Click(), frmLabProfileExams.cmdDown_Click(), frmLabProfileExams.cmdForward_Click(), frmLabProfileExams.cmdSave_Click() (+24 more)

### Community 3 - "RE: Edit Request"
Cohesion: 0.06
Nodes (37): ClassForms.printRequestForm(), frmEditRequest (frm), frmEditRequest.blnMandatoryClear(), frmEditRequest.CheckWithSTAT(), frmEditRequest.chkStat_Click(), frmEditRequest.cmbmsfgResult_KeyDown(), frmEditRequest.cmbmsfgResult_KeyPress(), frmEditRequest.DelExam() (+29 more)

### Community 4 - "RE: Wards Rooms"
Cohesion: 0.06
Nodes (26): frmWardsRooms (frm), frmWardsRooms.checkDefRate(), frmWardsRooms.cmdDeptClose_Click(), frmWardsRooms.cmdRoomClose_Click(), frmWardsRooms.cmdUpdateDept_Click(), frmWardsRooms.cmdUpdateRoom_Click(), frmWardsRooms.cmdUpdateWard_Click(), frmWardsRooms.cmdWardClose_Click() (+18 more)

### Community 5 - "SA: Normal Values"
Cohesion: 0.05
Nodes (26): frmNormalValues (frm), frmNormalValues.clearItems(), frmNormalValues.cmdAddSave_Click(), frmNormalValues.cmdCancel_Click(), frmNormalValues.cmdCopyFrom_Click(), frmNormalValues.cmdDeleteItems_Click(), frmNormalValues.displayExams(), frmNormalValues.msfgResult_DblClick() (+18 more)

### Community 6 - "ComSet: Com Settings"
Cohesion: 0.06
Nodes (26): ComSettings (cls), ComSettings.showSettings(), frmSelectPlugins (frm), frmSettings (frm), frmSettings2 (frm), frmSettings2.AssignValues(), frmSettings2.cboControl_Change(), frmSettings2.cboControl_Click() (+18 more)

### Community 7 - "SA: JSON"
Cohesion: 0.10
Nodes (29): cJSONScript (cls), cJSONScript.Eval(), cJSONScript.ExecCommand(), cStringBuilder (cls), cStringBuilder.Append(), cStringBuilder.AppendByVal(), cStringBuilder.AppendNL(), cStringBuilder.Insert() (+21 more)

### Community 8 - "Interface: Receive Data"
Cohesion: 0.10
Nodes (33): frmReceiveData (frm), frmReceiveData.ArchiveUnprocessed(), frmReceiveData.checkComStatus(), frmReceiveData.Command1_Click(), frmReceiveData.Command2_Click(), frmReceiveData.Command3_Click(), frmReceiveData.Command4_Click(), frmReceiveData.consoleCmd() (+25 more)

### Community 9 - "Interface: Request"
Cohesion: 0.06
Nodes (4): frmRequest (frm), frmRequest.cmdAdd_Click(), frmRequest.Form_Load(), mdiMain.mnuRequest_Click()

### Community 10 - "SA: Lab Exam"
Cohesion: 0.09
Nodes (23): ClassForms.showLabExams(), frmLabExam (frm), frmLabExam.cboResultType_Click(), frmLabExam.ClearValues(), frmLabExam.cmdAddSave_Click(), frmLabExam.cmdCancel_Click(), frmLabExam.cmdDown_Click(), frmLabExam.cmdGenerateSort_Click() (+15 more)

### Community 11 - "Interface: Main"
Cohesion: 0.09
Nodes (16): frmSettings.cmdSet_Click(), mdiMain (frm), mdiMain.mnuBackUp_Click(), mdiMain.mnuComSet_Click(), mdiMain.mnuControlPanel_Click(), mdiMain.mnuDisconnect_Click(), mdiMain.mnuDoctors_Click(), mdiMain.mnuExit_Click() (+8 more)

### Community 12 - "SA: Form Settings"
Cohesion: 0.10
Nodes (15): ClassForms.showFormSettings(), ClassForms.showSelectionLists(), frmFormSettings (frm), frmFormSettings.fillCboSelection(), frmFormSettings.findCboIndex(), frmFormSettings.Form_Load(), frmSelectionLists (frm), frmSelectionLists.cmdAdd_Click() (+7 more)

### Community 13 - "Interface: OBISetiings"
Cohesion: 0.11
Nodes (19): frmOBISetiings (frm), frmOBISetiings.cboServerType_Change(), frmOBISetiings.cboServerType_Click(), frmOBISetiings.cmdConnect_Click(), frmOBISetiings.cmdTestApi_Click(), frmOBISetiings.FillCboList(), frmOBISetiings.findCboIndex(), frmOBISetiings.findcboType() (+11 more)

### Community 14 - "Interface: Settings"
Cohesion: 0.08
Nodes (5): frmSettings (frm), frmSettings.cmdSetPlugins_Click(), frmSettings.List1_Click(), frmSettings.LoadFileToTB(), frmSettings.txtDetails_DblClick()

### Community 15 - "Interface: Search Result"
Cohesion: 0.11
Nodes (18): frmSearchResult (frm), frmSearchResult.cmdCancel_Click(), frmSearchResult.HighlightOver(), frmSearchResult.Image4_Click(), frmSearchResult.Image5_Click(), frmSearchResult.Image6_Click(), frmSearchResult.lblArchive_Click(), frmSearchResult.lblArchive_MouseMove() (+10 more)

### Community 16 - "SA: Report Formats"
Cohesion: 0.15
Nodes (17): ClassForms.showReportFormats(), frmReportFormats (frm), frmReportFormats.assignValues(), frmReportFormats.ClearValues(), frmReportFormats.cmdBrowse_Click(), frmReportFormats.cmdCancel_Click(), frmReportFormats.cmdSave_Click(), frmReportFormats.cmdUpdate_Click() (+9 more)

### Community 17 - "Interface: Db Man"
Cohesion: 0.16
Nodes (17): frmDbMan (frm), frmDbMan.cmdArchive_Click(), frmDbMan.cmdChangePassword_Click(), frmDbMan.cmdCompactDB_MouseUp(), frmDbMan.cmdFixLink_Click(), frmDbMan.cmdGenerateResultID_Click(), frmDbMan.cmdRestoreDB_Click(), frmDbMan.showStatus() (+9 more)

### Community 18 - "Interface: Registration"
Cohesion: 0.12
Nodes (15): frmRegistration (frm), frmRegistration.cmdNewCode_Click(), frmRegistration.EnControls(), frmRegistration.ISHEX(), frmRegistration.ProcessResponse(), frmRegistration.RetrieveRegistration(), frmRegistration.SavePayload(), frmRegistration.txtCode1_Change() (+7 more)

### Community 19 - "Interface: Main (2)"
Cohesion: 0.11
Nodes (15): frmDailyBackup.Form_Activate(), frmSettings.cmdSave_MouseUp(), modMain (bas), modMain.CLargeInt(), modMain.DailyBackup(), modMain.EnumRegistryKeys(), modMain.EnumRegistryValuesEx(), modMain.FilenameToDate() (+7 more)

### Community 20 - "RE: Search Request"
Cohesion: 0.17
Nodes (17): ClassForms.SearchRequestFormTop(), frmSearchRequest (frm), frmSearchRequest.cmdCancelSearchHIS_Click(), frmSearchRequest.cmdNew_Click(), frmSearchRequest.cmdSearch_Click(), frmSearchRequest.delPrevExtDB(), frmSearchRequest.Form_Load(), frmSearchRequest.initGrid() (+9 more)

### Community 21 - "Interface: Search"
Cohesion: 0.10
Nodes (10): frmSearch (frm), frmSearch.cboStatus_Click(), frmSearch.cmdClear_Click(), frmSearch.cmdSearch_Click(), frmSearch.Form_Load(), frmSearch.Form_Resize(), frmSearch.Form_Unload(), frmSearch.getSearchString() (+2 more)

### Community 22 - "RE: Blowfish"
Cohesion: 0.22
Nodes (20): clsBlowfish (cls), clsBlowfish.Append(), clsBlowfish.DecryptBlock(), clsBlowfish.DecryptByte(), clsBlowfish.DecryptFile(), clsBlowfish.DecryptString(), clsBlowfish.DeHex(), clsBlowfish.EncryptBlock() (+12 more)

### Community 23 - "RE: Request Setting"
Cohesion: 0.12
Nodes (12): frmRequestSetting (frm), frmRequestSetting.chkEditSort_Click(), frmRequestSetting.cmdDeptWardRoom_Click(), frmRequestSetting.cmdSave_Click(), frmRequestSetting.msfgRequestField_Click(), frmRequestSetting.msfgRequestField_LeaveCell(), frmRequestSetting.msfgRequestField_Scroll(), frmRequestSetting.txtLabel_KeyDown() (+4 more)

### Community 24 - "SA: PIDsearch"
Cohesion: 0.14
Nodes (13): ClassForms.showPIDsearch(), frmPIDsearch (frm), frmPIDsearch.assignValues(), frmPIDsearch.ClearValues(), frmPIDsearch.cmdSelect_Click(), frmPIDsearch.Form_Load(), frmPIDsearch.GenerateTestPacket(), frmPIDsearch.GetError() (+5 more)

### Community 25 - "Interface: Blowfish"
Cohesion: 0.22
Nodes (20): clsBlowfish (cls), clsBlowfish.Append(), clsBlowfish.DecryptBlock(), clsBlowfish.DecryptByte(), clsBlowfish.DecryptFile(), clsBlowfish.DecryptString(), clsBlowfish.DeHex(), clsBlowfish.EncryptBlock() (+12 more)

### Community 26 - "SA: Result Entry (2)"
Cohesion: 0.12
Nodes (21): frmResultEntry.ConvertToNumeric(), frmResultEntry.DisplayResultForm(), frmResultEntry.FormulaValue(), frmResultEntry.getResultValue(), frmResultEntry.getTotalDiff(), frmResultEntry.hasDiff(), frmResultEntry.lblClassName_Change(), frmResultEntry.lblDiff_Change() (+13 more)

### Community 27 - "SA: Result Entry (3)"
Cohesion: 0.18
Nodes (21): frmResultEntry.txtAge_GotFocus(), frmResultEntry.txtBirthDate_GotFocus(), frmResultEntry.txtCivilStatus_GotFocus(), frmResultEntry.txtClassID_GotFocus(), frmResultEntry.txtEntry_GotFocus(), frmResultEntry.txtFirstName_GotFocus(), frmResultEntry.txtLastName_GotFocus(), frmResultEntry.txtMedtech_GotFocus() (+13 more)

### Community 28 - "RE: Selection"
Cohesion: 0.13
Nodes (9): frmSelection (frm), frmSelection.cmdSelect_Click(), frmSelection.Form_Load(), frmSelection.msfgResult_DblClick(), frmSelection.msfgResult_KeyPress(), frmSelection.setGrid(), frmSelection.txtFilter_Change(), frmSelection.txtFilter_KeyDown() (+1 more)

### Community 29 - "RE: Class Main"
Cohesion: 0.16
Nodes (18): ClassForms.getDispTestReqSum(), ClassForms.wConnection(), frmEditRequest.Form_Load(), frmEditRequest.ReqIsLocked(), frmEditRequest.ReqLock(), frmEditRequest.setmsfgResultDemo(), frmEditRequest.setmsfgResultReq(), frmEditRequest.sizeFix() (+10 more)

### Community 30 - "SA: Lab Profiles"
Cohesion: 0.12
Nodes (10): ClassForms.showLabProfiles(), frmLabProfiles (frm), frmLabProfiles.cmdSave_Click(), frmLabProfiles.cmdTemplate_Click(), frmLabProfiles.Form_Load(), frmLabProfiles.setFlexGrid(), frmTemplate (frm), frmTemplate.cmdSelect_Click() (+2 more)

### Community 31 - "Interface: Main (3)"
Cohesion: 0.18
Nodes (17): frmProcess (frm), frmProcess.processData(), frmProcess.SaveResult(), frmProcess.Timer1_Timer(), frmReceiveData.processData(), frmReceiveData.SaveResult(), frmReceiveData.saveUcontrol(), frmReceiveData.tmrAutoPrint_Timer() (+9 more)

### Community 32 - "Interface: Work List"
Cohesion: 0.21
Nodes (16): frmWorkList (frm), frmWorkList.cboFilter_Click(), frmWorkList.cmdPos_Click(), frmWorkList.cmdRefresh_Click(), frmWorkList.cmdRemoveFromWorklist_Click(), frmWorkList.cmdSend_Click(), frmWorkList.dtpDate_Click(), frmWorkList.fillCboFilter() (+8 more)

### Community 33 - "RE: Class Forms"
Cohesion: 0.12
Nodes (10): ClassForms (cls), ClassForms.CheckInSample(), ClassForms.Class_Initialize(), ClassForms.EditRequest(), ClassForms.printBarCode(), ClassForms.SearchRequestFormLeft(), ClassForms.showRequestSettings(), ClassForms.showSearchForm() (+2 more)

### Community 34 - "RE: Exam Selection"
Cohesion: 0.16
Nodes (11): frmExamSelection (frm), frmExamSelection.cmbSpecimen_Click(), frmExamSelection.cmdSelect_Click(), frmExamSelection.Form_Load(), frmExamSelection.msfgResult_DblClick(), frmExamSelection.msfgResult_KeyPress(), frmExamSelection.setGrid(), frmExamSelection.setListCmbSelection() (+3 more)

### Community 35 - "SA: Worklist Report"
Cohesion: 0.12
Nodes (11): ClassForms.showOtherReports(), ClassForms.showWorklistReport(), frmInputBox.Form_Load(), frmLabProfileExams.Form_Load(), frmNormalValues.Form_Load(), frmOtherReports (frm), frmOtherReports.Form_Load(), frmPrintCopies.Form_Load() (+3 more)

### Community 36 - "Interface: DBsettings"
Cohesion: 0.13
Nodes (6): frmDBsettings (frm), frmDBsettings.cmdSave_Click(), frmDBsettings.fillcboServerType(), frmDBsettings.Form_Load(), frmDBsettings.setCBOServer(), frmDBsettings.txtSchema_GotFocus()

### Community 37 - "RE: Edit Request (2)"
Cohesion: 0.23
Nodes (16): frmEditRequest.AddExamDet(), frmEditRequest.blnExamOccured(), frmEditRequest.cmdExamMenu_Click(), frmEditRequest.RemoveFromTestToDel(), frmEditRequest.txtmsfgExam_GotFocus(), frmEditRequest.txtmsfgExam_KeyDown(), frmEditRequest.txtmsfgExam_KeyPress(), frmEditRequest.txtmsfgResult_GotFocus() (+8 more)

### Community 38 - "SA: Input Box"
Cohesion: 0.12
Nodes (3): ClassForms.showInputBox(), frmInputBox (frm), frmInputBox.txtValue_GotFocus()

### Community 39 - "SA: Result Entry (4)"
Cohesion: 0.14
Nodes (16): frmResultEntry.ClearValues(), frmResultEntry.cmdNext_Click(), frmResultEntry.cmdPIDSearch_Click(), frmResultEntry.cmdPrevious_Click(), frmResultEntry.cmdSave_Click(), frmResultEntry.fillLists(), frmResultEntry.Form_Activate(), frmResultEntry.Form_Load() (+8 more)

### Community 40 - "Interface: Plug In Test"
Cohesion: 0.23
Nodes (14): frmPlugInTest (frm), frmPlugInTest.cboFormatting_Click(), frmPlugInTest.cmdNextRecord_Click(), frmPlugInTest.cmdPreviousRecord_Click(), frmPlugInTest.fileHosts_Click(), frmPlugInTest.fillcboFormatting(), frmPlugInTest.Form_Load(), frmPlugInTest.ViewItems() (+6 more)

### Community 41 - "Shared: Logbook"
Cohesion: 0.13
Nodes (12): Dataset.cls (outside repo), Datasets.cls (outside repo), Graph.ctl (outside repo), modConstants.bas (outside repo), modFiles.bas (outside repo), modTypes.bas (outside repo), Point.cls (outside repo), Points.cls (outside repo) (+4 more)

### Community 42 - "RE: Del Req"
Cohesion: 0.21
Nodes (12): ClassForms.DelReq(), frmDelReq (frm), frmDelReq.AddExamDet(), frmDelReq.blnExamOccured(), frmDelReq.DelExamDet(), frmDelReq.msfgResult_Click(), frmDelReq.msfgResult_EnterCell(), frmDelReq.msfgResult_LeaveCell() (+4 more)

### Community 43 - "SA: Class Forms"
Cohesion: 0.13
Nodes (3): ClassForms (cls), ClassForms.showNormalValues(), ClassForms.showResultEntry()

### Community 44 - "SA: Print Copies"
Cohesion: 0.13
Nodes (7): ClassForms.showPrintCopies(), frmPrintCopies (frm), frmPrintCopies.txtCopy1_Change(), frmPrintCopies.txtCopy1_GotFocus(), frmPrintCopies.txtCopy2_Change(), frmPrintCopies.txtCopy2_GotFocus(), frmPrintCopies.txtCopy2Tag_GotFocus()

### Community 45 - "Interface: License Manager"
Cohesion: 0.21
Nodes (12): frmLicenseManager (frm), frmLicenseManager.cmdGetLicense_Click(), frmLicenseManager.cmdRegister_Click(), frmLicenseManager.DeletePayload(), frmLicenseManager.GetLicences(), frmLicenseManager.LoadDetails(), frmLicenseManager.OpenRegistrationServer(), frmLicenseManager.ProcessResponse() (+4 more)

### Community 46 - "Interface: Check"
Cohesion: 0.25
Nodes (15): modCheck (bas), modCheck.ASTMChecksum(), modCheck.bin2dec(), modCheck.CheckSum8(), modCheck.CRC16_CCITT(), modCheck.CRC16_CCITT_M(), modCheck.dec2bin(), modCheck.isValidCS() (+7 more)

### Community 47 - "SA: Rem Template"
Cohesion: 0.22
Nodes (10): ClassForms.GetSearchRecordSet(), ClassForms.SearchCode(), ClassForms.showRemarksTemplate(), ClassForms.showSearchCode(), frmClassifications.cmdSave_Click(), frmRemTemplate (frm), frmRemTemplate.cmdAdd_Click(), frmRemTemplate.Form_Load() (+2 more)

### Community 48 - "SA: Pathologist"
Cohesion: 0.22
Nodes (9): ClassForms.showPathologist(), frmPathologist (frm), frmPathologist.cmdAdd_Click(), frmPathologist.cmdDelete_Click(), frmPathologist.Form_Load(), frmPathologist.msfgResult_KeyDown(), frmPathologist.setFlexGrid(), frmResultEntry.Label1_Click() (+1 more)

### Community 49 - "SA: Search Code"
Cohesion: 0.16
Nodes (6): frmSearchCode (frm), frmSearchCode.cmdSelect_Click(), frmSearchCode.Form_Load(), frmSearchCode.msfgResult_DblClick(), frmSearchCode.setFlexGrid(), frmSearchCode.txtFilter_GotFocus()

### Community 50 - "RE: Summary Request"
Cohesion: 0.17
Nodes (7): ClassForms.showReqSummary(), frmRequest (frm), frmSummaryRequest (frm), frmSummaryRequest.Form_Load(), frmSummaryRequest.setMsfgSummary(), modClassMain.showSummaryRequest(), HostForms_RE (HostForms_RE.dll, OleDll)

### Community 51 - "SA: Hospital"
Cohesion: 0.19
Nodes (8): ClassForms.showHospitalInfo(), ClassForms.wConnection(), frmHospital (frm), frmHospital.cmdSave_Click(), frmHospital.Form_Load(), frmHospital.replaceStr(), frmDataBaseCopy.cmdCheck_Click(), frmDbMan.Form_Load()

### Community 53 - "Interface: Settings (2)"
Cohesion: 0.15
Nodes (13): frmSettings.File1_Click(), frmSettings.File2_Click(), frmSettings.fillcboExport(), frmSettings.fillcboScheme(), frmSettings.findcboType(), frmSettings.Form_Load(), frmSettings.initCdgExportLocation(), frmSettings.ListSubDirs() (+5 more)

### Community 54 - "SA: Medtech"
Cohesion: 0.27
Nodes (7): ClassForms.showMedTech(), frmMedtech (frm), frmMedtech.cmdAdd_Click(), frmMedtech.cmdDelete_Click(), frmMedtech.Form_Load(), frmMedtech.msfgResult_KeyDown(), frmMedtech.setFlexGrid()

### Community 55 - "SA: Result Entry (5)"
Cohesion: 0.18
Nodes (12): frmResultEntry.FillCBOResult(), frmResultEntry.findCboIndex(), frmResultEntry.getIResultDate(), frmResultEntry.hideBlankInstrumentResults(), frmResultEntry.msfgResult_Click(), frmResultEntry.msfgResult_EnterCell(), frmResultEntry.msfgResult_GotFocus(), frmResultEntry.msfgResult_Scroll() (+4 more)

### Community 56 - "Interface: Control Panel"
Cohesion: 0.20
Nodes (11): frmControlPanel (frm), frmControlPanel.Form_Load(), frmControlPanel.lvwItems_DblClick(), frmControlPanel.lvwItems_KeyDown(), frmInputBox.Form_Load(), frmProcess.Form_Load(), frmReceiveData.Form_Activate(), frmSearch.Form_Activate() (+3 more)

### Community 57 - "Interface: Registration (2)"
Cohesion: 0.17
Nodes (12): frmDBsettings.txtDatabase_GotFocus(), frmDBsettings.txtPassword_GotFocus(), frmDBsettings.txtServerName_GotFocus(), frmDBsettings.txtUsername_GotFocus(), frmRegistration.txtCode1_GotFocus(), frmRegistration.txtCode2_GotFocus(), frmRegistration.txtCode3_GotFocus(), frmRegistration.txtCode4_GotFocus() (+4 more)

### Community 58 - "RE: Req Rates"
Cohesion: 0.24
Nodes (7): ClassForms.showReqRates(), frmReqRates (frm), frmReqRates.cmdSaveDiscount_Click(), frmReqRates.Form_Load(), frmReqRates.setmsfgRates(), frmReqRates.setmsfgReqDiscounts(), frmReqRates.txtDiscountName_GotFocus()

### Community 59 - "SA: Physician"
Cohesion: 0.31
Nodes (7): ClassForms.showPhysician(), frmPhysician (frm), frmPhysician.cmdAdd_Click(), frmPhysician.cmdDelete_Click(), frmPhysician.Form_Load(), frmPhysician.msfgResult_KeyDown(), frmPhysician.setFlexGrid()

### Community 60 - "Interface: Search Result (2)"
Cohesion: 0.25
Nodes (9): frmResultEntry.getUpdatedValues(), frmSearchResult.ArchiveResults(), frmSearchResult.cmdDelete_Click(), frmSearchResult.DeleteResults(), frmSearchResult.Image7_Click(), frmSearchResult.lblUpload_Click(), frmSearchResult.MergeResults(), frmSearchResult.msfgResult_KeyDown() (+1 more)

### Community 61 - "Interface: Main (4)"
Cohesion: 0.24
Nodes (10): frmLicenseManager.Form_Load(), frmRegistration.cmdRegister_Click(), frmRegistration.Form_Load(), frmRegistration.OpenRegistrationServer(), mdiMain.mnuactivation_Click(), modMain.GetCpuID(), modMain.GetSAVersion(), modMain.GetSerialNumber() (+2 more)

### Community 62 - "Interface: Main (5)"
Cohesion: 0.24
Nodes (10): frmReceiveData.Form_Unload(), mdiMain.MDIForm_Load(), mdiMain.MDIForm_Resize(), mdiMain.mnuConnect_Click(), mdiMain.mnuSendID_Click(), mdiMain.mnuSendReceive_Click(), mdiMain.tbrMain_ButtonMenuClick(), modMain.InitCommonControlsVB() (+2 more)

### Community 63 - "Interface: Activation"
Cohesion: 0.22
Nodes (5): frmActivation (frm), frmActivation.cmdEnter_Click(), frmActivation.saveData(), frmActivation.txtSerial_Change(), frmActivation.txtSerial_GotFocus()

### Community 64 - "Interface: Warning"
Cohesion: 0.20
Nodes (6): frmDailyBackup (frm), frmSelectInstrument (frm), frmWarning (frm), frmWorkListAdd (frm), mdiMain.tmrDailyBackup_Timer(), HostInterface (Interface.exe, Exe)

### Community 65 - "Interface: Result Viewer"
Cohesion: 0.24
Nodes (6): frmResultViewer (frm), frmResultViewer.cmdSave_Click(), frmResultViewer.Form_Load(), frmResultViewer.setmsfgResult(), frmResultViewer.txtSampleNum_GotFocus(), frmResultViewer.UpdateData()

### Community 66 - "Interface: Search Result (3)"
Cohesion: 0.20
Nodes (10): frmSearchResult.cmdPreview_Click(), frmSearchResult.Form_Load(), frmSearchResult.Form_Resize(), frmSearchResult.Image1_Click(), frmSearchResult.lblPreview_Click(), frmSearchResult.msfgResult_Click(), frmSearchResult.msfgResult_MouseDown(), frmSearchResult.msfgResult_Scroll() (+2 more)

### Community 67 - "RE: Sample In"
Cohesion: 0.25
Nodes (3): frmSampleIn (frm), frmSampleIn.Form_Load(), frmSampleIn.setmsfgSamples()

### Community 68 - "SA: Classifications"
Cohesion: 0.25
Nodes (4): ClassForms.showClassifications(), frmClassifications (frm), frmClassifications.Form_Load(), frmClassifications.setFlexGrid()

### Community 69 - "SA: Result Entry (6)"
Cohesion: 0.25
Nodes (9): frmResultEntry.AssignPIDValues(), frmResultEntry.GetIdFromName(), frmResultEntry.getNVclass(), frmResultEntry.setValues(), frmResultEntry.txtBirthDate_Change(), frmResultEntry.txtClassID_KeyDown(), frmResultEntry.txtPatientID_KeyDown(), frmResultEntry.txtSex_Change() (+1 more)

### Community 70 - "Interface: Base64"
Cohesion: 0.47
Nodes (9): modBase64 (bas), modBase64.Base64Decode(), modBase64.Base64DecodeString(), modBase64.Base64Encode(), modBase64.Base64Encode2(), modBase64.Base64EncodeString(), modBase64.ConvertBytesToString(), modBase64.ConvertStringToBytes() (+1 more)

### Community 71 - "SA: Specimen"
Cohesion: 0.43
Nodes (5): frmSpecimen (frm), frmSpecimen.cmdAdd_Click(), frmSpecimen.Form_Load(), frmSpecimen.msfgResult_KeyDown(), frmSpecimen.setFlexGrid()

### Community 72 - "Interface: Class Forms"
Cohesion: 0.29
Nodes (7): ClassForms (cls), ClassForms.showInputBox(), ClassForms.ShowResultViewer(), frmDbMan.cmdClearBuffer_Click(), frmSettings.cmdTestPlugin_Click(), mdiMain.mnuClearBuffer_Click(), mdiMain.mnuPlugTest_Click()

### Community 73 - "Interface: Main (6)"
Cohesion: 0.25
Nodes (7): frmSettings.cmdClose_Click(), mdiMain.mnuReportFormats_Click(), modMain.checkResultString(), modMain.InitPlugin(), modMain.Main(), modMain.setResultReport(), modMain.VerifyRequestDB()

### Community 74 - "SA: Add Profiles"
Cohesion: 0.33
Nodes (4): ClassForms.showAddProfile(), frmAddProfiles (frm), frmAddProfiles.Form_Load(), frmAddProfiles.setListView()

### Community 75 - "Interface: Settings (3)"
Cohesion: 0.29
Nodes (7): frmSettings.Check1_Click(), frmSettings.chkRegistry_Click(), frmSettings.chkTaskM_Click(), frmSettings.chkUsbStor_Click(), frmSettings.cmdRegistry_Click(), frmSettings.cmdTaskMAn_Click(), frmSettings.SaveValue()

### Community 76 - "Interface: Statistics"
Cohesion: 0.43
Nodes (5): frmStatistics (frm), frmStatistics.countData(), frmStatistics.Form_Load(), frmStatistics.getAverage(), frmStatistics.setmsfgResult()

### Community 77 - "Interface: Wall"
Cohesion: 0.43
Nodes (6): frmWall (frm), frmWall.changeSize(), frmWall.Form_Load(), frmWall.Form_Resize(), frmWall.getIP(), mdiMain.MDIForm_Unload()

### Community 78 - "Interface: About"
Cohesion: 0.47
Nodes (4): frmAbout (frm), frmAbout.cmdSysInfo_Click(), frmAbout.GetKeyValue(), frmAbout.StartSysInfo()

### Community 79 - "Interface: Backup List"
Cohesion: 0.40
Nodes (4): frmBackupList (frm), frmBackupList.Form_Load(), frmBackupList.ListSubDirs(), modMain.showBackupList()

### Community 80 - "Interface: Bar Code"
Cohesion: 0.40
Nodes (5): frmBarCode (frm), frmBarCode.cmdPrint_Click(), frmBarCode.cmdPrint_GotFocus(), mdiMain.mnuBarCode_Click(), modMain.Mod43()

### Community 81 - "Interface: Charmap"
Cohesion: 0.40
Nodes (3): frmCharmap (frm), frmCharmap.cmdInsert_Click(), frmCharmap.msfgResult_Click()

### Community 82 - "Interface: Archive"
Cohesion: 0.40
Nodes (3): ClassForms.showArchive(), frmArchive (frm), frmArchive.Form_Load()

### Community 84 - "Interface: Search Result (4)"
Cohesion: 0.40
Nodes (5): frmSearchResult.CanMerge(), frmSearchResult.GetResultID(), frmSearchResult.Image8_Click(), frmSearchResult.lblMerge_Click(), frmSearchResult.lblMerge_MouseMove()

### Community 85 - "Interface: Search Result (5)"
Cohesion: 0.40
Nodes (5): frmSearchResult.cmdOpen_Click(), frmSearchResult.Image2_Click(), frmSearchResult.lblEdit_Click(), frmSearchResult.msfgResult_DblClick(), frmSearchResult.setSelectedItems()

### Community 86 - "Interface: Search Result (6)"
Cohesion: 0.40
Nodes (5): frmSearchResult.cmdPrint_Click(), frmSearchResult.Image3_Click(), frmSearchResult.lblPrint_Click(), frmSearchResult.msfgResult_SelChange(), frmSearchResult.PrintResults()

### Community 88 - "Interface: Settings (4)"
Cohesion: 0.50
Nodes (4): frmSettings.chkPlugin_Click(), frmSettings.clearDetails(), frmSettings.fileList_Click(), frmSettings.setDetails()

## Knowledge Gaps
- **10 isolated node(s):** `frmSelectInstrument (frm)`, `frmRequest (frm)`, `Graph.ctl (outside repo)`, `Dataset.cls (outside repo)`, `Datasets.cls (outside repo)` (+5 more)
  These have ≤1 connection - possible missing edges. (Counts symbols only; 625 node(s) total have ≤1 connection when file, concept and rationale nodes are included.)
- **4 thin communities (<3 nodes) omitted from report** — run `graphify query` to explore isolated nodes.

## Suggested Questions
_Questions this graph is uniquely positioned to answer:_

- **Why does `modMain.setHostForms()` connect `Interface: Db Man` to `Interface: Main (4)`?**
  _High betweenness centrality (0.011) - this node is a cross-community bridge._
- **Why does `ClassForms (cls)` connect `SA: Class Forms` to `SA: Rpt Export`, `SA: Lab Profile Exams`, `SA: Worklist Report`, `SA: Classifications`, `SA: Input Box`, `Shared: Logbook`, `SA: Add Profiles`, `SA: Lab Exam`, `SA: Form Settings`, `SA: Print Copies`, `SA: Rem Template`, `SA: Pathologist`, `SA: Report Formats`, `SA: Hospital`, `SA: Medtech`, `SA: PIDsearch`, `SA: Physician`, `SA: Lab Profiles`?**
  _High betweenness centrality (0.011) - this node is a cross-community bridge._
- **Why does `modMain.setButtons()` connect `Interface: Main (5)` to `Interface: Main`?**
  _High betweenness centrality (0.008) - this node is a cross-community bridge._
- **What connects `frmSelectInstrument (frm)`, `frmRequest (frm)`, `Graph.ctl (outside repo)` to the rest of the system?**
  _10 weakly-connected nodes found - possible documentation gaps or missing edges._
- **Should `SA: Result Entry` be split into smaller, more focused modules?**
  _Cohesion score 0.019798210546354465 - nodes in this community are weakly interconnected._
- **Should `SA: Rpt Export` be split into smaller, more focused modules?**
  _Cohesion score 0.05764411027568922 - nodes in this community are weakly interconnected._
- **Should `SA: Lab Profile Exams` be split into smaller, more focused modules?**
  _Cohesion score 0.05075187969924812 - nodes in this community are weakly interconnected._