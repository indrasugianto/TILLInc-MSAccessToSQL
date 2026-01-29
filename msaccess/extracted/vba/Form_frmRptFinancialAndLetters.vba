' Module Name: Form_frmRptFinancialAndLetters
' Module Type: Document Module
' Lines of Code: 1310
' Extracted: 1/29/2026 4:12:26 PM

Option Compare Database
Option Explicit

Public ExportFileName As String, Loc As String, RetValue As Variant, CommandLine As String
Private MRSelectedMonth As String, MRSelectedYear As String, ErrorOccurred As Boolean, ReportName As String
Private Const MAX_PATH As Long = 260

Dim SigOnFileClientIndexedName As Variant

Private Sub AmendmentsReport_Click()
    If IsNull(MRSelectMonth) Or Len(MRSelectMonth) <= 0 Then
        MsgBox "You must select a month above.", vbOKOnly, "Error!"
        Exit Sub
    End If
    If IsNull(DCount("Pending", "tblContractsAmendments", "FY=" & MRFY)) Or DCount("Pending", "tblContractsAmendments", "FY=" & MRFY) = 0 Then
        MsgBox "There are currently no amendments for FY" & MRFY & ".", vbOKOnly, "No amendments to report"
        Exit Sub
    End If
    Call BriefDelay
    DoCmd.OpenReport "rptAMENDMENTS", acViewPreview, , "FY=" & MRFY, , MRFY & " " & MRNumDays & " " & MRSelectMonth
    Call LoopUntilClosed("rptAMENDMENTS", 3)
    Call DropTempTables
End Sub

Private Sub AutoDepositAuth_AfterUpdate()
    Dim WhichLetter As String, RepPayee As Variant

    Call BuildAllClients(False)
    SelectedClientIndexedName = DLookup("IndexedName", "AllClients", "ReverseDisplayName = """ & AutoDepositAuth & """")
    RepPayee = DLookup("RepresentativePayee", "tblPeopleClientsDemographics", "IndexedName=""" & SelectedClientIndexedName & """")
    If IsNull(RepPayee) Or Len(RepPayee) <= 0 Then
        MsgBox "Individual has no Representative Payee identified.  Please correct this.", vbOKOnly, "Error"
        AutoDepositAuth = Null
        Exit Sub
    End If
    If Len(SelectedClientIndexedName) > 0 Then
        If Left(RepPayee, 4) = "TILL" Then WhichLetter = "ltrAUTODEPOSITAUTHTILL" Else WhichLetter = "ltrAUTODEPOSITAUTH"
        Call BriefDelay
        Call ExecReport(WhichLetter)
    End If
    AutoDepositAuth = Null: Me.Refresh: Me.Requery: Me.Repaint
End Sub

Private Sub CFCMSortBy_AfterUpdate()
    Select Case CFCMSortBy
        Case "Client"
            Call BriefDelay: DoCmd.OpenReport "rptCFCMETHOD", acViewPreview, , , , "Client"
        Case "Location"
            Call BriefDelay: DoCmd.OpenReport "rptCFCMETHOD", acViewPreview, , , , "Location"
        Case Else
    End Select
    Call LoopUntilClosed("rptCFCMETHOD", 3)
    CFCMSortBy.DefaultValue = "<Select sort criteria>": Me.Refresh:     Me.Requery:     Me.Repaint
End Sub

Private Sub ContractsByDirector_Click()
    Call BriefDelay
    DoCmd.OpenReport "rptDIRCONTRACTS", acViewPreview, , "FY=" & MRFY, , MRFY
    Call LoopUntilClosed("rptDIRCONTRACTS", 3)
    Call DropTempTables
End Sub

Private Sub DDSAmendedMaxObligation_Click()
    Dim ExportFileName As String
    
    ExportFileName = Application.CurrentProject.Path & "\" & "TILLDB-Export-DDSMaxObligations-" & Format(Date, "yyyymmdd") & ".xls"
    If IsFileOpen(ExportFileName) Then
        If MsgBox(ExportFileName & " is already open.  Please close it and click OK to continue or Cancel to abort.", vbOKCancel, "ERROR!") = vbCancel Then
            MsgBox "Export aborted.", vbOKOnly, "Aborted"
            Exit Sub
        End If
    End If
    If Dir(ExportFileName) <> "" Then Kill ExportFileName
    DoCmd.TransferSpreadsheet acExport, acSpreadsheetTypeExcel9, "qryDDSMaxObligation", ExportFileName
    MsgBox "The requested information has been exported to " & ExportFileName & "." & vbCrLf & vbCrLf & "This export may contain information that is protected under HIPAA and other privacy laws.  This export must be securely stored at all times and must be deleted when no longer being used.", vbOKOnly, "Export Complete"
End Sub

Private Sub DDSChangesInWorks_Click()
    If IsNull(MRSelectMonth) Or Len(MRSelectMonth) <= 0 Then
        MsgBox "You must select a month above.", vbOKOnly, "Error!"
        Exit Sub
    End If
    Call BriefDelay
    DoCmd.OpenReport "rptCONINWORKS", acViewPreview, , "FY=" & MRFY, , MRFY & " " & MRNumDays & " " & MRSelectMonth
    Call LoopUntilClosed("rptCONINWORKS", 3)
    Call DropTempTables
End Sub

Private Sub DDSChangesInWorksSummary_Click()
On Error GoTo ShowMeError
    Dim Criteria As String, UnitsBilled As Variant, Multiplier As Integer, rst As Recordset, dbs As Database, SelectedMonth As String
    
    If IsNull(MRSelectMonth) Or Len(MRSelectMonth) <= 0 Then
        MsgBox "You must select a month above.", vbOKOnly, "Error!"
        Exit Sub
    End If
    SelectedMonth = MRSelectMonth
    
    TILLDataBase.Execute "DELETE * FROM [~CONINWORKSSummary]", dbSeeChanges: Call BriefDelay
    DoCmd.SetWarnings False
    DoCmd.OpenQuery "qrySeedCONINWORKSSummary"
    DoCmd.SetWarnings True

    Set dbs = CurrentDb
    Set rst = dbs.OpenRecordset("~CONINWORKSSummary", , dbSeeChanges): Call BriefDelay
        Do
            Select Case rst![ContractUnit]
                Case "Day", "Quarter Hour", "Month"
                    Select Case Left(SelectedMonth, 3)
                        Case "Jul": Multiplier = 1
                        Case "Aug": Multiplier = 2
                        Case "Sep": Multiplier = 3
                        Case "Oct": Multiplier = 4
                        Case "Nov": Multiplier = 5
                        Case "Dec": Multiplier = 6
                        Case "Jan": Multiplier = 7
                        Case "Feb": Multiplier = 8
                        Case "Mar": Multiplier = 9
                        Case "Apr": Multiplier = 10
                        Case "May": Multiplier = 11
                        Case "Jun": Multiplier = 12
                    End Select
                    UnitsBilled = Round((rst![Units] / 12) * Multiplier)
                Case "Day (9 months)"
                    Select Case Left(SelectedMonth, 3)
                        Case "Jul": Multiplier = 1
                        Case "Aug": Multiplier = 2
                        Case "Sep": Multiplier = 3
                        Case "Oct": Multiplier = 4
                        Case "Nov": Multiplier = 5
                        Case "Dec": Multiplier = 6
                        Case "Jan": Multiplier = 7
                        Case "Feb": Multiplier = 8
                        Case "Mar": Multiplier = 9
                        Case "Apr": Multiplier = 9
                        Case "May": Multiplier = 9
                        Case "Jun": Multiplier = 9
                    End Select
                    UnitsBilled = Round((rst![Units] / 9) * Multiplier)
                Case "Day (6 months)"
                    Select Case Left(SelectedMonth, 3)
                        Case "Jul": Multiplier = 1
                        Case "Aug": Multiplier = 2
                        Case "Sep": Multiplier = 3
                        Case "Oct": Multiplier = 4
                        Case "Nov": Multiplier = 5
                        Case "Dec": Multiplier = 6
                        Case "Jan": Multiplier = 6
                        Case "Feb": Multiplier = 6
                        Case "Mar": Multiplier = 6
                        Case "Apr": Multiplier = 6
                        Case "May": Multiplier = 6
                        Case "Jun": Multiplier = 6
                    End Select
                    UnitsBilled = Round((rst![Units] / 6) * Multiplier)
                Case "Day (3 months)"
                    Select Case Left(SelectedMonth, 3)
                        Case "Jul": Multiplier = 1
                        Case "Aug": Multiplier = 2
                        Case "Sep": Multiplier = 3
                        Case "Oct": Multiplier = 4
                        Case "Nov": Multiplier = 5
                        Case "Dec": Multiplier = 6
                        Case "Jan": Multiplier = 6
                        Case "Feb": Multiplier = 6
                        Case "Mar": Multiplier = 6
                        Case "Apr": Multiplier = 6
                        Case "May": Multiplier = 6
                        Case "Jun": Multiplier = 6
                    End Select
                    UnitsBilled = Round((rst![Units] / 6) * Multiplier)
                Case Else
                    UnitsBilled = -1
            End Select

            rst.Edit
            If UnitsBilled >= 0 Then
                rst![UnitsShouldHaveBilled] = UnitsBilled
            Else
                rst![UnitsShouldHaveBilled] = Null
            End If
            rst.Update
            rst.MoveNext
        Loop Until rst.EOF
    
    rst.Close
    dbs.Close
    SysCmdResult = SysCmd(5)

    ExportFileName = Application.CurrentProject.Path & "\" & "TILLDB-Export-ContractsInWorks-" & Format(Date, "yyyymmdd") & ".xls"
    If IsFileOpen(ExportFileName) Then
        If MsgBox(ExportFileName & " is already open.  Please close it and click OK to continue or Cancel to abort.", vbOKCancel, "ERROR!") = vbCancel Then
            MsgBox "Export aborted.", vbOKOnly, "Aborted"
            Exit Sub
        End If
    End If
    
    If Dir(ExportFileName) <> "" Then Kill ExportFileName
    DoCmd.TransferSpreadsheet acExport, acSpreadsheetTypeExcel9, "~CONINWORKSSummary", ExportFileName
    MsgBox "The requested information has been exported to " & ExportFileName & "." & vbCrLf & vbCrLf & "This export may contain information that is protected under HIPAA and other privacy laws.  This export must be securely stored at all times and must be deleted when no longer being used.", _
        vbOKOnly, "Export Complete"

    Exit Sub
ShowMeError:
    MsgBox "Error # " & Str(Err.Number) & " was generated by " & Me.Name & Chr(13) & Err.Description, vbOKOnly, "Error", Err.HelpFile, Err.HelpContext
End Sub

Private Sub GPMonthlyEntries_Click()
    If IsNull(MRSelectMonth) Or Len(MRSelectMonth) <= 0 Then
        MsgBox "You must select a month above.", vbOKOnly, "Error!"
        Exit Sub
    End If
    Call BriefDelay: DoCmd.OpenReport "rptGPMONTHLY", acViewPreview, , "FY=" & MRFY, , MRFY & " " & MRNumDays & " " & MRSelectMonth
    Call LoopUntilClosed("rptGPMONTHLY", 3)
    Call DropTempTables
End Sub

Private Sub GUARDIANINFORMATIONAllClients_Click()
    ReportName = "ltrGUARDIANINFOAllClients"
    Call DropTempTables
    TILLDataBase.Execute "SELECT tblPeopleFamily.*, tblPeopleClientsDemographics.RepPayeeIsTILL INTO temptbl " & _
        "FROM (((((tblPeople LEFT JOIN tblPeopleFamily ON tblPeople.IndexedName = tblPeopleFamily.IndexedName) LEFT JOIN tblPeopleClientsDemographics ON tblPeopleFamily.ClientIndexedName = tblPeopleClientsDemographics.IndexedName) LEFT JOIN tblPeopleClientsCLOServices ON tblPeopleClientsDemographics.IndexedName = tblPeopleClientsCLOServices.IndexedName) LEFT JOIN tblPeopleClientsIndividualSupportServices ON tblPeopleClientsDemographics.IndexedName = tblPeopleClientsIndividualSupportServices.IndexedName) LEFT JOIN tblPeopleClientsResidentialServices ON tblPeopleClientsDemographics.IndexedName = tblPeopleClientsResidentialServices.IndexedName) LEFT JOIN tblPeopleClientsSHCServices ON tblPeopleClientsDemographics.IndexedName = tblPeopleClientsSHCServices.IndexedName " & _
        "WHERE (((tblPeople.IsFamilyGuardian)=True) AND ((tblPeople.IsDeceased)=False) AND ((tblPeopleFamily.Inactive)=False) AND " & _
        "((tblPeopleFamily.Guardian)=True) AND ((tblPeopleFamily.ClientIndexedName) Is Not Null) AND " & _
        "((tblPeopleClientsDemographics.RepPayeeIsTILL)=True) AND ((tblPeopleClientsDemographics.ActiveResidentialServices)=True) AND " & _
        "((tblPeopleClientsResidentialServices.Inactive)=False)) OR (((tblPeople.IsFamilyGuardian)=True) AND " & _
        "((tblPeople.IsDeceased)=False) AND ((tblPeopleFamily.Inactive)=False) AND ((tblPeopleFamily.Guardian)=True) AND " & _
        "((tblPeopleFamily.ClientIndexedName) Is Not Null) AND ((tblPeopleClientsDemographics.RepPayeeIsTILL)=True) AND " & _
        "((tblPeopleClientsDemographics.ActiveSHC)=True) AND ((tblPeopleClientsSHCServices.Inactive)=False)) OR " & _
        "(((tblPeople.IsFamilyGuardian)=True) AND ((tblPeople.IsDeceased)=False) AND ((tblPeopleFamily.Inactive)=False) AND " & _
        "((tblPeopleFamily.Guardian)=True) AND ((tblPeopleFamily.ClientIndexedName) Is Not Null) AND " & _
        "((tblPeopleClientsDemographics.RepPayeeIsTILL)=True) AND " & _
        "((tblPeopleClientsDemographics.ActiveCLO)=True) AND ((tblPeopleClientsCLOServices.Inactive)=False)) OR " & _
        "(((tblPeople.IsFamilyGuardian)=True) AND ((tblPeople.IsDeceased)=False) AND ((tblPeopleFamily.Inactive)=False) AND " & _
        "((tblPeopleFamily.Guardian)=True) AND ((tblPeopleFamily.ClientIndexedName) Is Not Null) AND " & _
        "((tblPeopleClientsDemographics.RepPayeeIsTILL)=True) AND ((tblPeopleClientsDemographics.ActiveIndivSupport)=True) AND " & _
        "((tblPeopleClientsIndividualSupportServices.Inactive)=False)) " & _
        "ORDER BY tblPeopleFamily.IndexedName;", dbSeeChanges: Call BriefDelay
    Call BriefDelay
    Call ExecReport("ltrGUARDIANINFOAllClients")
    Me.Refresh: Me.Requery: Me.Repaint
    Call DropTempTables
End Sub

Private Sub GuardSelectClient_AfterUpdate()
    ReportName = "ltrGUARDIANINFO"
    Call DropTempTables
    SelectedClientIndexedName = DLookup("IndexedName", "AllClients", "ReverseDisplayName = """ & GuardSelectClient & """")
    SelectedFirstName = DLookup("FirstName", "AllClients", "ReverseDisplayName = """ & GuardSelectClient & """")
    SelectedLastName = DLookup("LastName", "AllClients", "ReverseDisplayName = """ & GuardSelectClient & """")
    TILLDataBase.Execute "SELECT tblPeopleFamily.* INTO temptbl " & _
        "FROM tblPeople LEFT JOIN tblPeopleFamily ON tblPeople.IndexedName = tblPeopleFamily.IndexedName " & _
        "WHERE (((tblPeople.IsFamilyGuardian) = True) And ((tblPeople.IsDeceased) = False) And ((tblPeopleFamily.Inactive) = False) And ((tblPeopleFamily.Guardian) = True)) " & _
        "ORDER BY tblPeopleFamily.IndexedName;", dbSeeChanges: Call BriefDelay
    GuardianIndexedName = DLookup("IndexedName", "temptbl", "ClientIndexedName=""" & SelectedClientIndexedName & """")
    If Len(SelectedClientIndexedName) > 0 Then
        Call ExecReport(ReportName)
    Else
        MsgBox "This individual has no guardians.  Form cannot be created.", vbOKOnly, "Error!"
    End If
    GuardSelectClient = Null: Me.Refresh: Me.Requery: Me.Repaint
End Sub

Private Sub IFCompletelyBlank_Click()
    If ABILetter Then ReportName = "ltrINDIVFINANCESCOMPLETELYBLANKABI" Else ReportName = "ltrINDIVFINANCESCOMPLETELYBLANK"
    Call ExecReport(ReportName)
    Call DropTempTables
End Sub

Private Sub MRSelectMonth_Enter()
    If Not (IsNull(MRSelectMonth) Or Len(MRSelectMonth)) <= 0 Then
        MsgBox "You cannot change this value once it has been set.  Close this form and select ""Financial Reports and Letters"" again to use a different month/year.", vbOKOnly, "Error!"
        GPMonthlyEntries.SetFocus
        Exit Sub
    End If
End Sub

Private Sub NHBillings_Click()
    If IsNull(MRSelectMonth) Or Len(MRSelectMonth) <= 0 Then
        MsgBox "You must select a month above.", vbOKOnly, "Error!"
        Exit Sub
    End If
    ReportFY = CLng(MRFY): ReportMonth = MRSelectMonth
    Call BriefDelay: DoCmd.OpenReport "rptNHBILLS", acViewPreview, , "FY=" & MRFY, , MRFY & " " & MRSelectMonth: Call LoopUntilClosed("rptNHBILLS", 3)
    Call DropTempTables
End Sub

Private Sub VictoriaBillings_Click()
    If IsNull(MRSelectMonth) Or Len(MRSelectMonth) <= 0 Then
        MsgBox "You must select a month above.", vbOKOnly, "Error!"
        Exit Sub
    End If
    ReportFY = CLng(MRFY): ReportMonth = MRSelectMonth
    Call BriefDelay: DoCmd.OpenReport "rptVICTORIABILLS", acViewPreview, , "FY=" & MRFY, , MRFY & " " & MRSelectMonth: Call LoopUntilClosed("rptJOHNMBILLS", 3)
    Call DropTempTables
End Sub

Private Sub JulieBillings_Click()
    If IsNull(MRSelectMonth) Or Len(MRSelectMonth) <= 0 Then
        MsgBox "You must select a month above.", vbOKOnly, "Error!"
        Exit Sub
    End If
    ReportFY = CLng(MRFY): ReportMonth = MRSelectMonth
    Call BriefDelay: DoCmd.OpenReport "rptJULIEBILLS", acViewPreview, , "FY=" & MRFY, , MRFY & " " & MRSelectMonth: Call LoopUntilClosed("rptJULIEBILLS", 3)
    Call DropTempTables
End Sub

Private Sub MSAMiscReport_Click()
    If IsNull(MRSelectMonth) Or Len(MRSelectMonth) <= 0 Then
        MsgBox "You must select a month above.", vbOKOnly, "Error!"
        Exit Sub
    End If
    DoCmd.SetWarnings False
    DoCmd.OpenQuery "qryMSAMISCTEMP"
    DoCmd.SetWarnings True
    Call BriefDelay: DoCmd.OpenReport "rptMSAMISC", acViewPreview, , "FY=" & MRFY, , MRFY & " " & MRNumDays & " " & MRSelectMonth: Call LoopUntilClosed("rptMSAMISC", 3)
End Sub

Private Sub REPPAYEEFOLLOWUPSelectAllClients_Click()
    ReportName = "ltrREPPAYEEFOLLOWUPTillIsNotRepPayee"
    Call BriefDelay
    Call ExecReport(ReportName)
End Sub

Private Sub REPPAYEEFOLLOWUPSelectCLient_AfterUpdate()
    ReportName = "ltrREPPAYEEFOLLOWUPSelectedClient"
    SelectedClientIndexedName = DLookup("IndexedName", "AllClients", "ReverseDisplayName = """ & REPPAYEEFOLLOWUPSelectCLient & """")
    SelectedFirstName = DLookup("FirstName", "AllClients", "ReverseDisplayName = """ & REPPAYEEFOLLOWUPSelectCLient & """")
    SelectedLastName = DLookup("LastName", "AllClients", "ReverseDisplayName = """ & REPPAYEEFOLLOWUPSelectCLient & """")
    If Len(SelectedClientIndexedName) > 0 Then Call ExecReport(ReportName)
    REPPAYEEFOLLOWUPSelectCLient = Null: Me.Refresh: Me.Requery: Me.Repaint
End Sub

Private Sub RPYGo_Click()
On Error GoTo ShowMeError
    Call DropTempTables
    Call UpdateProgressMessages("Building table of Rep Payees...please wait.", True)
    TILLDataBase.Execute "SELECT tblPeople.IndexedName, tblPeople.LastName, tblPeople.FirstName, tblPeople.MiddleInitial, tblPeopleClientsDemographics.RepresentativePayee, " & _
        "tblPeopleClientsDemographics.ActiveDayServices, tblPeopleClientsDayServices.Inactive AS InactDAY, [tblPeopleClientsDayServices]![CityTown] & ' - ' & [tblPeopleClientsDayServices]![LocationName] AS DayProgramLocation, tblPeopleClientsDemographics.ActiveVocationalServices, tblPeopleClientsVocationalServices.Inactive AS InactVOC, [tblPeopleClientsVocationalServices]![CityTown] & ' - ' & [tblPeopleClientsVocationalServices]![Location] AS VocationalProgramLocation, " & _
        "IIf([ActiveDayServices] And Not [InactDay],[DayProgramLocation],IIf([ActiveVocationalServices] And Not [InactVOC],[VocationalProgramLocation],Null)) AS DayLocationForDisplay, " & _
        "tblPeopleClientsDemographics.ActiveResidentialServices, tblPeopleClientsResidentialServices.Inactive AS InactRES, [tblPeopleClientsResidentialServices]![CityTown] & ' - ' & [tblPeopleClientsResidentialServices]![Location] AS ResidentialProgramLocation, tblPeopleClientsDemographics.ActiveCLO, tblPeopleClientsCLOServices.Inactive AS InactCLO, [tblPeopleClientsCLOServices]![CityTown] & ' - ' & [tblPeopleClientsCLOServices]![Location] AS CLOProgramLocation, " & _
        "IIf([ActiveResidentialServices] And Not [InactRES],[ResidentialProgramLocation],IIf([ActiveCLO] And Not [InactCLO],[CLOProgramLocation],Null)) AS ResLocationForDisplay, " & _
        "tblPeopleClientsDemographics.ActiveSharedLiving, tblPeopleClientsSharedLivingServices.Inactive AS InactSL, tblPeopleClientsDemographics.ActiveAFC, tblPeopleClientsAFCServices.Inactive AS InactAFC, tblPeopleClientsDemographics.ActiveIndivSupport, tblPeopleClientsIndividualSupportServices.Inactive AS InactISS, tblPeopleClientsDemographics.ActiveAutismServices, tblPeopleClientsAutismServices.Inactive AS InactAUT, tblPeopleClientsDemographics.ActivePCA, tblPeopleClientsPCAServices.Inactive AS InactPCA, tblPeopleClientsDemographics.ActiveSpringboard, tblPeopleClientsSpringboardServices.Inactive AS InactSPR, tblPeopleClientsDemographics.ActiveTRASE, tblPeopleClientsTRASEServices.Inactive AS InactTRASE, tblPeopleClientsDemographics.ActiveCommunityConnections, tblPeopleClientsCommunityConnectionsServices.Inactive AS InactCC, " & _
        "tblPeopleClientsVendors.ResVendorLocation, tblPeopleClientsVendors.DayVendorLocation, IIf(InStr(1,[RepresentativePayee],'TILL'),'TILL','Not TILL') AS TILLRepPayee, tblPeopleClientsDemographics.RepPayeeAddress, tblPeopleClientsDemographics.RepPayeeCity, tblPeopleClientsDemographics.RepPayeeState, tblPeopleClientsDemographics.RepPayeeZIP, tblPeopleClientsDemographics.RepPayeeReportDate, tblPeopleClientsDemographics.RepPayeePhone, tblPeopleClientsDemographics.DateOfBirth, tblPeopleClientsDemographics.SocialSecurityNumber, tblPeopleClientsDemographics.LegalStatus INTO temptbl0 " & _
        "FROM (((((((((((((tblPeople INNER JOIN tblPeopleClientsDemographics ON tblPeople.IndexedName = tblPeopleClientsDemographics.IndexedName) " & _
        "INNER JOIN tblPeopleClientsVendors ON tblPeople.IndexedName = tblPeopleClientsVendors.IndexedName) " & _
        "LEFT JOIN tblPeopleClientsCLOServices ON tblPeopleClientsDemographics.IndexedName = tblPeopleClientsCLOServices.IndexedName) " & _
        "LEFT JOIN tblPeopleClientsDayServices ON tblPeopleClientsDemographics.IndexedName = tblPeopleClientsDayServices.IndexedName) " & _
        "LEFT JOIN tblPeopleClientsResidentialServices ON tblPeopleClientsDemographics.IndexedName = tblPeopleClientsResidentialServices.IndexedName) " & _
        "LEFT JOIN tblPeopleClientsSharedLivingServices ON tblPeopleClientsDemographics.IndexedName = tblPeopleClientsSharedLivingServices.IndexedName) " & _
        "LEFT JOIN tblPeopleClientsAFCServices ON tblPeopleClientsDemographics.IndexedName = tblPeopleClientsAFCServices.IndexedName) " & _
        "LEFT JOIN tblPeopleClientsVocationalServices ON tblPeopleClientsDemographics.IndexedName = tblPeopleClientsVocationalServices.IndexedName) " & _
        "LEFT JOIN tblPeopleClientsIndividualSupportServices ON tblPeopleClientsDemographics.IndexedName = tblPeopleClientsIndividualSupportServices.IndexedName) " & _
        "LEFT JOIN tblPeopleClientsAutismServices ON tblPeopleClientsDemographics.IndexedName = tblPeopleClientsAutismServices.IndexedName) " & _
        "LEFT JOIN tblPeopleClientsPCAServices ON tblPeopleClientsDemographics.IndexedName = tblPeopleClientsPCAServices.IndexedName) " & _
        "LEFT JOIN tblPeopleClientsSpringboardServices ON tblPeopleClientsDemographics.IndexedName = tblPeopleClientsSpringboardServices.IndexedName) " & _
        "LEFT JOIN tblPeopleClientsTRASEServices ON tblPeopleClientsDemographics.IndexedName = tblPeopleClientsTRASEServices.IndexedName) " & _
        "LEFT JOIN tblPeopleClientsCommunityConnectionsServices ON tblPeopleClientsDemographics.IndexedName = tblPeopleClientsCommunityConnectionsServices.IndexedName " & _
        "WHERE tblPeopleClientsDemographics.RepresentativePayee Is Not Null AND tblPeople.IsDeceased=False AND " & _
        "((tblPeopleClientsDemographics.ActiveDayServices=True AND tblPeopleClientsDayServices.Inactive=False) OR (tblPeopleClientsDemographics.ActiveVocationalServices=True AND tblPeopleClientsVocationalServices.Inactive=False) OR (tblPeopleClientsDemographics.ActiveResidentialServices=True AND tblPeopleClientsResidentialServices.Inactive=False) OR (tblPeopleClientsDemographics.ActiveSharedLiving=True AND tblPeopleClientsSharedLivingServices.Inactive=False) OR " & _
        " (tblPeopleClientsDemographics.ActiveAFC=True AND tblPeopleClientsAFCServices.Inactive=False) OR (tblPeopleClientsDemographics.ActiveIndivSupport=True AND tblPeopleClientsIndividualSupportServices.Inactive=False) OR (tblPeopleClientsDemographics.ActiveCLO=True AND tblPeopleClientsCLOServices.Inactive=False)) ORDER BY tblPeople.IndexedName;"
    SysCmdResult = SysCmd(4, "Preparing report.")
    Call UpdateProgressMessages("Preparing report.", , True)
    Call BriefDelay
    DoCmd.OpenReport "rptREPPAYEESbyClient", acViewPreview, , , , "ALL CLI"
    Exit Sub
ShowMeError:
    MsgBox "Error # " & Str(Err.Number) & " was generated by " & Me.Name & Chr(13) & Err.Description, vbOKOnly, "Error", Err.HelpFile, Err.HelpContext
End Sub

Private Sub Form_Load()
    CFCMSortBy.DefaultValue = "<Select sort criteria>"
    EWDLSelectClient = Null: AutoDepositAuth = Null: GuardSelectClient = Null: IntroSelectClient = Null
    IndivFinEnterEffectiveMonth.Visible = True: IndivFinEnterEffectiveMonth = Null: IndivFinSelectRepPayee.Visible = False
    IndivFinSelectRepPayee = Null: IndivFinAllOrSingle.Visible = False: IndivFinAllOrSingle = Null: IndivFinDollarAmounts.Visible = False
    IndivFinDollarAmounts = Null: IndivFinSelectClient.Visible = False: IndivFinSelectClient = Null: MAFY = Null
    MedSelectClient = Null: MedEnterEffectiveMonth = Null: MedSelectedLetterType = Null: MEDIASelectClient = Null
    MEDIAEnterEffectiveMonth = Null: MRMask.Visible = True: NHFY = Null: RCBSelectClient = Null: REPPAYEEFEESelectClient = Null
    SECT8SelectCLient = Null: SSASelectClient = Null: SUPREPSelectClient = Null: SUPREPEnterEffectiveMonth = Null
    TRVLSelectClient = Null: TRVLEnterEffectiveMonth = Null: ErrorOccurred = False
    SigOnFileClient.Visible = False
    RepPayeeFee = DLookup("ParameterValue", "appParameters", "ParameterName='RepPayeeFee'")
    
    Select Case Month(Date)
        Case 1, 2, 3, 4, 5, 6: MRFY = Year(Date)
        Case Else:             MRFY = Year(Date) + 1
    End Select
    
    MRSelectMonth.RowSource = "April " & Format(MRFY - 1, "####") & "; May " & Format(MRFY - 1, "####") & "; June " & Format(MRFY - 1, "####") & "; July " & Format(MRFY - 1, "####") & _
        "; August " & Format(MRFY - 1, "####") & "; September " & Format(MRFY - 1, "####") & _
        "; October " & Format(MRFY - 1, "####") & "; November " & Format(MRFY - 1, "####") & _
        "; December " & Format(MRFY - 1, "####") & "; January " & Format(MRFY, "####") & _
        "; February " & Format(MRFY, "####") & "; March " & Format(MRFY, "####") & _
        "; April " & Format(MRFY, "####") & "; May " & Format(MRFY, "####") & _
        "; June " & Format(MRFY, "####")
End Sub

Private Sub CONSUMMARY_Click()
    DoCmd.OpenForm "frmRptCONSUMMARY"
End Sub

Private Sub AMENDMENTS_Click()
    DoCmd.OpenForm "frmRptAMENDMENTS"
End Sub

Private Sub CFCMETHOD_Click()
    DoCmd.OpenForm "frmRptCFC"
End Sub

Private Sub CLIENTCC_Click()
On Error GoTo ShowMeError
    Dim ExportFileName As String
    
    Call DropTempTables
    TILLDataBase.Execute "CREATE TABLE temptbl0 (IndexedName CHAR(160), ClientService CHAR(25), LastName CHAR(25), FirstName CHAR(25), MiddleInitial CHAR(1), CityTown CHAR(25), LocationName CHAR(30), CostCenter CHAR(4));", dbSeeChanges: Call BriefDelay
    TILLDataBase.Execute "ALTER TABLE temptbl0 ADD CONSTRAINT temptblconstraint PRIMARY KEY (IndexedName,ClientService);", dbSeeChanges: Call BriefDelay
    '   Residential.
    TILLDataBase.Execute "INSERT INTO temptbl0 ( IndexedName, LastName, FirstName, MiddleInitial, ClientService, CityTown, LocationName, CostCenter ) " & _
        "SELECT tblPeople.IndexedName, tblPeople.LastName, tblPeople.FirstName, tblPeople.MiddleInitial, 'Residential' AS ClientService, Null AS CityTown, Null AS LocationName, Null AS CostCenter " & _
        "FROM tblPeople INNER JOIN tblPeopleClientsResidentialServices ON tblPeople.IndexedName = tblPeopleClientsResidentialServices.IndexedName " & _
        "WHERE tblPeopleClientsResidentialServices.Inactive=False AND tblPeople.IsDeceased=False AND tblPeople.IsClientRes=True;", dbSeeChanges: Call BriefDelay
    TILLDataBase.Execute "UPDATE (temptbl0 INNER JOIN qrytblPeopleClientsResidentialServices ON temptbl0.IndexedName = qrytblPeopleClientsResidentialServices.IndexedName) INNER JOIN qrytblLocations ON (qrytblPeopleClientsResidentialServices.Location = qrytblLocations.LocationName) AND (qrytblPeopleClientsResidentialServices.CityTown = qrytblLocations.CityTown) SET temptbl0.CityTown = [qrytblPeopleClientsResidentialServices]![CityTown], temptbl0.LocationName = [qrytblPeopleClientsResidentialServices]![Location], temptbl0.CostCenter = [qrytblLocations].[CostCenter] " & _
        "WHERE temptbl0.ClientService='Residential';", dbSeeChanges: Call BriefDelay
    '   Day.
    TILLDataBase.Execute "INSERT INTO temptbl0 ( IndexedName, LastName, FirstName, MiddleInitial, ClientService, CityTown, LocationName, CostCenter ) " & _
        "SELECT tblPeople.IndexedName, tblPeople.LastName, tblPeople.FirstName, tblPeople.MiddleInitial, 'Day Services' AS ClientService, Null AS CityTown, Null AS LocationName, Null AS CostCenter " & _
        "FROM tblPeople INNER JOIN tblPeopleClientsDayServices ON tblPeople.IndexedName = tblPeopleClientsDayServices.IndexedName " & _
        "WHERE tblPeopleClientsDayServices.Inactive=False AND tblPeople.IsDeceased=False AND tblPeople.IsClientDay=True;", dbSeeChanges: Call BriefDelay
    TILLDataBase.Execute "UPDATE (temptbl0 INNER JOIN qrytblPeopleClientsDayServices ON temptbl0.IndexedName = qrytblPeopleClientsDayServices.IndexedName) INNER JOIN qrytblLocations ON (qrytblPeopleClientsDayServices.LocationName = qrytblLocations.LocationName) AND (qrytblPeopleClientsDayServices.CityTown = qrytblLocations.CityTown) SET temptbl0.CityTown = [qrytblPeopleClientsDayServices]![CityTown], temptbl0.LocationName = [qrytblPeopleClientsDayServices]![LocationName], temptbl0.CostCenter = [qrytblLocations].[CostCenter] " & _
        "WHERE temptbl0.ClientService='Day Services';", dbSeeChanges: Call BriefDelay
    '   Vocational.
    TILLDataBase.Execute "INSERT INTO temptbl0 ( IndexedName, LastName, FirstName, MiddleInitial, ClientService, CityTown, LocationName, CostCenter ) " & _
        "SELECT tblPeople.IndexedName, tblPeople.LastName, tblPeople.FirstName, tblPeople.MiddleInitial, 'Vocational' AS ClientService, Null AS CityTown, Null AS LocationName, Null AS CostCenter " & _
        "FROM tblPeople INNER JOIN tblPeopleClientsVocationalServices ON tblPeople.IndexedName = tblPeopleClientsVocationalServices.IndexedName " & _
        "WHERE tblPeopleClientsVocationalServices.Inactive=False AND tblPeople.IsDeceased=False AND tblPeople.IsClientVocat=True;", dbSeeChanges: Call BriefDelay
    TILLDataBase.Execute "UPDATE (temptbl0 INNER JOIN qrytblPeopleClientsVocationalServices ON temptbl0.IndexedName = qrytblPeopleClientsVocationalServices.IndexedName) INNER JOIN qrytblLocations ON (qrytblPeopleClientsVocationalServices.Location = qrytblLocations.LocationName) AND (qrytblPeopleClientsVocationalServices.CityTown = qrytblLocations.CityTown) SET temptbl0.CityTown = [qrytblPeopleClientsVocationalServices].[CityTown], temptbl0.LocationName = [qrytblPeopleClientsVocationalServices].[Location], temptbl0.CostCenter = [qrytblLocations].[CostCenter] " & _
        "WHERE temptbl0.ClientService='Vocational';", dbSeeChanges: Call BriefDelay
    '   Shared Living.
    TILLDataBase.Execute "INSERT INTO temptbl0 ( IndexedName, LastName, FirstName, MiddleInitial, ClientService, CityTown, LocationName, CostCenter ) " & _
        "SELECT tblPeople.IndexedName, tblPeople.LastName, tblPeople.FirstName, tblPeople.MiddleInitial, 'Shared Living' AS ClientService, Null AS CityTown, Null AS LocationName, Null AS CostCenter " & _
        "FROM tblPeople INNER JOIN tblPeopleClientsSharedLivingServices ON tblPeople.IndexedName = tblPeopleClientsSharedLivingServices.IndexedName " & _
        "WHERE tblPeopleClientsSharedLivingServices.Inactive=False AND tblPeople.IsDeceased=False AND tblPeople.IsClientSharedLiving=True;", dbSeeChanges: Call BriefDelay
    TILLDataBase.Execute "UPDATE (temptbl0 INNER JOIN qrytblPeopleClientsSharedLivingServices ON temptbl0.IndexedName = qrytblPeopleClientsSharedLivingServices.IndexedName) INNER JOIN tblContractsBillingBook ON qrytblPeopleClientsSharedLivingServices.ContractNumber = tblContractsBillingBook.ContractID SET temptbl0.CostCenter = [tblContractsBillingBook].[CostCenter] " & _
        "WHERE temptbl0.ClientService='Shared Living';", dbSeeChanges: Call BriefDelay
    '   Adult Foster Care.
    TILLDataBase.Execute "INSERT INTO temptbl0 ( IndexedName, LastName, FirstName, MiddleInitial, ClientService, CityTown, LocationName, CostCenter ) " & _
        "SELECT tblPeople.IndexedName, tblPeople.LastName, tblPeople.FirstName, tblPeople.MiddleInitial, 'Adult Foster Care' AS ClientService, Null AS CityTown, Null AS LocationName, Null AS CostCenter " & _
        "FROM tblPeople INNER JOIN tblPeopleClientsAFCServices ON tblPeople.IndexedName = tblPeopleClientsAFCServices.IndexedName " & _
        "WHERE tblPeopleClientsAFCServices.Inactive=False AND tblPeople.IsDeceased=False AND tblPeople.IsClientAFC=True;", dbSeeChanges: Call BriefDelay
    TILLDataBase.Execute "UPDATE (temptbl0 INNER JOIN qrytblPeopleClientsAFCServices ON temptbl0.IndexedName = qrytblPeopleClientsAFCServices.IndexedName) INNER JOIN tblContractsBillingBook ON qrytblPeopleClientsAFCServices.ContractNumber = tblContractsBillingBook.ContractID SET temptbl0.CostCenter = [tblContractsBillingBook].[CostCenter] " & _
        "WHERE temptbl0.ClientService='Adult Foster Care';", dbSeeChanges: Call BriefDelay
    '   CLO.
    TILLDataBase.Execute "INSERT INTO temptbl0 ( IndexedName, LastName, FirstName, MiddleInitial, ClientService, CityTown, LocationName, CostCenter ) " & _
        "SELECT tblPeople.IndexedName, tblPeople.LastName, tblPeople.FirstName, tblPeople.MiddleInitial, 'CLO' AS ClientService, Null AS CityTown, Null AS LocationName, Null AS CostCenter " & _
        "FROM tblPeople INNER JOIN tblPeopleClientsCLOServices ON tblPeople.IndexedName = tblPeopleClientsCLOServices.IndexedName " & _
        "WHERE tblPeopleClientsCLOServices.Inactive=False AND tblPeople.IsDeceased=False AND tblPeople.IsCilentCLO=True;", dbSeeChanges: Call BriefDelay
    TILLDataBase.Execute "UPDATE (temptbl0 INNER JOIN qrytblPeopleClientsCLOServices ON temptbl0.IndexedName = qrytblPeopleClientsCLOServices.IndexedName) INNER JOIN qrytblLocations ON (qrytblPeopleClientsCLOServices.Location = qrytblLocations.LocationName) AND (qrytblPeopleClientsCLOServices.CityTown = qrytblLocations.CityTown) SET temptbl0.CityTown = [qrytblPeopleClientsCLOServices].[CityTown], temptbl0.LocationName = [qrytblPeopleClientsCLOServices].[Location], temptbl0.CostCenter = [qrytblLocations].[CostCenter] " & _
        "WHERE temptbl0.ClientService='CLO';", dbSeeChanges: Call BriefDelay
    '   Individual Support.
    TILLDataBase.Execute "INSERT INTO temptbl0 ( IndexedName, LastName, FirstName, MiddleInitial, ClientService, CityTown, LocationName, CostCenter ) " & _
        "SELECT tblPeople.IndexedName, tblPeople.LastName, tblPeople.FirstName, tblPeople.MiddleInitial, 'Individual Support' AS ClientService, Null AS CityTown, Null AS LocationName, Null AS CostCenter " & _
        "FROM tblPeople INNER JOIN tblPeopleClientsIndividualSupportServices ON tblPeople.IndexedName = tblPeopleClientsIndividualSupportServices.IndexedName " & _
        "WHERE tblPeopleClientsIndividualSupportServices.Inactive=False AND tblPeople.IsDeceased=False AND tblPeople.IsClientIndiv=True;", dbSeeChanges: Call BriefDelay
    TILLDataBase.Execute "UPDATE (temptbl0 INNER JOIN qrytblPeopleClientsIndividualSupportServices ON temptbl0.IndexedName = qrytblPeopleClientsIndividualSupportServices.IndexedName) INNER JOIN tblContractsBillingBook ON qrytblPeopleClientsIndividualSupportServices.ContractNumber = tblContractsBillingBook.ContractID SET temptbl0.CostCenter = [tblContractsBillingBook].[CostCenter] " & _
        "WHERE temptbl0.ClientService='Individual Support';", dbSeeChanges: Call BriefDelay
    '   Autism.
    TILLDataBase.Execute "INSERT INTO temptbl0 ( IndexedName, LastName, FirstName, MiddleInitial, ClientService, CityTown, LocationName, CostCenter ) " & _
        "SELECT tblPeople.IndexedName, tblPeople.LastName, tblPeople.FirstName, tblPeople.MiddleInitial, 'Autism' AS ClientService, Null AS CityTown, Null AS LocationName, Null AS CostCenter " & _
        "FROM tblPeople INNER JOIN tblPeopleClientsAutismServices ON tblPeople.IndexedName = tblPeopleClientsAutismServices.IndexedName " & _
        "WHERE tblPeopleClientsAutismServices.Inactive=False AND tblPeople.IsDeceased=False AND tblPeople.IsClientAutism=True;", dbSeeChanges: Call BriefDelay
    '   Springboard.
    TILLDataBase.Execute "INSERT INTO temptbl0 ( IndexedName, LastName, FirstName, MiddleInitial, ClientService, CityTown, LocationName, CostCenter ) " & _
        "SELECT tblPeople.IndexedName, tblPeople.LastName, tblPeople.FirstName, tblPeople.MiddleInitial, 'Springboard' AS ClientService, Null AS CityTown, Null AS LocationName, Null AS CostCenter " & _
        "FROM tblPeople INNER JOIN tblPeopleClientsSpringboardServices ON tblPeople.IndexedName = tblPeopleClientsSpringboardServices.IndexedName " & _
        "WHERE tblPeopleClientsSpringboardServices.Inactive=False AND tblPeople.IsDeceased=False AND tblPeople.IsClientSpring=True;", dbSeeChanges: Call BriefDelay
    
    If MsgBox("Do you want a spreadsheet for this?", vbYesNo, "Spreadsheet?") = vbYes Then
        ExportFileName = Application.CurrentProject.Path & "\" & "TILLDB-Export-ClientCostCenters-" & Format(Date, "yyyymmdd") & ".xls"
        If IsFileOpen(ExportFileName) Then
            If MsgBox(ExportFileName & " is already open.  Please close it and click OK to continue or Cancel to abort.", vbOKCancel, "ERROR!") = vbCancel Then
                MsgBox "Export aborted.", vbOKOnly, "Aborted"
                Exit Sub
            End If
        End If
        If Dir(ExportFileName) <> "" Then Kill ExportFileName
        DoCmd.TransferSpreadsheet acExport, acSpreadsheetTypeExcel9, "temptbl0", ExportFileName
        MsgBox "The requested information has been exported to " & ExportFileName & "." & vbCrLf & vbCrLf & "This export may contain information that is protected under HIPAA and other privacy laws.  This export must be securely stored at all times and must be deleted when no longer being used.", vbOKOnly, "Export Complete"
    Else
        Call BriefDelay: Call ExecReport("rptCLIENTCC")
    End If
    Call DropTempTables
    Exit Sub
ShowMeError:
    MsgBox "Error # " & Str(Err.Number) & " was generated by " & Me.Name & Chr(13) & Err.Description, vbOKOnly, "Error", Err.HelpFile, Err.HelpContext
End Sub

Private Sub EWDLS_Click()
    Call BriefDelay: DoCmd.OpenForm "frmRptLtrEWDLS"
End Sub

Private Sub EWDLSTILL_Click()
    Call BriefDelay: DoCmd.OpenForm "frmRptLtrEWDLSTILL"
End Sub

Private Sub EWDLSelectClient_AfterUpdate()
    Dim WhichLetter As String, RepPayee As Variant

    Call BuildAllClients(False)
    SelectedClientIndexedName = DLookup("IndexedName", "AllClients", "ReverseDisplayName = """ & EWDLSelectClient & """")
    RepPayee = DLookup("RepresentativePayee", "tblPeopleClientsDemographics", "IndexedName=""" & SelectedClientIndexedName & """")
    If IsNull(RepPayee) Or Len(RepPayee) <= 0 Then
        MsgBox "Individual has no Representative Payee identified.  Please correct this.", vbOKOnly, "Error"
        EWDLSelectClient = Null
        Exit Sub
    End If
    
    If Len(SelectedClientIndexedName) > 0 Then
        If Left(RepPayee, 4) = "TILL" Then WhichLetter = "ltrEWDLSTILL" Else WhichLetter = "ltrEWDLS"
        Call BriefDelay: Call ExecReport(WhichLetter)
    End If
    
    EWDLSelectClient = Null: Me.Refresh: Me.Requery: Me.Repaint
End Sub

Private Sub IndivFinDoReport()
On Error GoTo ShowMeError
    Dim RepPayee As Variant, WhichLetter As String, Criteria1 As Variant, Criteria2 As Variant

    ' Load data.
    TILLDataBase.Execute "DELETE [~IndivFinances].* FROM [~IndivFinances];"
    TILLDataBase.Execute "INSERT INTO [~IndivFinances] ( IndexedName, LastName, FirstName, MiddleInitial, WorkingABI, RepresentativePayee, RepPayeeIsClient, Service, Program, LegalName, SSAMonthlyAmount, SSIMonthlyAmount, SSPMonthlyAmount, Pension, OtherMonthlyAmount, WdlPercent, TotalBenefits, DiscountedBenefits, RemainingBenefits, SubtotalA, MonthlyEarnedIncome, Less65Dollars, SubtotalB, TotalMonthlyCharge ) " & _
        "SELECT tblPeople.IndexedName, tblPeople.LastName, tblPeople.FirstName, tblPeople.MiddleInitial, False AS WorkingABI, " & _
        "tblPeopleClientsDemographics.RepresentativePayee, tblPeopleClientsDemographics.RepPayeeIsClient, 'CLO' AS Service, " & _
        "[tblPeopleClientsCLOServices]![CityTown] & '-' & [tblPeopleClientsCLOServices]![Location] AS Program, " & _
        "[tblPeople]![FirstName] & ' ' & [tblPeople]![MiddleInitial] & ' ' & [tblPeople]![LastName] AS LegalName, " & _
        "IIf(IsNull([tblPeopleClientsDemographics]![SSAMonthlyAmount]),0,[tblPeopleClientsDemographics]![SSAMonthlyAmount]) AS SSAMonthlyAmount, " & _
        "IIf(IsNull([tblPeopleClientsDemographics]![SSIMonthlyAmount]),0,[tblPeopleClientsDemographics]![SSIMonthlyAmount]) AS SSIMonthlyAmount, " & _
        "IIf(IsNull([tblPeopleClientsDemographics]![SSPMonthlyAmount]),0,[tblPeopleClientsDemographics]![SSPMonthlyAmount]) AS SSPMonthlyAmount, " & _
        "IIf(IsNull([tblPeopleClientsDemographics]![Pension]),0,[tblPeopleClientsDemographics]![Pension]) AS Pension, " & _
        "IIf(IsNull([tblPeopleClientsDemographics]![OtherMonthlyAmount]),0,[tblPeopleClientsDemographics]![OtherMonthlyAmount]) AS OtherMonthlyAmount, tblPeopleClientsDemographics.WdlPercent, " & _
        "[SSAMonthlyAmount]+[SSIMonthlyAmount]+[SSPMonthlyAmount]+[OtherMonthlyAmount]+[Pension] AS TotalBenefits, " & _
        "[TotalBenefits]*[WdlPercent]/100 AS DiscountedBenefits, " & _
        "IIf([TotalBenefits]-[DiscountedBenefits]>=200,[TotalBenefits]-[DiscountedBenefits],200) AS RemainingBenefits, " & _
        "[DiscountedBenefits] AS SubtotalA, Null AS MonthlyEarnedIncome, Null AS Less65Dollars, Null AS SubtotalB, " & _
        "IIf(IsNull([SubtotalB]),[SubtotalA],[SubtotalA]+[SubtotalB]) AS TotalMonthlyCharge " & _
        "FROM (tblPeople INNER JOIN tblPeopleClientsDemographics ON tblPeople.IndexedName = tblPeopleClientsDemographics.IndexedName) LEFT JOIN tblPeopleClientsCLOServices ON tblPeopleClientsDemographics.IndexedName = tblPeopleClientsCLOServices.IndexedName " & _
        "WHERE (((tblPeopleClientsDemographics.ActiveCLO) = True) And ((tblPeopleClientsCLOServices.Inactive) = False) And ((tblPeople.IsDeceased) = False));", dbSeeChanges: Call BriefDelay
    TILLDataBase.Execute "INSERT INTO [~IndivFinances] ( IndexedName, LastName, FirstName, MiddleInitial, WorkingABI, RepresentativePayee, RepPayeeIsClient, Service, Program, LegalName, SSAMonthlyAmount, SSIMonthlyAmount, SSPMonthlyAmount, Pension, OtherMonthlyAmount, WdlPercent, TotalBenefits, DiscountedBenefits, RemainingBenefits, SubtotalA, MonthlyEarnedIncome, Less65Dollars, SubtotalB, TotalMonthlyCharge ) " & _
        "SELECT tblPeople.IndexedName, tblPeople.LastName, tblPeople.FirstName, tblPeople.MiddleInitial, IIf([tblPeopleClientsDemographics]![ActiveResidentialServices],[tblLocations]![ABI],False) AS WorkingABI, " & _
        "tblPeopleClientsDemographics.RepresentativePayee, tblPeopleClientsDemographics.RepPayeeIsClient, 'RES' AS Service, " & _
        "[tblPeopleClientsResidentialServices]![CityTown] & '-' & [tblPeopleClientsResidentialServices]![Location] AS Program, " & _
        "[tblPeople]![FirstName] & ' ' & [tblPeople]![MiddleInitial] & ' ' & [tblPeople]![LastName] AS LegalName, " & _
        "IIf(IsNull([tblPeopleClientsDemographics]![SSAMonthlyAmount]),0,[tblPeopleClientsDemographics]![SSAMonthlyAmount]) AS SSAMonthlyAmount, " & _
        "IIf(IsNull([tblPeopleClientsDemographics]![SSIMonthlyAmount]),0,[tblPeopleClientsDemographics]![SSIMonthlyAmount]) AS SSIMonthlyAmount, " & _
        "IIf(IsNull([tblPeopleClientsDemographics]![SSPMonthlyAmount]),0,[tblPeopleClientsDemographics]![SSPMonthlyAmount]) AS SSPMonthlyAmount, " & _
        "IIf(IsNull([tblPeopleClientsDemographics]![Pension]),0,[tblPeopleClientsDemographics]![Pension]) AS Pension, " & _
        "IIf(IsNull([tblPeopleClientsDemographics]![OtherMonthlyAmount]),0,[tblPeopleClientsDemographics]![OtherMonthlyAmount]) AS OtherMonthlyAmount, tblPeopleClientsDemographics.WdlPercent, " & _
        "[SSAMonthlyAmount]+[SSIMonthlyAmount]+[SSPMonthlyAmount]+[OtherMonthlyAmount]+[Pension] AS TotalBenefits, " & _
        "[TotalBenefits]*[WdlPercent]/100 AS DiscountedBenefits, " & _
        "IIf([TotalBenefits]-[DiscountedBenefits]>=200,[TotalBenefits]-[DiscountedBenefits],200) AS RemainingBenefits, " & _
        "IIf([ABI],[TotalBenefits]-[RemainingBenefits],[DiscountedBenefits]) AS SubtotalA, Null AS MonthlyEarnedIncome, Null AS Less65Dollars, Null AS SubtotalB, " & _
        "IIf(IsNull([SubtotalB]),[SubtotalA],[SubtotalA]+[SubtotalB]) AS TotalMonthlyCharge " & _
        "FROM ((tblPeople INNER JOIN tblPeopleClientsDemographics ON tblPeople.IndexedName = tblPeopleClientsDemographics.IndexedName) LEFT JOIN tblPeopleClientsResidentialServices ON tblPeopleClientsDemographics.IndexedName = tblPeopleClientsResidentialServices.IndexedName) LEFT JOIN tblLocations ON (tblPeopleClientsResidentialServices.Location = tblLocations.LocationName) AND (tblPeopleClientsResidentialServices.CityTown = tblLocations.CityTown) " & _
        "WHERE (((tblPeopleClientsDemographics.ActiveResidentialServices) = True) And ((tblPeopleClientsResidentialServices.Inactive) = False) And ((tblPeople.IsDeceased) = False))", dbSeeChanges: Call BriefDelay
    TILLDataBase.Execute "INSERT INTO [~IndivFinances] ( IndexedName, LastName, FirstName, MiddleInitial, WorkingABI, RepresentativePayee, RepPayeeIsClient, Service, Program, LegalName, SSAMonthlyAmount, SSIMonthlyAmount, SSPMonthlyAmount, Pension, OtherMonthlyAmount, WdlPercent, TotalBenefits, DiscountedBenefits, RemainingBenefits, SubtotalA, MonthlyEarnedIncome, Less65Dollars, SubtotalB, TotalMonthlyCharge ) " & _
        "SELECT tblPeople.IndexedName, tblPeople.LastName, tblPeople.FirstName, tblPeople.MiddleInitial, False AS WorkingABI, " & _
        "tblPeopleClientsDemographics.RepresentativePayee, tblPeopleClientsDemographics.RepPayeeIsClient, 'ISS' AS Service, " & _
        "Null AS Program, " & _
        "[tblPeople]![FirstName] & ' ' & [tblPeople]![MiddleInitial] & ' ' & [tblPeople]![LastName] AS LegalName, " & _
        "IIf(IsNull([tblPeopleClientsDemographics]![SSAMonthlyAmount]),0,[tblPeopleClientsDemographics]![SSAMonthlyAmount]) AS SSAMonthlyAmount, " & _
        "IIf(IsNull([tblPeopleClientsDemographics]![SSIMonthlyAmount]),0,[tblPeopleClientsDemographics]![SSIMonthlyAmount]) AS SSIMonthlyAmount, " & _
        "IIf(IsNull([tblPeopleClientsDemographics]![SSPMonthlyAmount]),0,[tblPeopleClientsDemographics]![SSPMonthlyAmount]) AS SSPMonthlyAmount, " & _
        "IIf(IsNull([tblPeopleClientsDemographics]![Pension]),0,[tblPeopleClientsDemographics]![Pension]) AS Pension, " & _
        "IIf(IsNull([tblPeopleClientsDemographics]![OtherMonthlyAmount]),0,[tblPeopleClientsDemographics]![OtherMonthlyAmount]) AS OtherMonthlyAmount, tblPeopleClientsDemographics.WdlPercent, " & _
        "[SSAMonthlyAmount]+[SSIMonthlyAmount]+[SSPMonthlyAmount]+[OtherMonthlyAmount]+[Pension] AS TotalBenefits, " & _
        "[TotalBenefits]*[WdlPercent]/100 AS DiscountedBenefits, " & _
        "IIf([TotalBenefits]-[DiscountedBenefits]>=200,[TotalBenefits]-[DiscountedBenefits],200) AS RemainingBenefits, " & _
        "[DiscountedBenefits] AS SubtotalA, Null AS MonthlyEarnedIncome, Null AS Less65Dollars, Null AS SubtotalB, " & _
        "IIf(IsNull([SubtotalB]),[SubtotalA],[SubtotalA]+[SubtotalB]) AS TotalMonthlyCharge " & _
        "FROM (tblPeople INNER JOIN tblPeopleClientsDemographics ON tblPeople.IndexedName = tblPeopleClientsDemographics.IndexedName) LEFT JOIN tblPeopleClientsIndividualSupportServices ON tblPeopleClientsDemographics.IndexedName = tblPeopleClientsIndividualSupportServices.IndexedName " & _
        "WHERE (((tblPeopleClientsDemographics.ActiveIndivSupport) = True) And ((tblPeopleClientsIndividualSupportServices.Inactive) = False) And ((tblPeople.IsDeceased) = False));", dbSeeChanges: Call BriefDelay
    TILLDataBase.Execute "INSERT INTO [~IndivFinances] ( IndexedName, LastName, FirstName, MiddleInitial, WorkingABI, RepresentativePayee, RepPayeeIsClient, Service, Program, LegalName, SSAMonthlyAmount, SSIMonthlyAmount, SSPMonthlyAmount, Pension, OtherMonthlyAmount, WdlPercent, TotalBenefits, DiscountedBenefits, RemainingBenefits, SubtotalA, MonthlyEarnedIncome, Less65Dollars, SubtotalB, TotalMonthlyCharge ) " & _
        "SELECT tblPeople.IndexedName, tblPeople.LastName, tblPeople.FirstName, tblPeople.MiddleInitial, False AS WorkingABI, " & _
        "tblPeopleClientsDemographics.RepresentativePayee, tblPeopleClientsDemographics.RepPayeeIsClient, 'SL' AS Service, " & _
        "Null AS Program, " & _
        "[tblPeople]![FirstName] & ' ' & [tblPeople]![MiddleInitial] & ' ' & [tblPeople]![LastName] AS LegalName, " & _
        "IIf(IsNull([tblPeopleClientsDemographics]![SSAMonthlyAmount]),0,[tblPeopleClientsDemographics]![SSAMonthlyAmount]) AS SSAMonthlyAmount, " & _
        "IIf(IsNull([tblPeopleClientsDemographics]![SSIMonthlyAmount]),0,[tblPeopleClientsDemographics]![SSIMonthlyAmount]) AS SSIMonthlyAmount, " & _
        "IIf(IsNull([tblPeopleClientsDemographics]![SSPMonthlyAmount]),0,[tblPeopleClientsDemographics]![SSPMonthlyAmount]) AS SSPMonthlyAmount, " & _
        "IIf(IsNull([tblPeopleClientsDemographics]![Pension]),0,[tblPeopleClientsDemographics]![Pension]) AS Pension, " & _
        "IIf(IsNull([tblPeopleClientsDemographics]![OtherMonthlyAmount]),0,[tblPeopleClientsDemographics]![OtherMonthlyAmount]) AS OtherMonthlyAmount, tblPeopleClientsDemographics.WdlPercent, " & _
        "[SSAMonthlyAmount]+[SSIMonthlyAmount]+[SSPMonthlyAmount]+[OtherMonthlyAmount]+[Pension] AS TotalBenefits, " & _
        "[TotalBenefits]*[WdlPercent]/100 AS DiscountedBenefits, " & _
        "IIf([TotalBenefits]-[DiscountedBenefits]>=200,[TotalBenefits]-[DiscountedBenefits],200) AS RemainingBenefits, " & _
        "[DiscountedBenefits] AS SubtotalA, Null AS MonthlyEarnedIncome, Null AS Less65Dollars, Null AS SubtotalB, " & _
        "IIf(IsNull([SubtotalB]),[SubtotalA],[SubtotalA]+[SubtotalB]) AS TotalMonthlyCharge " & _
        "FROM (tblPeople INNER JOIN tblPeopleClientsDemographics ON tblPeople.IndexedName = tblPeopleClientsDemographics.IndexedName) LEFT JOIN tblPeopleClientsSharedLivingServices ON tblPeopleClientsDemographics.IndexedName = tblPeopleClientsSharedLivingServices.IndexedName " & _
        "WHERE (((tblPeopleClientsDemographics.ActiveSharedLiving) = True) And ((tblPeopleClientsSharedLivingServices.Inactive) = False) And ((tblPeople.IsDeceased) = False));", dbSeeChanges: Call BriefDelay
    
    If IndivFinDollarAmounts = "Blank" Then WhichLetter = "ltrINDIVFINANCESBLANK" Else WhichLetter = "ltrINDIVFINANCES"

    If IndivFinAllOrSingle = "ALL" Then
        RepPayee = Null
        Select Case IndivFinSelectRepPayee
            Case "All"
                Call BriefDelay
                Call ExecReport(WhichLetter)
            Case "TILL"
                Criteria2 = "RepresentativePayee Like 'TILL*'"
                Call BriefDelay
                DoCmd.OpenReport WhichLetter, acViewPreview, , Criteria2
            Case "Not TILL"
                Criteria2 = "(RepresentativePayee Not Like 'TILL*' OR RepPayeeIsClient=True) AND RepresentativePayee Is Not Null"
                Call BriefDelay
                DoCmd.OpenReport WhichLetter, acViewPreview, , Criteria2
        End Select
    Else
        RepPayee = DLookup("RepresentativePayee", "tblPeopleClientsDemographics", "IndexedName = """ & SelectedClientIndexedName & """")
        If IsNull(RepPayee) Or Len(RepPayee) <= 0 Then
            MsgBox "Individual has no Representative Payee identified.  Please correct this.", vbOKOnly, "Error"
            IndivFinEnterEffectiveMonth.Visible = True: IndivFinEnterEffectiveMonth = Null: IndivFinEnterEffectiveMonth.SetFocus
            IndivFinSelectRepPayee.Visible = False: IndivFinSelectRepPayee = Null: IndivFinDollarAmounts.Visible = False: IndivFinDollarAmounts = Null
            IndivFinAllOrSingle.Visible = False: IndivFinAllOrSingle = Null: IndivFinSelectClient.Visible = False: IndivFinSelectClient = Null
            Exit Sub
        End If
        Call BriefDelay
        DoCmd.OpenReport WhichLetter, acViewPreview, , "IndexedName = """ & SelectedClientIndexedName & """"
    End If
    ' Wait for report to be closed before contining.
    Call LoopUntilClosed(WhichLetter, 3)
    ' Reset all form fields.
    IndivFinEnterEffectiveMonth.Visible = True: IndivFinEnterEffectiveMonth = Null: IndivFinEnterEffectiveMonth.SetFocus
    IndivFinSelectRepPayee.Visible = False: IndivFinSelectRepPayee = Null: IndivFinDollarAmounts.Visible = False: IndivFinDollarAmounts = Null
    IndivFinAllOrSingle.Visible = False: IndivFinAllOrSingle = Null: IndivFinSelectClient.Visible = False: IndivFinSelectClient = Null
    Me.Refresh: Me.Requery: Me.Repaint
    TILLDataBase.Execute "DELETE [~IndivFinances].* FROM [~IndivFinances];"
    Exit Sub
ShowMeError:
    MsgBox "Error # " & Str(Err.Number) & " was generated by " & Me.Name & Chr(13) & Err.Description, vbOKOnly, "Error", Err.HelpFile, Err.HelpContext
End Sub

Private Sub IndivFinAllOrSingle_AfterUpdate()
    If IndivFinAllOrSingle = "All" Then
        Call IndivFinDoReport
    Else
        IndivFinSelectClient.Visible = True
        If IndivFinSelectRepPayee = "TILL" Then
            IndivFinSelectClient.RowSource = "SELECT AllClients.ReverseDisplayName, tblPeopleClientsDemographics.RepPayeeIsTILL FROM tblPeopleClientsDemographics INNER JOIN AllClients ON tblPeopleClientsDemographics.IndexedName = AllClients.IndexedName WHERE (AllClients.IsClientRes=True Or AllClients.IsCilentCLO=True Or AllClients.IsClientSharedLiving=True Or AllClients.IsClientIndiv=True) AND tblPeopleClientsDemographics.RepPayeeIsTILL = True ORDER BY AllClients.LastName, AllClients.FirstName;"
        ElseIf IndivFinSelectRepPayee = "Not TILL" Then
            IndivFinSelectClient.RowSource = "SELECT AllClients.ReverseDisplayName, tblPeopleClientsDemographics.RepPayeeIsTILL FROM tblPeopleClientsDemographics INNER JOIN AllClients ON tblPeopleClientsDemographics.IndexedName = AllClients.IndexedName WHERE (AllClients.IsClientRes=True Or AllClients.IsCilentCLO=True Or AllClients.IsClientSharedLiving=True Or AllClients.IsClientIndiv=True) AND tblPeopleClientsDemographics.RepPayeeIsTILL = False ORDER BY AllClients.LastName, AllClients.FirstName;"
        ElseIf IndivFinSelectRepPayee = "All" Then
            IndivFinSelectClient.RowSource = "SELECT AllClients.ReverseDisplayName FROM tblPeopleClientsDemographics INNER JOIN AllClients ON tblPeopleClientsDemographics.IndexedName = AllClients.IndexedName WHERE (((AllClients.IsClientRes)=True)) Or (((AllClients.IsCilentCLO)=True)) Or (((AllClients.[IsClientSharedLiving])=True)) Or (((AllClients.IsClientIndiv)=True)) ORDER BY AllClients.LastName, AllClients.FirstName;"
        End If
        IndivFinSelectClient.Requery
        IndivFinSelectClient.SetFocus
    End If
End Sub

Private Sub IndivFinDollarAmounts_AfterUpdate()
    IndivFinAllOrSingle.Visible = True: IndivFinAllOrSingle.SetFocus
End Sub

Private Sub IndivFinEnterEffectiveMonth_AfterUpdate()
    IndivFinSelectRepPayee.Visible = True: IndivFinSelectRepPayee.SetFocus
End Sub

Private Sub IndivFinSelectClient_AfterUpdate()
    Dim RepPayee As Variant, ClientFirstName As Variant, ClientLastName As Variant
    
    Call BuildAllClients(False)
    SelectedClientIndexedName = DLookup("IndexedName", "AllClients", "ReverseDisplayName = """ & IndivFinSelectClient & """")
    ClientLastName = Left(IndivFinSelectClient, InStr(1, IndivFinSelectClient, ",") - 1)
    ClientFirstName = Mid(IndivFinSelectClient, InStr(1, IndivFinSelectClient, ",") + 1, 255)
    RepPayee = DLookup("RepresentativePayee", "tblPeopleClientsDemographics", "IndexedName = """ & SelectedClientIndexedName & """")
    If IsNull(RepPayee) Or Len(RepPayee) <= 0 Then
        MsgBox "Individual has no Representative Payee identified.  Please correct this.", vbOKOnly, "Error"
        IndivFinEnterEffectiveMonth.Visible = True: IndivFinEnterEffectiveMonth = Null: IndivFinEnterEffectiveMonth.SetFocus
        IndivFinSelectRepPayee.Visible = False: IndivFinSelectRepPayee = Null: IndivFinDollarAmounts.Visible = False: IndivFinDollarAmounts = Null
        IndivFinAllOrSingle.Visible = False: IndivFinAllOrSingle = Null: IndivFinSelectClient.Visible = False: IndivFinSelectClient = Null
        Exit Sub
    End If

    If InStr(1, RepPayee, "TILL") > 0 Then
        MsgBox ClientFirstName & " " & ClientLastName & "'s representative payee is TILL.", vbOKOnly, "Representative Payee Status"
    ElseIf IsNull(RepPayee) Then
        MsgBox ClientFirstName & " " & ClientLastName & " has no representative payee.", vbOKOnly, "Representative Payee Status"
    Else
        MsgBox ClientFirstName & " " & ClientLastName & "'s representative payee is " & RepPayee & ".", vbOKOnly, "Representative Payee Status"
    End If
    
    Call IndivFinDoReport
End Sub

Private Function IndivFinEverythingFilled()
    IndivFinEverythingFilled = True
    If IsNull(IndivFinEnterEffectiveMonth) Then IndivFinEverythingFilled = False
    If IsNull(IndivFinEnterEffectiveYear) Then IndivFinEverythingFilled = False
    If IsNull(IndivFinDollarAmounts) Then IndivFinEverythingFilled = False
    If IndivFinAllOrSingle = "Single" Then If IsNull(IndivFinSelectClient) Then IndivFinEverythingFilled = False
End Function

Private Sub IndivFinSelectRepPayee_AfterUpdate()
    IndivFinDollarAmounts.Visible = True: IndivFinDollarAmounts.SetFocus
End Sub

Private Sub MAFY_AfterUpdate()
    Call BriefDelay: DoCmd.OpenReport "rptCONSUMMARY", acViewPreview, , "FY=" & MAFY, , MAFY: Call LoopUntilClosed("rptCONSUMMARY", 3)
    MAFY = Null: Me.Refresh: Me.Requery: Me.Repaint
End Sub

Private Sub MedDoReport(Optional DoBoth As Boolean = False)
On Error GoTo ShowMeError
    Dim Comp As Boolean, WhichLetter As String, RptName1 As String, RptName2 As String, RptName3 As String, RptName4 As String
    
    Call BuildAllClients(False)
    Call DropTempTables
    Comp = False
    
    RptName1 = "ltrEMGAUTH": RptName2 = "ltrRTNAUTH": RptName3 = "ltrEMGAUTHSelfApproval": RptName4 = "ltrRTNAUTHSelfApproval"
    
    DoCmd.CopyObject , "temptbl1", acTable, "tblPeople"
    
    TILLDataBase.Execute "CREATE TABLE temptbl2 (FamilyIndexedName CHAR(160), FamilyLastName CHAR(50), FamilyFirstName CHAR(50), FamilyMiddleInitial CHAR(1), FamilyCompanyOrganization CHAR(40), ClientIndexedName CHAR(160), ClientLastName CHAR(50), ClientFirstName CHAR(50), ClientMiddleInitial CHAR(1), Pronoun CHAR(3));", dbSeeChanges: Call BriefDelay
    TILLDataBase.Execute "INSERT INTO temptbl2 (FamilyIndexedName, FamilyLastName, FamilyFirstName, FamilyMiddleInitial, FamilyCompanyOrganization, ClientIndexedName, ClientLastName, ClientFirstName, ClientMiddleInitial, Pronoun ) " & vbCrLf & _
        "SELECT tblPeopleFamily.IndexedName AS FamilyIndexedName, tblPeople.LastName AS FamilyLastName, tblPeople.FirstName AS FamilyFirstName, tblPeople.MiddleInitial AS FamilyMiddleInitial, tblPeople.CompanyOrganization AS FamilyCompanyOrganization, tblPeopleFamily.ClientIndexedName, temptbl1.LastName AS ClientLastName, temptbl1.FirstName AS ClientFirstName, temptbl1.MiddleInitial AS ClientMiddleInitial, IIf([tblPeopleClientsDemographics]![Gender]='M','his','her') AS Pronoun " & vbCrLf & _
        "FROM ((tblPeopleFamily INNER JOIN tblPeople ON tblPeopleFamily.IndexedName = tblPeople.IndexedName) INNER JOIN temptbl1 ON tblPeopleFamily.ClientIndexedName = temptbl1.IndexedName) INNER JOIN tblPeopleClientsDemographics ON temptbl1.IndexedName = tblPeopleClientsDemographics.IndexedName " & vbCrLf & _
        "WHERE tblPeopleFamily.ClientIndexedName=""" & SelectedClientIndexedName & """ AND tblPeople.IsDeceased=False AND tblPeopleFamily.Guardian=Yes AND tblPeopleFamily.Inactive=False;", dbSeeChanges: Call BriefDelay
    TILLDataBase.Execute "CREATE TABLE temptbl3 (DisplayName CHAR(160), ClientIndexedName CHAR(160), ClientLastName CHAR(25), ClientFirstName CHAR(25), ClientMiddleInitial CHAR(1));", dbSeeChanges: Call BriefDelay
    TILLDataBase.Execute "INSERT INTO temptbl3 (DisplayName, ClientIndexedName, ClientLastName, ClientFirstName, ClientMiddleInitial) " & vbCrLf & _
        "SELECT DisplayName, IndexedName, LastName, FirstName, MiddleInitial " & vbCrLf & _
        "FROM AllClients WHERE IndexedName=""" & Form_frmRptFinancialAndLetters.SelectedClientIndexedName & """;", dbSeeChanges: Call BriefDelay
    
    If DLookup("LegalStatus", "tblPeopleClientsDemographics", "IndexedName=""" & SelectedClientIndexedName & """") = "Competent" Or DLookup("LegalStatus", "tblPeopleClientsDemographics", "IndexedName=""" & SelectedClientIndexedName & """") = "Presumed Competent" Then Comp = True

    If DoBoth Then
        MedEnterEffectiveMonth = AllConsentFormsEnterEffectiveMonth
        If (Not Comp) And DCount("FamilyIndexedName", "temptbl2", "ClientIndexedName = """ & SelectedClientIndexedName & """") > 0 Then
            ' Incompetent with at least one Guardian.
            Call BriefDelay
            DoCmd.OpenReport RptName1, acViewPreview, , "ClientIndexedName = """ & SelectedClientIndexedName & """"
            Call LoopUntilClosed(RptName1, 3)
            Call BriefDelay
            DoCmd.OpenReport RptName2, acViewPreview, , "ClientIndexedName = """ & SelectedClientIndexedName & """"
            Call LoopUntilClosed(RptName2, 3)
        ElseIf Comp Then
            ' Competent
            Call BriefDelay
            DoCmd.OpenReport RptName3, acViewPreview, , "ClientIndexedName = """ & SelectedClientIndexedName & """"
            Call LoopUntilClosed(RptName3, 3)
            Call BriefDelay
            DoCmd.OpenReport RptName4, acViewPreview, , "ClientIndexedName = """ & SelectedClientIndexedName & """"
            Call LoopUntilClosed(RptName4, 3)
        Else
            ' Incompetent but no guardians.  Flag as error and get out.
            MsgBox "Individual is flagged as incompetent but has no guardian(s).  Form cannot be produced.", vbOKOnly, "Error!"
            MedSelectClient = Null: MedEnterEffectiveMonth = Null: MedSelectedLetterType = Null: Me.Refresh: Me.Requery: Me.Repaint
            ErrorOccurred = True
            Exit Sub
        End If
    Else
        Select Case MedSelectedLetterType
            Case "Emergency"
                If (Not Comp) And DCount("FamilyIndexedName", "temptbl2", "ClientIndexedName = """ & SelectedClientIndexedName & """") > 0 Then
                    ' Incompetent with at least one Guardian.
                    Call BriefDelay: DoCmd.OpenReport RptName1, acViewPreview, , "ClientIndexedName = """ & SelectedClientIndexedName & """"
                    Call LoopUntilClosed(RptName1, 3)
                ElseIf Comp Then
                    ' Competent.
                    Call BriefDelay: DoCmd.OpenReport RptName3, acViewPreview, , "ClientIndexedName = """ & SelectedClientIndexedName & """"
                    Call LoopUntilClosed(RptName3, 3)
                Else
                    ' Incompetent but no guardians.  Flag as error and get out.
                    MsgBox "Individual is flagged as incompetent but has no guardian(s).  Form cannot be produced.", vbOKOnly, "Error!"
                    MedSelectClient = Null: MedEnterEffectiveMonth = Null: MedSelectedLetterType = Null: Me.Refresh: Me.Requery: Me.Repaint
                    Exit Sub
                End If
            Case "Routine"
                If (Not Comp) And DCount("FamilyIndexedName", "temptbl2", "ClientIndexedName = """ & SelectedClientIndexedName & """") > 0 Then
                    ' Incompetent with at least one Guardian.
                    Call BriefDelay: DoCmd.OpenReport RptName2, acViewPreview, , "ClientIndexedName = """ & SelectedClientIndexedName & """"
                    Call LoopUntilClosed(RptName2, 3)
                ElseIf Comp Then
                    ' Competent.
                    Call BriefDelay: DoCmd.OpenReport RptName4, acViewPreview, , "ClientIndexedName = """ & SelectedClientIndexedName & """"
                    Call LoopUntilClosed(RptName4, 3)
                Else
                    ' Incompetent but no guardians.  Flag as error and get out.
                    MsgBox "Individual is flagged as incompetent but has no guardian(s).  Form cannot be produced.", vbOKOnly, "Error!"
                    MedSelectClient = Null: MedEnterEffectiveMonth = Null: MedSelectedLetterType = Null: Me.Refresh: Me.Requery: Me.Repaint
                    Exit Sub
                End If
            Case "Both"
                If (Not Comp) And DCount("FamilyIndexedName", "temptbl2", "ClientIndexedName = """ & SelectedClientIndexedName & """") > 0 Then
                    ' Incompetent with at least one Guardian.
                    Call BriefDelay: DoCmd.OpenReport RptName1, acViewPreview, , "ClientIndexedName = """ & SelectedClientIndexedName & """"
                    Call LoopUntilClosed(RptName1, 3)
                    Call BriefDelay: DoCmd.OpenReport RptName2, acViewPreview, , "ClientIndexedName = """ & SelectedClientIndexedName & """"
                    Call LoopUntilClosed(RptName2, 3)
                ElseIf Comp Then
                    ' Competent.
                    Call BriefDelay: DoCmd.OpenReport RptName3, acViewPreview, , "ClientIndexedName = """ & SelectedClientIndexedName & """"
                    Call LoopUntilClosed(RptName3, 3)
                    Call BriefDelay: DoCmd.OpenReport RptName4, acViewPreview, , "ClientIndexedName = """ & SelectedClientIndexedName & """"
                    Call LoopUntilClosed(RptName4, 3)
                Else
                    ' Incompetent but no guardians.  Flag as error and get out.
                    MsgBox "Individual is flagged as incompetent but has no guardian(s).  Form cannot be produced.", vbOKOnly, "Error!"
                    MedSelectClient = Null: MedEnterEffectiveMonth = Null: MedSelectedLetterType = Null: Me.Refresh: Me.Requery: Me.Repaint
                    Exit Sub
                End If
        End Select
    End If
    Call DropTempTables
    MedSelectClient = Null: MedEnterEffectiveMonth = Null: MedSelectedLetterType = Null: Me.Refresh: Me.Requery: Me.Repaint
    Exit Sub
ShowMeError:
    MsgBox "Error # " & Str(Err.Number) & " was generated by " & Me.Name & Chr(13) & Err.Description, vbOKOnly, "Error", Err.HelpFile, Err.HelpContext
End Sub

Private Sub AllConsentFormsEnterEffectiveMonth_AfterUpdate()
    If (Len(SelectedClientIndexedName) > 0) And (Len(AllConsentFormsEnterEffectiveMonth) > 0) Then
        Call MedDoReport(True)
        If ErrorOccurred Then
            ErrorOccurred = False: AllConsentFormsEnterEffectiveMonth = Null: AllConsentFormsSelectClient = Null
            Exit Sub
        End If
        Call MEDIADoReport(True)
        Call SUPREPDoReport(True)
        Call TRVLDoReport(True)
        Call BuildAllClients(False)
        AllConsentFormsEnterEffectiveMonth = Null: AllConsentFormsSelectClient = Null
    End If
End Sub

Private Sub AllConsentFormsSelectClient_AfterUpdate()
    Call BuildAllClients(False)
    SelectedClientIndexedName = DLookup("IndexedName", "AllClients", "ReverseDisplayName = """ & AllConsentFormsSelectClient & """")
    If (Len(SelectedClientIndexedName) > 0) And (Len(AllConsentFormsEnterEffectiveMonth) > 0) Then
        Call MedDoReport(True)
        If ErrorOccurred Then
            ErrorOccurred = False: AllConsentFormsEnterEffectiveMonth = Null: AllConsentFormsSelectClient = Null
            Exit Sub
        End If
        Call MEDIADoReport(True)
        Call SUPREPDoReport(True)
        Call TRVLDoReport(True)
        Call BuildAllClients(False)
        AllConsentFormsEnterEffectiveMonth = Null: AllConsentFormsSelectClient = Null
    End If
End Sub

Private Sub MedEnterEffectiveMonth_AfterUpdate()
    If (Len(SelectedClientIndexedName) > 0) And (Len(MedEnterEffectiveMonth) > 0) And (Len(MedSelectedLetterType) > 0) Then
        Call MedDoReport
        Call BuildAllClients(False)
    End If
End Sub

Private Sub MedSelectClient_AfterUpdate()
    Call BuildAllClients(False)
    SelectedClientIndexedName = DLookup("IndexedName", "AllClients", "ReverseDisplayName = """ & MedSelectClient & """")
    If (Len(SelectedClientIndexedName) > 0) And (Len(MedEnterEffectiveMonth) > 0) And (Len(MedSelectedLetterType) > 0) Then
        Call MedDoReport
        Call BuildAllClients(False)
    End If
End Sub

Private Sub MedSelectedLetterType_AfterUpdate()
    If (Len(SelectedClientIndexedName) > 0) And (Len(MedEnterEffectiveMonth) > 0) And (Len(MedSelectedLetterType) > 0) Then
        Call MedDoReport
        Call BuildAllClients(False)
    End If
End Sub

Private Sub MEDIASelectClient_AfterUpdate()
    SelectedClientIndexedName = DLookup("IndexedName", "AllClients", "ReverseDisplayName = """ & MEDIASelectClient & """")
    If Len(SelectedClientIndexedName) > 0 And Len(MEDIAEnterEffectiveMonth) > 0 Then Call MEDIADoReport
End Sub

Private Sub MEDIAEnterEffectiveMonth_AfterUpdate()
    If Len(SelectedClientIndexedName) > 0 And Len(MEDIAEnterEffectiveMonth) > 0 Then Call MEDIADoReport
End Sub

Private Sub MEDIADoReport(Optional FromAllConsentForms As Boolean = False)
On Error GoTo ShowMeError
    Dim Comp As Boolean, WhichLetter As String
    
    Call BuildAllClients(False)
    Call DropTempTables
    Comp = False
    
    DoCmd.CopyObject , "temptbl1", acTable, "tblPeople"
    TILLDataBase.Execute "CREATE TABLE temptbl2 (FamilyIndexedName CHAR(160), FamilyLastName CHAR(50), FamilyFirstName CHAR(50), FamilyMiddleInitial CHAR(1), FamilyCompanyOrganization CHAR(40), ClientIndexedName CHAR(160), ClientLastName CHAR(50), ClientFirstName CHAR(50), ClientMiddleInitial CHAR(1), Pronoun CHAR(3));", dbSeeChanges: Call BriefDelay
    TILLDataBase.Execute "INSERT INTO temptbl2 ( FamilyIndexedName, FamilyLastName, FamilyFirstName, FamilyMiddleInitial, FamilyCompanyOrganization, ClientIndexedName, ClientLastName, ClientFirstName, ClientMiddleInitial, Pronoun ) " & vbCrLf & _
        "SELECT tblPeopleFamily.IndexedName AS FamilyIndexedName, tblPeople.LastName AS FamilyLastName, tblPeople.FirstName AS FamilyFirstName, tblPeople.MiddleInitial AS FamilyMiddleInitial, tblPeople.CompanyOrganization AS FamilyCompanyOrganization, tblPeopleFamily.ClientIndexedName, temptbl1.LastName AS ClientLastName, temptbl1.FirstName AS ClientFirstName, temptbl1.MiddleInitial AS ClientMiddleInitial, IIf([tblPeopleClientsDemographics]![Gender]='M','his','her') AS Pronoun " & vbCrLf & _
        "FROM ((tblPeopleFamily INNER JOIN tblPeople ON tblPeopleFamily.IndexedName = tblPeople.IndexedName) INNER JOIN temptbl1 ON tblPeopleFamily.ClientIndexedName = temptbl1.IndexedName) INNER JOIN tblPeopleClientsDemographics ON temptbl1.IndexedName = tblPeopleClientsDemographics.IndexedName " & vbCrLf & _
        "WHERE tblPeople.IsDeceased=False AND tblPeopleFamily.Guardian=Yes AND tblPeopleFamily.Inactive=False AND " & _
            "(tblPeopleClientsDemographics.ActiveResidentialServices=True OR tblPeopleClientsDemographics.ActiveCLO=True OR tblPeopleClientsDemographics.ActiveSharedLiving=True OR tblPeopleClientsDemographics.ActiveAFC=True OR tblPeopleClientsDemographics.ActiveIndivSupport=True);", dbSeeChanges: Call BriefDelay
    TILLDataBase.Execute "CREATE TABLE temptbl3 (DisplayName CHAR(160), ClientIndexedName CHAR(160), ClientLastName CHAR(25), ClientFirstName CHAR(25), ClientMiddleInitial CHAR(1));", dbSeeChanges: Call BriefDelay
    TILLDataBase.Execute "INSERT INTO temptbl3 (DisplayName, ClientIndexedName, ClientLastName, ClientFirstName, ClientMiddleInitial) " & vbCrLf & _
        "SELECT DisplayName, IndexedName, LastName, FirstName, MiddleInitial " & vbCrLf & _
        "FROM AllClients WHERE IndexedName=""" & SelectedClientIndexedName & """;", dbSeeChanges: Call BriefDelay

    If FromAllConsentForms Then MEDIAEnterEffectiveMonth = AllConsentFormsEnterEffectiveMonth
    If DLookup("LegalStatus", "tblPeopleClientsDemographics", "IndexedName=""" & SelectedClientIndexedName & """") = "Competent" Or DLookup("LegalStatus", "tblPeopleClientsDemographics", "IndexedName=""" & SelectedClientIndexedName & """") = "Presumed Competent" Then Comp = True
    
    If (Not Comp) And DCount("FamilyIndexedName", "temptbl2", "ClientIndexedName = """ & SelectedClientIndexedName & """") > 0 Then
        ' Incompetent and has guardian(s).
        WhichLetter = "ltrMEDIARELEASEINCOMPETENT"
    ElseIf Comp Then
        ' Competent.
        WhichLetter = "ltrMEDIARELEASECOMPETENT"
    Else
        ' Incompetent but no guardians.  Flag as error and get out.
        MsgBox "Individual is flagged as incompetent but has no guardian(s).  Form cannot be produced.", vbOKOnly, "Error!"
        MEDIASelectClient = Null: MEDIAEnterEffectiveMonth = Null: Me.Refresh: Me.Requery: Me.Repaint
        Exit Sub
    End If
    
    Call BriefDelay
    DoCmd.OpenReport WhichLetter, acViewPreview, , "ClientIndexedName = """ & SelectedClientIndexedName & """"
    Call LoopUntilClosed(WhichLetter, 3)
    Call DropTempTables
    MEDIASelectClient = Null: MEDIAEnterEffectiveMonth = Null: Me.Refresh: Me.Requery: Me.Repaint
    Exit Sub
ShowMeError:
    MsgBox "Error # " & Str(Err.Number) & " was generated by " & Me.Name & Chr(13) & Err.Description, vbOKOnly, "Error", Err.HelpFile, Err.HelpContext
End Sub

Private Sub MRSelectMonth_AfterUpdate()
    MRSelectedMonth = Left(MRSelectMonth, InStr(1, MRSelectMonth, " ", vbTextCompare) - 1)
    MRSelectedYear = Right(MRSelectMonth, 4)
    
    Select Case MRSelectMonth
        Case "April " & Format(MRFY - 1, "####")
            MRFY = MRFY - 1
            MRNumDays = 31
        Case "May " & Format(MRFY - 1, "####")
            MRFY = MRFY - 1
            MRNumDays = 31
        Case "June " & Format(MRFY - 1, "####")
            MRFY = MRFY - 1
            MRNumDays = 30
        Case "July " & Format(MRFY - 1, "####")
            MRNumDays = 31
        Case "August " & Format(MRFY - 1, "####")
            MRNumDays = 31
        Case "September " & Format(MRFY - 1, "####")
            MRNumDays = 30
        Case "October " & Format(MRFY - 1, "####")
            MRNumDays = 31
        Case "November " & Format(MRFY - 1, "####")
            MRNumDays = 30
        Case "December " & Format(MRFY - 1, "####")
            MRNumDays = 31
        Case "January " & Format(MRFY, "####")
            MRNumDays = 31
        Case "February " & Format(MRFY, "####")
            If IsDate("02/29/" & Format(MRFY, "####")) Then MRNumDays = 29 Else MRNumDays = 28
        Case "March " & Format(MRFY, "####")
            MRNumDays = 31
        Case "April " & Format(MRFY, "####")
            MRNumDays = 30
        Case "May " & Format(MRFY, "####")
            MRNumDays = 31
        Case "June " & Format(MRFY, "####")
            MRNumDays = 30
    End Select
    MRMask.Visible = False
End Sub

Private Sub NHFY_AfterUpdate()
On Error GoTo ShowMeError
    Call DropTempTables
    TILLDataBase.Execute "SELECT tblContracts.FY, tblContracts.ContractID, tblContractsBillingBook.BIllingBookNumber, tblContracts.ActivityCode, tblContracts.GPNumber, tblContracts.State, tblContracts.UnitRate, tblContracts.Units AS ContractUnits, tblContracts.MaximumObligation AS ContractMaximumObligation, tblContracts.MaximumObligationAsAmended, tblContracts.TotalClients, tblContracts.TotalUnits, tblContracts.TotalUnitsAsAmended, tblContractsBillingBook.ProgramName, tblContractsBillingBook.CostCenter, tblContractsBillingBook.MaximumObligation AS BillingBookMaximumObligation, tblContractsBillingBook.Units AS BillingBookUnits, tblContractsBillingBook.BillingRate, tblContractsBillingBook.FundingSource, tblContractsBillingBook.DDSArea, tblContractsBillingBook.Staff " & _
        "INTO temptbl0 " & _
        "FROM tblContracts LEFT JOIN tblContractsBillingBook ON (tblContracts.FY = tblContractsBillingBook.FY) AND (tblContracts.ContractID = tblContractsBillingBook.ContractID) " & _
        "WHERE Pending = False AND tblContracts.FY = " & NHFY & " And tblContractsBillingBook.BillingBookNumber >= 600 And tblContractsBillingBook.BillingBookNumber <= 699 " & _
        "ORDER BY tblContractsBillingBook.BIllingBookNumber;", dbSeeChanges: Call BriefDelay
    Call BriefDelay: DoCmd.OpenReport "rptNHCONTRACTS", acViewPreview, , , , NHFY
    Call LoopUntilClosed("rptNHCONTRACTS", 3)
    NHFY = Null:    Me.Refresh:    Me.Requery:    Me.Repaint
    Exit Sub
ShowMeError:
    MsgBox "Error # " & Str(Err.Number) & " was generated by " & Me.Name & Chr(13) & Err.Description, vbOKOnly, "Error", Err.HelpFile, Err.HelpContext
End Sub

Private Sub RCBSelectAllClients_Click()
    Call ExecReport("rptRESBENEFITS")
End Sub

Private Sub RCBSelectClient_AfterUpdate()
    Call BuildAllClients(False)
    SelectedClientIndexedName = DLookup("IndexedName", "AllClients", "ReverseDisplayName = """ & RCBSelectClient & """")
    Call BriefDelay: DoCmd.OpenReport "rptRESBENEFITS", acViewPreview, , "IndexedName = """ & SelectedClientIndexedName & """": Call LoopUntilClosed("rptRESBENEFITS", 3)
    RCBSelectClient = Null: Me.Refresh: Me.Requery: Me.Repaint
End Sub

Private Sub REPPAYEEFEESelectClient_AfterUpdate()
    ReportName = "ltrREPPAYEEFEE"
    SelectedClientIndexedName = DLookup("IndexedName", "AllClients", "ReverseDisplayName = """ & REPPAYEEFEESelectClient & """")
    If Len(SelectedClientIndexedName) > 0 Then Call ExecReport(ReportName)
    REPPAYEEFEESelectClient = Null: Me.Refresh: Me.Requery: Me.Repaint
End Sub

Private Sub REPPAYEEFEEAllClients_Click()
    ReportName = "ltrREPPAYEEFEEAllClients"
    Call BriefDelay
    DoCmd.OpenReport ReportName, acViewPreview, , "RepPayeeIsTILL=True"
    Call LoopUntilClosed(ReportName, 3)
End Sub

Private Sub REPPAYEES_Click()
    DoCmd.OpenForm "frmRptREPPAYEES"
End Sub

Private Sub SECT8SelectClient_AfterUpdate()
    ReportName = "ltrSECTION8AUTH"
    SelectedClientIndexedName = DLookup("IndexedName", "AllClients", "ReverseDisplayName = """ & SECT8SelectCLient & """")
    SelectedFirstName = DLookup("FirstName", "AllClients", "ReverseDisplayName = """ & SECT8SelectCLient & """")
    SelectedLastName = DLookup("LastName", "AllClients", "ReverseDisplayName = """ & SECT8SelectCLient & """")
    If Len(SelectedClientIndexedName) > 0 Then Call ExecReport(ReportName)
    SECT8SelectCLient = Null: Me.Refresh: Me.Requery: Me.Repaint
End Sub

Private Sub IntroSelectClient_AfterUpdate()
    Dim IsRes As Boolean, IsCLO As Boolean

    ReportName = "ltrINTRODUCTION"
    Call BuildAllClients(False)
    SelectedClientIndexedName = DLookup("IndexedName", "AllClients", "ReverseDisplayName = """ & IntroSelectClient & """")
    SelectedFirstName = DLookup("FirstName", "AllClients", "ReverseDisplayName = """ & IntroSelectClient & """")
    SelectedLastName = DLookup("LastName", "AllClients", "ReverseDisplayName = """ & IntroSelectClient & """")
    IsRes = DLookup("isClientRes", "AllClients", "ReverseDisplayName = """ & IntroSelectClient & """")
    IsCLO = DLookup("isCilentCLO", "AllClients", "ReverseDisplayName = """ & IntroSelectClient & """")
    If IsRes Then
        SelectedLocation = DLookup("Location", "tblPeopleClientsResidentialServices", "IndexedName = """ & SelectedClientIndexedName & """")
    ElseIf IsCLO Then
        SelectedLocation = DLookup("Location", "tblPeopleClientsCLOServices", "IndexedName = """ & SelectedClientIndexedName & """")
    Else
        SelectedLocation = Null
    End If
    If Len(SelectedClientIndexedName) > 0 Then Call ExecReport(ReportName)
    IntroSelectClient = Null: Me.Refresh: Me.Requery: Me.Repaint
End Sub

Private Sub SigOnFileSelectClient_AfterUpdate()
On Error GoTo ShowMeError
    Dim RepPayee As Variant
    ReportName = "ltrSIGONFILE"
    Call DropTempTables
    Select Case SigOnFileSelectClient
        Case "TILL is Rep Payee"
            TILLDataBase.Execute "SELECT tblPeople.LastName, tblPeople.FirstName INTO temptbl0 FROM tblPeople LEFT JOIN tblPeopleClientsDemographics ON tblPeople.IndexedName = tblPeopleClientsDemographics.IndexedName " & _
                "WHERE tblPeopleClientsDemographics.RepresentativePayee Like 'TILL*' AND (tblPeople.IsClientRes=True OR tblPeople.IsCilentCLO=True) ORDER BY tblPeople.LastName, tblPeople.FirstName;", dbSeeChanges: Call BriefDelay
        Case "TILL is not Rep Payee"
            TILLDataBase.Execute "SELECT tblPeople.LastName, tblPeople.FirstName INTO temptbl0 FROM tblPeople LEFT JOIN tblPeopleClientsDemographics ON tblPeople.IndexedName = tblPeopleClientsDemographics.IndexedName " & _
                "WHERE tblPeopleClientsDemographics.RepresentativePayee NOT Like 'TILL*' AND (tblPeople.IsClientRes=True OR tblPeople.IsCilentCLO=True) ORDER BY tblPeople.LastName, tblPeople.FirstName;", dbSeeChanges: Call BriefDelay
        Case "Single Person"
            If SigOnFileClient.Visible = False Then
                SigOnFileClient.Visible = True
                Exit Sub
            End If
            If IsNull(SigOnFileClient) Then
                Exit Sub
            Else
                SigOnFileClientIndexedName = DLookup("IndexedName", "AllClients", "ReverseDisplayName=""" & SigOnFileClient & """")
                TILLDataBase.Execute "SELECT tblPeople.LastName, tblPeople.FirstName INTO temptbl0 FROM tblPeople LEFT JOIN tblPeopleClientsDemographics ON tblPeople.IndexedName = tblPeopleClientsDemographics.IndexedName " & _
                    "WHERE tblPeople.IndexedName= """ & SigOnFileClientIndexedName & """", dbSeeChanges: Call BriefDelay
            End If
    End Select
    Call BriefDelay: DoCmd.OpenReport ReportName, acViewPreview: Call LoopUntilClosed(ReportName, 3)
    SigOnFileSelectClient = Null: SigOnFileSelectClient.SetFocus: SigOnFileClient.Visible = False: SigOnFileClient = Null: Me.Refresh: Me.Requery: Me.Repaint
    Call DropTempTables
    Exit Sub
ShowMeError:
    MsgBox "Error # " & Str(Err.Number) & " was generated by " & Me.Name & Chr(13) & Err.Description, vbOKOnly, "Error", Err.HelpFile, Err.HelpContext
End Sub

Private Sub SigOnFileClient_AfterUpdate()
    Call SigOnFileSelectClient_AfterUpdate
End Sub

Private Sub SSASelectClient_AfterUpdate()
    Dim RepPayee As Variant
    
    ReportName = "ltrSSAWARD"
    Call BuildAllClients(False)
    SelectedClientIndexedName = DLookup("IndexedName", "AllClients", "ReverseDisplayName = """ & SSASelectClient & """")
    SelectedLastName = DLookup("LastName", "AllClients", "ReverseDisplayName = """ & SSASelectClient & """")
    SelectedFirstName = DLookup("FirstName", "AllClients", "ReverseDisplayName = """ & SSASelectClient & """")
    RepPayee = DLookup("RepresentativePayee", "tblPeopleClientsDemographics", "IndexedName = """ & SelectedClientIndexedName & """")
    If IsNull(RepPayee) Or Len(RepPayee) <= 0 Then
        MsgBox "Individual has no Representative Payee identified.  Please correct this.", vbOKOnly, "Error"
        SSASelectClient = Null
        Exit Sub
    End If
    
    Call BriefDelay: DoCmd.OpenReport ReportName, acViewPreview: Call LoopUntilClosed(ReportName, 3)
    SSASelectClient = Null: Me.Refresh: Me.Requery: Me.Repaint
End Sub

Private Sub SUPREPDoReport(Optional FromAllConsentForms As Boolean = False)
On Error GoTo ShowMeError
    Dim Comp As Boolean, WhichLetter As String
    
    Call BuildAllClients(False)
    Call DropTempTables
    Comp = False
    
    DoCmd.CopyObject , "temptbl1", acTable, "tblPeople"
    TILLDataBase.Execute "CREATE TABLE temptbl2 (FamilyIndexedName CHAR(160), FamilyLastName CHAR(50), FamilyFirstName CHAR(50), FamilyMiddleInitial CHAR(1), FamilyCompanyOrganization CHAR(40), ClientIndexedName CHAR(160), ClientLastName CHAR(50), ClientFirstName CHAR(50), ClientMiddleInitial CHAR(1), Pronoun CHAR(3));", dbSeeChanges: Call BriefDelay
    TILLDataBase.Execute "INSERT INTO temptbl2 (FamilyIndexedName, FamilyLastName, FamilyFirstName, FamilyMiddleInitial, FamilyCompanyOrganization, ClientIndexedName, ClientLastName, ClientFirstName, ClientMiddleInitial, Pronoun ) " & vbCrLf & _
        "SELECT tblPeopleFamily.IndexedName AS FamilyIndexedName, tblPeople.LastName AS FamilyLastName, tblPeople.FirstName AS FamilyFirstName, tblPeople.MiddleInitial AS FamilyMiddleInitial, tblPeople.CompanyOrganization AS FamilyCompanyOrganization, tblPeopleFamily.ClientIndexedName, temptbl1.LastName AS ClientLastName, temptbl1.FirstName AS ClientFirstName, temptbl1.MiddleInitial AS ClientMiddleInitial, IIf([tblPeopleClientsDemographics]![Gender]='M','his','her') AS Pronoun " & vbCrLf & _
        "FROM ((tblPeopleFamily INNER JOIN tblPeople ON tblPeopleFamily.IndexedName = tblPeople.IndexedName) INNER JOIN temptbl1 ON tblPeopleFamily.ClientIndexedName = temptbl1.IndexedName) INNER JOIN tblPeopleClientsDemographics ON temptbl1.IndexedName = tblPeopleClientsDemographics.IndexedName " & vbCrLf & _
        "WHERE tblPeople.IsDeceased=False AND tblPeopleFamily.Guardian=Yes AND tblPeopleFamily.Inactive=False AND " & _
            "(tblPeopleClientsDemographics.ActiveResidentialServices=True OR tblPeopleClientsDemographics.ActiveCLO=True OR tblPeopleClientsDemographics.ActiveSharedLiving=True OR tblPeopleClientsDemographics.ActiveAFC=True OR tblPeopleClientsDemographics.ActiveIndivSupport=True);", dbSeeChanges: Call BriefDelay
    TILLDataBase.Execute "CREATE TABLE temptbl3 (DisplayName CHAR(160), ClientIndexedName CHAR(160), ClientLastName CHAR(25), ClientFirstName CHAR(25), ClientMiddleInitial CHAR(1));", dbSeeChanges: Call BriefDelay
    TILLDataBase.Execute "INSERT INTO temptbl3 (DisplayName, ClientIndexedName, ClientLastName, ClientFirstName, ClientMiddleInitial) " & vbCrLf & _
        "SELECT DisplayName, IndexedName, LastName, FirstName, MiddleInitial " & vbCrLf & _
        "FROM AllClients WHERE IndexedName=""" & SelectedClientIndexedName & """;", dbSeeChanges: Call BriefDelay

    If FromAllConsentForms Then SUPREPEnterEffectiveMonth = AllConsentFormsEnterEffectiveMonth
    If DLookup("LegalStatus", "tblPeopleClientsDemographics", "IndexedName=""" & SelectedClientIndexedName & """") = "Competent" Or DLookup("LegalStatus", "tblPeopleClientsDemographics", "IndexedName=""" & SelectedClientIndexedName & """") = "Presumed Competent" Then Comp = True
    
    If (Not Comp) And DCount("FamilyIndexedName", "temptbl2", "ClientIndexedName = """ & SelectedClientIndexedName & """") > 0 Then
        ' Incompetent and has guardian(s).
        WhichLetter = "ltrSUPREL"
    ElseIf Comp Then
        ' Competent.
        WhichLetter = "ltrSUPRELSelfApproval"
    Else
        ' Incompetent but no guardians.  Flag as error and get out.
        MsgBox "Individual is flagged as incompetent but has no guardian(s).  Form cannot be produced.", vbOKOnly, "Error!"
        SUPREPSelectClient = Null: SUPREPEnterEffectiveMonth = Null: Me.Refresh: Me.Requery: Me.Repaint
        Exit Sub
    End If

    Call BriefDelay: DoCmd.OpenReport WhichLetter, acViewPreview, , "ClientIndexedName = """ & SelectedClientIndexedName & """": Call LoopUntilClosed(WhichLetter, 3)
    SUPREPSelectClient = Null: SUPREPEnterEffectiveMonth = Null: Me.Refresh: Me.Requery: Me.Repaint
    Call DropTempTables
    Exit Sub
ShowMeError:
    MsgBox "Error # " & Str(Err.Number) & " was generated by " & Me.Name & Chr(13) & Err.Description, vbOKOnly, "Error", Err.HelpFile, Err.HelpContext
End Sub

Private Sub SUPREPEnterEffectiveMonth_AfterUpdate()
    If (Not IsNull(SelectedClientIndexedName)) And (Not IsNull(SUPREPEnterEffectiveMonth)) Then Call SUPREPDoReport
End Sub

Private Sub SUPREPSelectClient_AfterUpdate()
    Call BuildAllClients(False)
    SelectedClientIndexedName = DLookup("IndexedName", "AllClients", "ReverseDisplayName = """ & SUPREPSelectClient & """")
    If Len(SelectedClientIndexedName) > 0 And Len(SUPREPEnterEffectiveMonth) > 0 Then Call SUPREPDoReport
End Sub

Private Sub TRVREL_Click()
    DoCmd.OpenForm "frmRptLtrTRVREL"
End Sub

Private Sub TRVLDoReport(Optional FromAllConsentForms As Boolean = False)
On Error GoTo ShowMeError
    Dim Comp As Boolean, WhichLetter As String
    
    Call BuildAllClients(False)
    Call DropTempTables
    DoCmd.CopyObject , "temptbl1", acTable, "tblPeople"
    Comp = False

    TILLDataBase.Execute "CREATE TABLE temptbl2 (FamilyIndexedName CHAR(160), FamilyLastName CHAR(50), FamilyFirstName CHAR(50), FamilyMiddleInitial CHAR(1), FamilyCompanyOrganization CHAR(40), ClientIndexedName CHAR(160), ClientLastName CHAR(50), ClientFirstName CHAR(50), ClientMiddleInitial CHAR(1), Pronoun CHAR(3));", dbSeeChanges: Call BriefDelay
    TILLDataBase.Execute "INSERT INTO temptbl2 ( FamilyIndexedName, FamilyLastName, FamilyFirstName, FamilyMiddleInitial, FamilyCompanyOrganization, ClientIndexedName, ClientLastName, ClientFirstName, ClientMiddleInitial, Pronoun ) " & vbCrLf & _
        "SELECT tblPeopleFamily.IndexedName AS FamilyIndexedName, tblPeople.LastName AS FamilyLastName, tblPeople.FirstName AS FamilyFirstName, tblPeople.MiddleInitial AS FamilyMiddleInitial, tblPeople.CompanyOrganization AS FamilyCompanyOrganization, tblPeopleFamily.ClientIndexedName, temptbl1.LastName AS ClientLastName, temptbl1.FirstName AS ClientFirstName, temptbl1.MiddleInitial AS ClientMiddleInitial, IIf([tblPeopleClientsDemographics]![Gender]='M','his','her') AS Pronoun " & vbCrLf & _
        "FROM ((tblPeopleFamily INNER JOIN tblPeople ON tblPeopleFamily.IndexedName = tblPeople.IndexedName) INNER JOIN temptbl1 ON tblPeopleFamily.ClientIndexedName = temptbl1.IndexedName) INNER JOIN tblPeopleClientsDemographics ON temptbl1.IndexedName = tblPeopleClientsDemographics.IndexedName " & vbCrLf & _
        "WHERE tblPeople.IsDeceased=False AND tblPeopleFamily.Guardian=Yes AND tblPeopleFamily.Inactive=False AND " & _
            "(tblPeopleClientsDemographics.ActiveResidentialServices=True OR tblPeopleClientsDemographics.ActiveCLO=True OR tblPeopleClientsDemographics.ActiveSharedLiving=True OR tblPeopleClientsDemographics.ActiveAFC=True OR tblPeopleClientsDemographics.ActiveIndivSupport=True);", dbSeeChanges: Call BriefDelay
    TILLDataBase.Execute "CREATE TABLE temptbl3 (DisplayName CHAR(160), ClientIndexedName CHAR(160), ClientLastName CHAR(25), ClientFirstName CHAR(25), ClientMiddleInitial CHAR(1));", dbSeeChanges: Call BriefDelay
    TILLDataBase.Execute "INSERT INTO temptbl3 (DisplayName, ClientIndexedName, ClientLastName, ClientFirstName, ClientMiddleInitial) " & vbCrLf & _
        "SELECT DisplayName, IndexedName, LastName, FirstName, MiddleInitial " & vbCrLf & _
        "FROM AllClients WHERE IndexedName=""" & SelectedClientIndexedName & """;", dbSeeChanges: Call BriefDelay

    If FromAllConsentForms Then TRVLEnterEffectiveMonth = AllConsentFormsEnterEffectiveMonth
    If DLookup("LegalStatus", "tblPeopleClientsDemographics", "IndexedName=""" & SelectedClientIndexedName & """") = "Competent" Or DLookup("LegalStatus", "tblPeopleClientsDemographics", "IndexedName=""" & SelectedClientIndexedName & """") = "Presumed Competent" Then Comp = True

    If DCount("FamilyIndexedName", "temptbl2", "ClientIndexedName = """ & SelectedClientIndexedName & """") > 0 Then WhichLetter = "ltrTRVREL" Else WhichLetter = "ltrTRVRELSelfApproval"
    
    If (Not Comp) And DCount("FamilyIndexedName", "temptbl2", "ClientIndexedName = """ & SelectedClientIndexedName & """") > 0 Then
        ' Incompetent and has guardian(s).
        WhichLetter = "ltrTRVREL"
    ElseIf Comp Then
        ' Competent.
        WhichLetter = "ltrTRVRELSelfApproval"
    Else
        ' Incompetent but no guardians.  Flag as error and get out.
        MsgBox "Individual is flagged as incompetent but has no guardian(s).  Form cannot be produced.", vbOKOnly, "Error!"
        TRVLSelectClient = Null: TRVLEnterEffectiveMonth = Null: Me.Refresh: Me.Requery: Me.Repaint
        Exit Sub
    End If
    
    Call BriefDelay: DoCmd.OpenReport WhichLetter, acViewPreview, , "ClientIndexedName = """ & SelectedClientIndexedName & """": Call LoopUntilClosed(WhichLetter, 3)
    Call DropTempTables
    TRVLSelectClient = Null: TRVLEnterEffectiveMonth = Null: Me.Refresh: Me.Requery: Me.Repaint
    Exit Sub
ShowMeError:
    MsgBox "Error # " & Str(Err.Number) & " was generated by " & Me.Name & Chr(13) & Err.Description, vbOKOnly, "Error", Err.HelpFile, Err.HelpContext
End Sub

Private Sub TRVLEnterEffectiveMonth_AfterUpdate()
    If (Not IsNull(SelectedClientIndexedName)) And (Not IsNull(TRVLEnterEffectiveMonth)) Then Call TRVLDoReport
End Sub

Private Sub TRVLSelectClient_AfterUpdate()
    Call BuildAllClients(False)
    SelectedClientIndexedName = DLookup("IndexedName", "AllClients", "ReverseDisplayName = """ & TRVLSelectClient & """")
    If Len(SelectedClientIndexedName) > 0 And Len(TRVLEnterEffectiveMonth) > 0 Then Call TRVLDoReport
End Sub

Private Sub RepPayeeTILL_Click()
On Error GoTo ShowMeError
    Call DropTempTables
    TILLDataBase.Execute "CREATE TABLE temptbl0 (ClientIndexedName CHAR(160), ClientLastName CHAR(25), ClientFirstName CHAR(25), ClientMiddleInitial CHAR(1), " & _
        "ClientMailingAddress CHAR(50), ClientMailingCity CHAR(25), ClientMailingState CHAR(2), ClientMailingZIP CHAR(10), " & _
        "IsClientRes BIT, IsCilentCLO BIT, IsClientSharedLiving BIT, IsClientAFC BIT, IsClientIndiv BIT, Location CHAR(40), FamilyIndexedName CHAR(160), Relationship CHAR(25), Guardian BIT, PrimaryContact BIT, ClientSSN CHAR(11));", dbSeeChanges: Call BriefDelay
    
    TILLDataBase.Execute "INSERT INTO temptbl0 ( ClientIndexedName, ClientLastName, ClientFirstName, ClientMiddleInitial, ClientMailingAddress, ClientMailingCity, ClientMailingState, ClientMailingZIP, " & _
        "IsClientRes, IsCilentCLO, IsClientSharedLiving, IsClientAFC, IsClientIndiv, Location, FamilyIndexedName, Relationship, Guardian, PrimaryContact, ClientSSN ) " & _
        "SELECT tblPeople.IndexedName, tblPeople.LastName, tblPeople.FirstName, tblPeople.MiddleInitial, tblPeople.MailingAddress, tblPeople.MailingCity, tblPeople.MailingState, tblPeople.MailingZIP, tblPeople.IsClientRes, tblPeople.IsCilentCLO, tblPeople.IsClientSharedLiving, tblPeople.IsClientAFC, tblPeople.IsClientIndiv, IIf([tblPeople]![IsClientRes],[tblPeopleClientsResidentialServices]![CityTown] & '-' & [tblPeopleClientsResidentialServices]![Location],IIf([tblPeople]![IsClientSharedLiving],Null,IIf([tblPeople]![IsClientAFC],Null,IIf([tblPeople]![IsClientIndiv],Null,IIf([tblPeople]![IsCilentCLO],[tblPeopleClientsCLOServices]![CityTown] & '-' & [tblPeopleClientsCLOServices]![Location]))))) AS Location, tblPeopleFamily.IndexedName, tblPeopleFamily.Relationship, tblPeopleFamily.Guardian, tblPeopleFamily.PrimaryContact, tblPeopleClientsDemographics.SocialSecurityNumber " & _
        "FROM ((((((tblPeople INNER JOIN tblPeopleClientsDemographics ON tblPeople.IndexedName = tblPeopleClientsDemographics.IndexedName) LEFT JOIN tblPeopleFamily ON tblPeopleClientsDemographics.IndexedName = tblPeopleFamily.ClientIndexedName) LEFT JOIN tblPeopleClientsCLOServices ON tblPeopleClientsDemographics.IndexedName = tblPeopleClientsCLOServices.IndexedName) LEFT JOIN tblPeopleClientsIndividualSupportServices ON tblPeopleClientsDemographics.IndexedName = tblPeopleClientsIndividualSupportServices.IndexedName) LEFT JOIN tblPeopleClientsResidentialServices ON tblPeopleClientsDemographics.IndexedName = tblPeopleClientsResidentialServices.IndexedName) LEFT JOIN tblPeopleClientsSharedLivingServices ON tblPeopleClientsDemographics.IndexedName = tblPeopleClientsSharedLivingServices.IndexedName) LEFT JOIN tblPeopleClientsAFCServices ON tblPeopleClientsDemographics.IndexedName = tblPeopleClientsAFCServices.IndexedName " & _
        "WHERE (((tblPeopleFamily.Guardian)=True) AND ((tblPeopleClientsDemographics.ActiveResidentialServices)=True) AND ((tblPeopleClientsResidentialServices.Inactive)=False) AND ((tblPeopleFamily.Inactive)=False) AND ((tblPeople.IsDeceased)=False) AND ((tblPeopleClientsDemographics.RepresentativePayee) Like '*TILL*')) OR (((tblPeopleFamily.Guardian)=True) AND ((tblPeopleFamily.Inactive)=False) AND ((tblPeople.IsDeceased)=False) AND ((tblPeopleClientsDemographics.RepresentativePayee) Like '*TILL*') AND ((tblPeopleClientsDemographics.ActiveSharedLiving)=True) AND ((tblPeopleClientsSharedLivingServices.Inactive)=False)) OR (((tblPeopleFamily.Guardian)=True) AND ((tblPeopleFamily.Inactive)=False) AND ((tblPeople.IsDeceased)=False) AND ((tblPeopleClientsDemographics.RepresentativePayee) Like '*TILL*') AND ((tblPeopleClientsDemographics.ActiveAFC)=True) AND ((tblPeopleClientsAFCServices.Inactive)=False)) OR (((tblPeopleFamily.Guardian)=True) AND " & _
        "((tblPeopleFamily.Inactive)=False) AND ((tblPeople.IsDeceased)=False) AND ((tblPeopleClientsDemographics.RepresentativePayee) Like '*TILL*') AND ((tblPeopleClientsDemographics.ActiveCLO)=True) AND ((tblPeopleClientsCLOServices.Inactive)=False)) OR (((tblPeopleFamily.PrimaryContact)=True) AND ((tblPeopleClientsDemographics.ActiveResidentialServices)=True) AND ((tblPeopleClientsResidentialServices.Inactive)=False) AND ((tblPeopleFamily.Inactive)=False) AND ((tblPeople.IsDeceased)=False) AND ((tblPeopleClientsDemographics.RepresentativePayee) Like '*TILL*')) OR (((tblPeopleFamily.PrimaryContact)=True) AND ((tblPeopleFamily.Inactive)=False) AND ((tblPeople.IsDeceased)=False) AND ((tblPeopleClientsDemographics.RepresentativePayee) Like '*TILL*') AND ((tblPeopleClientsDemographics.ActiveSharedLiving)=True) AND ((tblPeopleClientsSharedLivingServices.Inactive)=False)) OR (((tblPeopleFamily.PrimaryContact)=True) AND ((tblPeopleFamily.Inactive)=False) AND ((tblPeople.IsDeceased)=False) AND " & _
        "((tblPeopleClientsDemographics.RepresentativePayee) Like '*TILL*') AND ((tblPeopleClientsDemographics.ActiveAFC)=True) AND ((tblPeopleClientsAFCServices.Inactive)=False)) OR (((tblPeopleFamily.PrimaryContact)=True) AND ((tblPeopleFamily.Inactive)=False) AND ((tblPeople.IsDeceased)=False) AND ((tblPeopleClientsDemographics.RepresentativePayee) Like '*TILL*') AND ((tblPeopleClientsDemographics.ActiveCLO)=True) AND ((tblPeopleClientsCLOServices.Inactive)=False)) OR (((tblPeopleFamily.Guardian)=True) AND ((tblPeopleFamily.Inactive)=False) AND ((tblPeople.IsDeceased)=False) AND ((tblPeopleClientsDemographics.RepresentativePayee) Like '*TILL*') AND ((tblPeopleClientsDemographics.ActiveIndivSupport)=True) AND ((tblPeopleClientsIndividualSupportServices.Inactive)=False)) OR (((tblPeopleFamily.PrimaryContact)=True) AND ((tblPeopleFamily.Inactive)=False) AND ((tblPeople.IsDeceased)=False) AND ((tblPeopleClientsDemographics.RepresentativePayee) Like '*TILL*') AND " & _
        "((tblPeopleClientsDemographics.ActiveIndivSupport)=True) AND ((tblPeopleClientsIndividualSupportServices.Inactive)=False));", dbSeeChanges: Call BriefDelay
    
    TILLDataBase.Execute "CREATE TABLE temptbl1 (ClientIndexedName CHAR(160), ClientLastName CHAR(25), ClientFirstName CHAR(25), " & _
        "ClientMiddleInitial CHAR(1), ClientMailingAddress CHAR(50), " & _
        "ClientMailingCity CHAR(25), ClientMailingState CHAR(2), ClientMailingZIP CHAR(10), ClientSSN CHAR(11), " & _
        "IsClientRes BIT, IsCilentCLO BIT, IsClientSharedLiving BIT, IsClientAFC BIT, IsClientIndiv BIT, Location CHAR(40), " & _
        "FamilyIndexedName CHAR(160), Relationship CHAR(25), LN CHAR(25), FN CHAR(25), " & _
        "MI CHAR(1), CO CHAR(40), FamilyFamiliarGreeting CHAR(40), MailingAddress CHAR(50), " & _
        "MailingCity CHAR(25), MailingState CHAR(2), MailingZIP CHAR(10), GoGreen BIT, " & _
        "EmailAddress CHAR(50));", dbSeeChanges: Call BriefDelay
    
    TILLDataBase.Execute "INSERT INTO temptbl1 ( ClientIndexedName, ClientLastName, ClientFirstName, ClientMiddleInitial, " & _
        "ClientMailingAddress, ClientMailingCity, ClientMailingState, ClientMailingZIP, ClientSSN, " & _
        "IsClientRes, IsCilentCLO, IsClientSharedLiving, IsClientAFC, IsClientIndiv, Location, FamilyIndexedName, Relationship, LN, FN, " & _
        "MI, CO, FamilyFamiliarGreeting, MailingAddress, MailingCity, MailingState, " & _
        "MailingZIP, GoGreen, EmailAddress ) " & _
        "SELECT temptbl0.ClientIndexedName, temptbl0.ClientLastName, temptbl0.ClientFirstName, temptbl0.ClientMiddleInitial, " & _
        "temptbl0.ClientMailingAddress, temptbl0.ClientMailingCity, temptbl0.ClientMailingState, temptbl0.ClientMailingZIP, temptbl0.ClientSSN, " & _
        "temptbl0.IsClientRes, temptbl0.IsCilentCLO, temptbl0.IsClientSharedLiving, temptbl0.IsClientAFC, temptbl0.IsClientIndiv, temptbl0.Location, temptbl0.FamilyIndexedName, temptbl0.Relationship, " & _
        "tblPeople.LastName AS LN, tblPeople.FirstName AS FN, " & _
        "tblPeople.MiddleInitial AS MI, tblPeople.CompanyOrganization AS CO, " & _
        "tblPeople.FamiliarGreeting AS FamilyFamiliarGreeting, tblPeople.MailingAddress AS MailingAddress, tblPeople.MailingCity AS MailingCity, " & _
        "tblPeople.MailingState AS MailingState, tblPeople.MailingZIP AS MailingZIP, tblPeople.GoGreen, " & _
        "tblPeople.EmailAddress " & _
        "FROM temptbl0 LEFT JOIN tblPeople ON temptbl0.FamilyIndexedName = tblPeople.IndexedName " & _
        "WHERE tblPeople.IsDeceased=FALSE AND tblPeople.LastName Is Not Null AND " & _
        "(temptbl0.Guardian=True OR temptbl0.PrimaryContact=True);", dbSeeChanges: Call BriefDelay
    
    TILLDataBase.Execute "CREATE TABLE temptbl2 (MailingAddress CHAR(50), MailingCity CHAR(25), " & _
        "MailingState CHAR(2), MailingZIP CHAR(10), ClientIndexedName CHAR(160), ClientMailingAddress CHAR(50), " & _
        "ClientMailingCity CHAR(25), ClientMailingState CHAR(2), ClientMailingZIP CHAR(10), ClientSSN CHAR(11), " & _
        "IsClientRes BIT, IsCilentCLO BIT, " & _
        "IsClientSharedLiving BIT, IsClientAFC BIT, IsClientIndiv BIT, Location CHAR(40), LN1 CHAR(25), FN1 CHAR(25), MI1 CHAR(1), " & _
        "CO1 CHAR(50), FamilyFamiliarGreeting1 CHAR(30), FamilyGoGreen1 BIT, EmailAddress1 CHAR(50), LN2 CHAR(25), " & _
        "FN2 CHAR(25), MI2 CHAR(1), CO2 CHAR(50), FamilyFamiliarGreeting2 CHAR(30), FamilyGoGreen2 BIT, " & _
        "EmailAddress2 CHAR(50), LN3 CHAR(25), FN3 CHAR(25), MI3 CHAR(1), " & _
        "CO3 CHAR(50), FamilyFamiliarGreeting3 CHAR(30), FamilyGoGreen3 BIT, EmailAddress3 CHAR(50), LN4 CHAR(25), " & _
        "FN4 CHAR(25), MI4 CHAR(1), CO4 CHAR(50), FamilyFamiliarGreeting4 CHAR(30), FamilyGoGreen4 BIT, " & _
        "EmailAddress4 CHAR(50));", dbSeeChanges: Call BriefDelay
    
    TILLDataBase.Execute "ALTER TABLE temptbl2 ADD CONSTRAINT temptbl2constraint PRIMARY KEY (MailingAddress,MailingCity,MailingState,MailingZIP,ClientIndexedName);", dbSeeChanges: Call BriefDelay
    
    TILLDataBase.Execute "INSERT INTO temptbl2 (MailingAddress, MailingCity, MailingState, MailingZIP, " & _
        "ClientIndexedName, ClientMailingAddress, ClientMailingCity, ClientMailingState, ClientMailingZIP, ClientSSN, IsClientRes, IsCilentCLO, IsClientSharedLiving, IsClientAFC, IsClientIndiv, Location )" & _
        "SELECT MailingAddress, MailingCity, MailingState, MailingZIP, ClientIndexedName, ClientMailingAddress, ClientMailingCity, ClientMailingState, ClientMailingZIP, ClientSSN, " & _
        "IsClientRes, IsCilentCLO, IsClientSharedLiving, IsClientAFC, IsClientIndiv, Location FROM temptbl1 " & _
        "WHERE IsClientRes=True OR IsCilentCLO=True OR IsClientSharedLiving=True OR IsClientAFC=True OR IsClientIndiv=True;", dbSeeChanges: Call BriefDelay
    
    Call ParseFamilyMembersAndThenExport(0, "temptbl2", "RepPayeesTILL")
    SysCmdResult = SysCmd(5)
    Exit Sub
ShowMeError:
    MsgBox "Error # " & Str(Err.Number) & " was generated by " & Me.Name & Chr(13) & Err.Description, vbOKOnly, "Error", Err.HelpFile, Err.HelpContext
End Sub

Private Sub RepPayeeNotTILL_Click()
    Call BasicExport("SELECT tblPeople.FirstName & ' ' & tblPeople.MiddleInitial & ' ' & tblPeople.LastName AS ClientName, tblPeopleClientsDemographics.SocialSecurityNumber AS ClientSSN, tblPeopleClientsDemographics.RepresentativePayee, tblPeopleClientsDemographics.RepPayeeAddress, tblPeopleClientsDemographics.RepPayeeCity, tblPeopleClientsDemographics.RepPayeeState, tblPeopleClientsDemographics.RepPayeeZIP, " & _
        "IIf([tblPeople]![IsClientRes],[tblPeopleClientsResidentialServices]![CityTown] & '-' & [tblPeopleClientsResidentialServices]![Location],IIf([tblPeople]![IsClientSharedLiving],'SL',IIf([tblPeople]![IsClientIndiv],'ISS',IIf([tblPeople]![IsCilentCLO],[tblPeopleClientsCLOServices]![CityTown] & '-' & [tblPeopleClientsCLOServices]![Location])))) AS Location " & _
        "INTO temptbl " & _
        "FROM (((((tblPeople INNER JOIN tblPeopleClientsDemographics ON tblPeople.IndexedName = tblPeopleClientsDemographics.IndexedName) LEFT JOIN tblPeopleClientsResidentialServices ON tblPeopleClientsDemographics.IndexedName = tblPeopleClientsResidentialServices.IndexedName) LEFT JOIN tblPeopleClientsSharedLivingServices ON tblPeopleClientsDemographics.IndexedName = tblPeopleClientsSharedLivingServices.IndexedName) LEFT JOIN tblPeopleClientsAFCServices ON tblPeopleClientsDemographics.IndexedName = tblPeopleClientsAFCServices.IndexedName) LEFT JOIN tblPeopleClientsCLOServices ON tblPeopleClientsDemographics.IndexedName = tblPeopleClientsCLOServices.IndexedName) LEFT JOIN tblPeopleClientsIndividualSupportServices ON tblPeopleClientsDemographics.IndexedName = tblPeopleClientsIndividualSupportServices.IndexedName " & _
        "WHERE tblPeople.IsDeceased=False AND " & _
        " (tblPeopleClientsDemographics.RepresentativePayee Not Like '*TILL*' OR tblPeopleClientsDemographics.RepPayeeIsClient=True) AND " & _
        "((tblPeopleClientsDemographics.ActiveResidentialServices=True AND tblPeopleClientsResidentialServices.Inactive=False)  OR " & _
        " (tblPeopleClientsDemographics.ActiveSharedLiving=True        AND tblPeopleClientsSharedLivingServices.Inactive=False) OR " & _
        " (tblPeopleClientsDemographics.ActiveCLO=True                 AND tblPeopleClientsCLOServices.Inactive=False)          OR " & _
        " (tblPeopleClientsDemographics.ActiveIndivSupport=True        AND tblPeopleClientsIndividualSupportServices.Inactive=False));", "RepPayeesNotTILL")
End Sub

Private Sub UpdateProgressMessages(Message As String, Optional FirstMessage As Boolean = False, Optional LastMessage As Boolean = False)
    If FirstMessage Then
        ProgressMessages = ""
        ProgressMessages = ProgressMessages & Message & vbCrLf: ProgressMessages.Requery
    ElseIf LastMessage Then
        ProgressMessages = ProgressMessages & Message & vbCrLf: ProgressMessages.Requery
        Call BriefDelay(3)
        ProgressMessages = ""
    Else
        ProgressMessages = ProgressMessages & Message & vbCrLf: ProgressMessages.Requery
    End If
End Sub