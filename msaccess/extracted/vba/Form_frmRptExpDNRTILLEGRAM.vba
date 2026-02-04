' Module Name: Form_frmRptExpDNRTILLEGRAM
' Module Type: Document Module
' Lines of Code: 70
' Extracted: 2026-02-04 13:03:35

Option Compare Database
Option Explicit

Private RetValue As Variant, CommandLine As String
Private Const MAX_PATH As Long = 260
'Private Declare Function FindExecutable Lib "shell32" Alias "FindExecutableA" ( _
'  ByVal lpFile As String, _
'  ByVal lpDirectory As String, _
'  ByVal sResult As String _
') As Long

Private Sub Form_Load()
    Select Case Month(Now())
        Case 1:   StartingMonth = "JAN": Case 2:   StartingMonth = "FEB": Case 3:   StartingMonth = "MAR"
        Case 4:   StartingMonth = "APR": Case 5:   StartingMonth = "MAY": Case 6:   StartingMonth = "JUN"
        Case 7:   StartingMonth = "JUL": Case 8:   StartingMonth = "AUG": Case 9:   StartingMonth = "SEP"
        Case 10:  StartingMonth = "OCT": Case 11:  StartingMonth = "NOV": Case 12:  StartingMonth = "DEC"
    End Select
    StartingYear = Year(Now()): EndingYear = Year(Now()): EndingMonth = StartingMonth
End Sub

Private Function MonthNum(MonthAbbrev As String) As String
    Select Case MonthAbbrev
        Case "JAN": MonthNum = "01": Case "FEB": MonthNum = "02": Case "MAR": MonthNum = "03"
        Case "APR": MonthNum = "04": Case "MAY": MonthNum = "05": Case "JUN": MonthNum = "06"
        Case "JUL": MonthNum = "07": Case "AUG": MonthNum = "08": Case "SEP": MonthNum = "09"
        Case "OCT": MonthNum = "10": Case "NOV": MonthNum = "11": Case "DEC": MonthNum = "12"
    End Select
End Function

Private Sub OK_Click()
On Error GoTo ShowMeError
    Dim StartComp As Date, StartStr As String, EndStr As String, EndComp As Date, ExportFileName As String

    StartComp = CDate(MonthNum(StartingMonth) & "/01/" & CStr(StartingYear))
    If MonthNum(EndingMonth) = 12 Then EndComp = "01/01/" & CStr(EndingYear + 1) Else EndComp = CDate(CStr(CInt(MonthNum(EndingMonth)) + 1) & "/01/" & CStr(EndingYear))
    
    If StartComp > EndComp Then
        MsgBox "The selected starting date is greater than the select ending date.  Please correct.", vbOKOnly, "Error!"
    Else
        Call DropTempTables
        StartDate = DateValue(CStr(CInt(MonthNum(StartingMonth))) & "/01/" & CStr(StartingYear))
        If MonthNum(EndingMonth) = 12 Then EndDate = DateValue("01/01/" & CStr(EndingYear + 1)) Else EndDate = DateValue(CStr(CInt(MonthNum(EndingMonth)) + 1) & "/01/" & CStr(EndingYear))
        StartStr = "#" & CStr(CInt(MonthNum(StartingMonth))) & "/01/" & CStr(StartingYear) & "#"
        If MonthNum(EndingMonth) = 12 Then EndStr = "#" & "01/01/" & CStr(EndingYear + 1) & "#" Else EndStr = "#" & CStr(CInt(MonthNum(EndingMonth)) + 1) & "/01/" & CStr(EndingYear) & "#"
        TILLDataBase.Execute "SELECT DateValue([DateOfDonation]) AS DonationDate, DateValue([DateReceived]) AS ReceivedDate, tblPeopleDonors.IndexedName AS IndexedName, tblPeopleDonors.DonationType AS DonationType, tblPeopleDonors.SolicitationType AS SolicitationType, tblPeopleDonors.IsGrant AS IsGrant, tblPeopleDonors.DonationFrom1Salutation AS Donor1Sal, tblPeopleDonors.DonationFrom1FirstName AS Donor1FN, tblPeopleDonors.DonationFrom1LastName AS Donor1LN, tblPeopleDonors.DonationFrom2Salutation AS Donor2Sal, tblPeopleDonors.DonationFrom2FirstName AS Donor2FN, tblPeopleDonors.DonationFrom2LastName AS Donor2LN, tblPeopleDonors.DonationFromCompany AS DonorCompany, tblPeopleDonors.Description AS Description, tblPeopleDonors.Amount AS Amount" & vbCrLf & _
            "INTO temptbl " & vbCrLf & _
            "FROM tblPeopleDonors INNER JOIN tblPeople ON tblPeopleDonors.IndexedName = tblPeople.IndexedName" & vbCrLf & _
            "WHERE DateValue([DateOfDonation]) >=" & StartStr & " And DateValue([DateOfDonation]) <" & EndStr & vbCrLf & _
            "ORDER BY DateValue([DateOfDonation]);", dbSeeChanges: Call BriefDelay
        ExportFileName = Application.CurrentProject.Path & "\" & "TILLDB-Export-DonorListForTILLEGram-" & Format(Date, "yyyymmdd") & ".xls"
        If IsFileOpen(ExportFileName) Then
            If MsgBox(ExportFileName & " is already open.  Please close it and click OK to continue or Cancel to abort.", vbOKCancel, "ERROR!") = vbCancel Then
                MsgBox "Export aborted.", vbOKOnly, "Aborted"
                Exit Sub
            End If
        End If
        If Dir(ExportFileName) <> "" Then Kill ExportFileName
        DoCmd.TransferSpreadsheet acExport, acSpreadsheetTypeExcel9, "temptbl", ExportFileName
'       CommandLine = LocateExecutable(ExportFileName) & " """ & ExportFileName & """"
        MsgBox "The requested information has been exported to " & ExportFileName & "." & vbCrLf & vbCrLf & "This export may contain information that is protected under HIPAA and other privacy laws.  This export must be securely stored at all times and must be deleted when no longer being used.", _
            vbOKOnly, "Export Complete"
'       RetValue = Shell(CommandLine, 1)
        Call DropTempTables
        DoCmd.Close
    End If
    Exit Sub
ShowMeError:
    MsgBox "Error # " & Str(Err.Number) & " was generated by " & Me.Name & Chr(13) & Err.Description, vbOKOnly, "Error", Err.HelpFile, Err.HelpContext
End Sub
