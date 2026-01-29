' Module Name: Form_frmContractsBillingBook
' Module Type: Document Module
' Lines of Code: 123
' Extracted: 1/29/2026 4:12:27 PM

Option Compare Database
Option Explicit

Private Sub Form_Current()
    If Not Form_frmContractsBillingBook.NewRecord Then
        MaximumObligationAsAmended = DSum("AmendedAmount", "tblContractsAmendments", "ContractID = """ & ContractID & """ AND FY = " & Form_frmContracts.SelectFY & " AND BillingBookNumber = " & BillingBookNumber)
        If IsNull(MaximumObligationAsAmended) Then MaximumObligationAsAmended = MaximumObligation Else MaximumObligationAsAmended = MaximumObligationAsAmended + MaximumObligation
        UnitsAsAmended = DSum("NumUnits", "tblContractsAmendments", "ContractID = """ & ContractID & """ AND FY = " & Form_frmContracts.SelectFY & " AND BillingBookNumber = " & BillingBookNumber)
        If IsNull(UnitsAsAmended) Then UnitsAsAmended = Units Else UnitsAsAmended = UnitsAsAmended + Units
    End If
    
    If Not (IsNull(R01) Or IsNull(U01)) Then T01 = U01 * R01
    If Not (IsNull(R02) Or IsNull(U02)) Then T02 = U02 * R02
    If Not (IsNull(R03) Or IsNull(U03)) Then T03 = U03 * R03
    If Not (IsNull(R04) Or IsNull(U04)) Then T04 = U04 * R04
    If Not (IsNull(R05) Or IsNull(U05)) Then T05 = U05 * R05
    If Not (IsNull(R06) Or IsNull(U06)) Then T06 = U06 * R06
    If Not (IsNull(R07) Or IsNull(U07)) Then T07 = U07 * R07
    If Not (IsNull(R08) Or IsNull(U08)) Then T08 = U08 * R08
    If Not (IsNull(R09) Or IsNull(U09)) Then T09 = U09 * R09
    If Not (IsNull(R10) Or IsNull(U10)) Then T10 = U10 * R10
    If Not (IsNull(R11) Or IsNull(U11)) Then T11 = U11 * R11
    If Not (IsNull(R12) Or IsNull(U12)) Then T12 = U12 * R12
    
    Call URAfterUpdate(0)
End Sub

Private Sub BIllingBookNumber_AfterUpdate()
    If Form_frmContractsBillingBook.NewRecord Then
        RecordAddedDate = Format(Now(), "mm/dd/yyyy"): RecordAddedBy = Form_frmMainMenu.UserName
    End If
End Sub

Private Sub ProcessDeleteRecord_Click()
On Error GoTo ShowMeError
    If MsgBox("Do you really want to delete this entire billing book entry for FY " & FY & " including all amendments?", vbYesNo, "Confirm Deletion") = vbYes Then
        If MsgBox("Are you absolutely sure you want to delete this entire billing book entry for FY " & FY & " including all amendments?", vbYesNo, "Confirm Deletion") = vbYes Then
            ' Delete amendments.
            TILLDataBase.Execute "INSERT INTO tblDELETEDContractsAmendments ( FY, ContractID, AmendmentNumber, RecordDeletedDate, RecordDeletedBy, " & _
                "BillingBookNumber, Pending, PendingAmount, PendingUnits, DDSArea, ProgramName, PurposeReason, DateSubmitted, DateApproved, StartDate, " & _
                "EndDate, NumUnits, NewRate, AmendedAmount, AnnualizedAmount, Comments ) " & _
                "SELECT FY, ContractID, AmendmentNumber, """ & _
                Format(Now(), "mm/dd/yyyy") & """ AS RecordDeletedDate, """ & _
                Form_frmMainMenu.UserName & """ AS RecordDeletedBy, " & _
                "BillingBookNumber, Pending, PendingAmount, PendingUnits, DDSArea, ProgramName, PurposeReason, DateSubmitted, DateApproved, StartDate, " & _
                "EndDate, NumUnits, NewRate, AmendedAmount, AnnualizedAmount, Comments " & _
                "FROM tblContractsAmendments " & _
                "WHERE FY=" & Form_frmContractsBillingBook.FY & " AND " & "BillingBookNumber=" & Form_frmContractsBillingBook.BillingBookNumber, dbSeeChanges: Call BriefDelay
            TILLDataBase.Execute "DELETE * FROM tblContractsAmendments " & _
                "WHERE FY=" & Form_frmContractsBillingBook.FY & " AND " & "BillingBookNumber=" & Form_frmContractsBillingBook.BillingBookNumber, dbSeeChanges: Call BriefDelay    ' Delete the amendments.
            ' Delete billing book.
            TILLDataBase.Execute "INSERT INTO tblDELETEDContractsBillingBook ( FY, ContractID, BIllingBookNumber, RecordDeletedDate, RecordDeletedBy, " & _
                "ProgramName, CostCenter, StartDate, EndDate, 07Units, 07Rate, 08Units, 08Rate, 09Units, 09Rate, 10Units, 10Rate, 11Units, 11Rate, " & _
                "12Units, 12Rate, 01Units, 01Rate, 02Units, 02Rate, 03Units, 03Rate, 04Units, 04Rate, 05Units, 05Rate, 06Units, 06Rate, MaximumObligation, " & _
                "MaximumObligationAsAmended, Units, UnitsAsAmended, BillingRate, NumberOfClients, InternalRate, FundingSource, DDSArea, Staff, " & _
                "Comments )" & _
                "SELECT FY, ContractID, BIllingBookNumber, """ & _
                Format(Now(), "mm/dd/yyyy") & """ AS RecordDeletedDate, """ & _
                Form_frmMainMenu.UserName & """ AS RecordDeletedBy, " & _
                "ProgramName, CostCenter, StartDate, EndDate, " & _
                "[07Units], [07Rate], [08Units], [08Rate], [09Units], [09Rate], [10Units], [10Rate], [11Units], [11Rate], [12Units], [12Rate], " & _
                "[01Units], [01Rate], [02Units], [02Rate], [03Units], [03Rate], [04Units], [04Rate], [05Units], [05Rate], [06Units], [06Rate], " & _
                "MaximumObligation, MaximumObligationAsAmended, Units, UnitsAsAmended, BillingRate, NumberOfClients, InternalRate, FundingSource, DDSArea, " & _
                "Staff, Comments FROM tblContractsBillingBook " & _
                "WHERE FY=" & Form_frmContractsBillingBook.FY & " AND " & "BillingBookNumber=" & Form_frmContractsBillingBook.BillingBookNumber, dbSeeChanges: Call BriefDelay
            TILLDataBase.Execute "DELETE * FROM tblContractsBillingBook " & _
                "WHERE " & "FY=" & Form_frmContractsBillingBook.FY & " AND " & "BillingBookNumber=" & Form_frmContractsBillingBook.BillingBookNumber, dbSeeChanges: Call BriefDelay   ' Delete the billing book entries.
            DoCmd.GoToRecord , , acFirst
        End If
    End If
    Exit Sub
ShowMeError:
    MsgBox "Error # " & Str(Err.Number) & " was generated by " & Me.Name & Chr(13) & Err.Description, vbOKOnly, "Error", Err.HelpFile, Err.HelpContext
End Sub

Public Function RecalcContractsInBillingBook() As Boolean
    RecalcContractsInBillingBook = True
    With Form_frmContracts
        If IsNull(.ContractID) Or IsNull(.SelectFY) Or .SelectFY < 2000 Then GoTo ProceedNoAction
        .TotalClients = DSum("NumberOfClients", "tblContractsBillingBook", "ContractID = """ & .ContractID & """ AND FY = " & .SelectFY)
        .TotalUnits = DSum("Units", "tblContractsBillingBook", "ContractID = """ & .ContractID & """ AND FY = " & .SelectFY)
        .TotalUnitsAsAmended = DSum("UnitsAsAmended", "tblContractsBillingBook", "ContractID = """ & .ContractID & """ AND FY = " & .SelectFY)
        .MaximumObligation = DSum("MaximumObligation", "tblContractsBillingBook", "ContractID = """ & .ContractID & """ AND FY = " & .SelectFY)
        .MaximumObligationAsAmended = DSum("MaximumObligationAsAmended", "tblContractsBillingBook", "ContractID = """ & .ContractID & """ AND FY = " & .SelectFY)
        .Refresh
    End With
ProceedNoAction:
    Me.Refresh
End Function

Private Function URAfterUpdate(WhichMonth As Integer) As Boolean
    Select Case WhichMonth
        Case 1:  If Not (IsNull(R01) Or IsNull(U01)) Then T01 = U01 * R01
        Case 2:  If Not (IsNull(R02) Or IsNull(U02)) Then T02 = U02 * R02
        Case 3:  If Not (IsNull(R03) Or IsNull(U03)) Then T03 = U03 * R03
        Case 4:  If Not (IsNull(R04) Or IsNull(U04)) Then T04 = U04 * R04
        Case 5:  If Not (IsNull(R05) Or IsNull(U05)) Then T05 = U05 * R05
        Case 6:  If Not (IsNull(R06) Or IsNull(U06)) Then T06 = U06 * R06
        Case 7:  If Not (IsNull(R07) Or IsNull(U07)) Then T07 = U07 * R07
        Case 8:  If Not (IsNull(R08) Or IsNull(U08)) Then T08 = U08 * R08
        Case 9:  If Not (IsNull(R09) Or IsNull(U09)) Then T09 = U09 * R09
        Case 10: If Not (IsNull(R10) Or IsNull(U10)) Then T10 = U10 * R10
        Case 11: If Not (IsNull(R11) Or IsNull(U11)) Then T11 = U11 * R11
        Case 12: If Not (IsNull(R12) Or IsNull(U12)) Then T12 = U12 * R12
        Case Else ' WhichMonth = 0 means just skip to calculate the YTD totals.
    End Select
    
    YTDTotal = 0
    If Not (IsNull(R01) Or IsNull(U01)) Then YTDTotal = YTDTotal + T01
    If Not (IsNull(R02) Or IsNull(U02)) Then YTDTotal = YTDTotal + T02
    If Not (IsNull(R03) Or IsNull(U03)) Then YTDTotal = YTDTotal + T03
    If Not (IsNull(R04) Or IsNull(U04)) Then YTDTotal = YTDTotal + T04
    If Not (IsNull(R05) Or IsNull(U05)) Then YTDTotal = YTDTotal + T05
    If Not (IsNull(R06) Or IsNull(U06)) Then YTDTotal = YTDTotal + T06
    If Not (IsNull(R07) Or IsNull(U07)) Then YTDTotal = YTDTotal + T07
    If Not (IsNull(R08) Or IsNull(U08)) Then YTDTotal = YTDTotal + T08
    If Not (IsNull(R09) Or IsNull(U09)) Then YTDTotal = YTDTotal + T09
    If Not (IsNull(R10) Or IsNull(U10)) Then YTDTotal = YTDTotal + T10
    If Not (IsNull(R11) Or IsNull(U11)) Then YTDTotal = YTDTotal + T11
    If Not (IsNull(R12) Or IsNull(U12)) Then YTDTotal = YTDTotal + T12
    
    URAfterUpdate = True
End Function