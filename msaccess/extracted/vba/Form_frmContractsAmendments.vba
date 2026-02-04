' Module Name: Form_frmContractsAmendments
' Module Type: Document Module
' Lines of Code: 94
' Extracted: 2026-02-04 13:03:36

Option Compare Database
Option Explicit

Private Sub Form_Current()
    ' This code establishes what is visible in the form display and what color certain items should be.
    Pending.Visible = True
    If Not Form_frmContractsAmendments.NewRecord Then
        If Pending Then
            PendingAmount.Visible = True:  PendingUnits.Visible = True
        Else
            PendingAmount.Visible = False: PendingUnits.Visible = False
            RecordAddedBy = Form_frmMainMenu.UserName: RecordAddedDate = Format(Now, "mm/dd/yyyy")
        End If
        
        If IsNull(AmendedAmount) Then PendingUnits.Visible = True Else PendingUnits.Visible = False
        If NumUnits < 0 Then NumUnits.ForeColor = RGB(255, 0, 0) Else NumUnits.ForeColor = RGB(0, 0, 0)
        If AmendedAmount < 0 Then AmendedAmount.ForeColor = RGB(255, 0, 0) Else AmendedAmount.ForeColor = RGB(0, 0, 0)
    End If
End Sub

Private Sub Form_Dirty(Cancel As Integer)
    BillingBookNumber = Form_frmContractsBillingBook.BillingBookNumber
End Sub

Private Sub Pending_Click()
    PendingAmount.Visible = True
    If Pending Then
        PendingUnits.Visible = True: AmendedAmount.Enabled = False: NumUnits.Enabled = False
    Else
        If MsgBox("You are about to move this amendment to approved status.  Are you sure you want to do this?", vbYesNo, "NOTE") = vbYes Then
            PendingAmount.Visible = False: PendingUnits.Visible = False: AmendedAmount.Enabled = True
            NumUnits.Enabled = True: AmendedAmount = PendingAmount: NumUnits = PendingUnits
            Call RecalcContractsInAmendments
            PendingAmount = Null: PendingUnits = Null: DateApproved.SetFocus
            MsgBox "Be sure to set the 'Date Approved' field.", , "NOTE"
        Else
            Pending = True: PendingAmount.Visible = True: PendingUnits.Visible = True: AmendedAmount.Enabled = False: NumUnits.Enabled = False
        End If
    End If
End Sub

Public Function RecalcContractsInAmendments() As Boolean
    With Form_frmContractsBillingBook
        If IsNull(ContractID) Or IsNull(.BillingBookNumber) Then GoTo ProceedNoAction
        .MaximumObligationAsAmended = DSum("AmendedAmount", "tblContractsAmendments", "ContractID = """ & ContractID & """ AND FY = " & Form_frmContracts.SelectFY & " AND BillingBookNumber = " & .BillingBookNumber)
        If IsNull(.MaximumObligationAsAmended) Then .MaximumObligationAsAmended = .MaximumObligation Else .MaximumObligationAsAmended = .MaximumObligationAsAmended + .MaximumObligation
        .UnitsAsAmended = DSum("NumUnits", "tblContractsAmendments", "ContractID = """ & ContractID & """ AND FY = " & Form_frmContracts.SelectFY & " AND BillingBookNumber = " & .BillingBookNumber)
        If IsNull(.UnitsAsAmended) Then .UnitsAsAmended = .Units Else .UnitsAsAmended = .UnitsAsAmended + .Units
        .Refresh
    End With
    Call BriefDelay
    With Form_frmContracts
        If IsNull(.ContractID) Or IsNull(Form_frmContracts.SelectFY) Or Form_frmContracts.SelectFY < 2000 Then GoTo ProceedNoAction
        .TotalClients = DSum("NumberOfClients", "tblContractsBillingBook", "ContractID = """ & .ContractID & """ AND FY = " & Form_frmContracts.SelectFY)
        .TotalUnits = DSum("Units", "tblContractsBillingBook", "ContractID = """ & .ContractID & """ AND FY = " & Form_frmContracts.SelectFY)
        .TotalUnitsAsAmended = DSum("UnitsAsAmended", "tblContractsBillingBook", "ContractID = """ & .ContractID & """ AND FY = " & Form_frmContracts.SelectFY)
        .MaximumObligation = DSum("MaximumObligation", "tblContractsBillingBook", "ContractID = """ & .ContractID & """ AND FY = " & Form_frmContracts.SelectFY)
        .MaximumObligationAsAmended = DSum("MaximumObligationAsAmended", "tblContractsBillingBook", "ContractID = """ & .ContractID & """ AND FY = " & Form_frmContracts.SelectFY)
        .Refresh
    End With
    Call BriefDelay
ProceedNoAction:
    Me.Refresh
End Function

Private Sub AmendedAmount_AfterUpdate()
    If AmendedAmount < 0 Then AmendedAmount.ForeColor = RGB(255, 0, 0) Else AmendedAmount.ForeColor = RGB(0, 0, 0)
    Call RecalcContractsInAmendments
End Sub

Private Sub ProcessDeleteRecord_Click()
On Error GoTo ShowMeError
    Dim DeletionDate As Variant
    
    DeletionDate = Format(Now(), "mm/dd/yyyy")
    If MsgBox("Do you really want to delete this amendment for FY " & FY & "?", vbYesNo, "Confirm Deletion") = vbYes Then
        If MsgBox("Are you absolutely sure you want to delete this amendment for FY " & FY & "?", vbYesNo, "Confirm Deletion") = vbYes Then
            TILLDataBase.Execute "INSERT INTO tblDELETEDContractsAmendments ( FY, ContractID, AmendmentNumber, RecordDeletedDate, RecordDeletedBy, " & _
                "BillingBookNumber, Pending, PendingAmount, PendingUnits, DDSArea, ProgramName, PurposeReason, DateSubmitted, DateApproved, StartDate, " & _
                "EndDate, NumUnits, NewRate, AmendedAmount, AnnualizedAmount, Comments ) " & _
                "SELECT FY, ContractID, AmendmentNumber, """ & DeletionDate & """ AS RecordDeletedDate, """ & Form_frmMainMenu.UserName & """ AS RecordDeletedBy, " & _
                "BillingBookNumber, Pending, PendingAmount, PendingUnits, DDSArea, ProgramName, PurposeReason, DateSubmitted, DateApproved, StartDate, " & _
                "EndDate, NumUnits, NewRate, AmendedAmount, AnnualizedAmount, Comments " & _
                "FROM tblContractsAmendments WHERE FY=" & Form_frmContractsAmendments.FY & " AND " & "AmendmentNumber=" & Form_frmContractsAmendments.AmendmentNumber, dbSeeChanges: Call BriefDelay
            TILLDataBase.Execute "DELETE * FROM tblContractsAmendments " & _
                "WHERE FY=" & Form_frmContractsAmendments.FY & " AND " & "AmendmentNumber=" & Form_frmContractsAmendments.AmendmentNumber, dbSeeChanges: Call BriefDelay
            Me.Requery
            DoCmd.GoToRecord , , acFirst
        End If
    End If
    Exit Sub
ShowMeError:
    MsgBox "Error # " & Str(Err.Number) & " was generated by " & Me.Name & Chr(13) & Err.Description, vbOKOnly, "Error", Err.HelpFile, Err.HelpContext
End Sub
