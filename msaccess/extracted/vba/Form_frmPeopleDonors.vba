' Module Name: Form_frmPeopleDonors
' Module Type: Document Module
' Lines of Code: 134
' Extracted: 2026-02-04 13:03:35

Option Compare Database
Option Explicit

Dim RememberDonationFrom1FirstName As Variant, RememberDonationFrom1LastName As Variant, RememberDonationFrom2FirstName As Variant, RememberDonationFrom2LastName As Variant, RememberDonationFromCompany As Variant

Private Sub DonationFrom1Salutation_Change()
    If Right(DonationFrom1Salutation, 1) = "." Then DonationFrom1Salutation = Left(DonationFrom1Salutation, Len(DonationFrom1Salutation) - 1)
    Call UpdateChangeLog("DonorDonationFrom1Salutation", [DonationFrom1Salutation])
End Sub

Private Sub DonationFrom2Salutation_Change()
    If Right(DonationFrom2Salutation, 1) = "." Then DonationFrom2Salutation = Left(DonationFrom2Salutation, Len(DonationFrom2Salutation) - 1)
    Call UpdateChangeLog("DonorDonationFrom2Salutation", [DonationFrom2Salutation])
End Sub

Private Sub Form_Current()
    RememberDonationFrom1FirstName = DonationFrom1FirstName: RememberDonationFrom1LastName = DonationFrom1LastName
    RememberDonationFrom2FirstName = DonationFrom2FirstName: RememberDonationFrom2LastName = DonationFrom2LastName
    RememberDonationFromCompany = DonationFromCompany
    
    Me.Caption = "Donor"
    If Left(IndexedName, 3) = "///" Then
        ShowFullName.Visible = False: ShowCompanyOrganization.Visible = True: ShowCompanyOrganization = Form_frmPeople.CompanyOrganization
        Me.Caption = "Donor: " & ShowCompanyOrganization
        If (Len(DonationFrom) <= 0) Or IsNull(DonationFrom) Then DonationFrom = ShowCompanyOrganization
    Else
        ShowFullName.Visible = True: ShowCompanyOrganization.Visible = False: ShowFullName = Form_frmPeople.DisplayName
        Me.Caption = "Donor: " & ShowFullName
        If (Len(DonationFrom) <= 0) Or IsNull(DonationFrom) Then DonationFrom = ShowFullName
    End If
    If Inactive Then
        AddNewDonation.Visible = False: Me.Caption = Me.Caption & " (INACTIVE)"
    Else
        AddNewDonation.Visible = True
    End If
    Call BriefDelay
    Me.Repaint
End Sub

Private Sub AddNewDonation_Click()
    DoCmd.OpenForm "frmPeopleDonorsNewDonation"
    Call LoopUntilClosed("frmPeopleDonorsNewDonation", acForm)
    Call BriefDelay
    Me.Requery
End Sub

Private Sub DNRSUMMARY_Click()
    Call ExecReport("rptDNRSUMMARY")
End Sub

Private Sub DonationFrom1FirstName_AfterUpdate()
    If IsNull(RememberDonationFrom1FirstName) Or Len(RememberDonationFrom1FirstName) <= 0 Then
        DonationFrom1FirstName = CorrectProperNames(StrConv(DonationFrom1FirstName, vbProperCase))
        RememberDonationFrom1FirstName = DonationFrom1FirstName
    End If
    DonationFrom1FirstName = SpecialNames(DonationFrom1FirstName)
    Call UpdateChangeLog("DonorDonationFrom1FirstName", DonationFrom1FirstName)
End Sub

Private Sub DonationFrom1LastName_AfterUpdate()
    If IsNull(RememberDonationFrom1LastName) Or Len(RememberDonationFrom1LastName) <= 0 Then
        DonationFrom1LastName = CorrectProperNames(StrConv(DonationFrom1LastName, vbProperCase))
        RememberDonationFrom1LastName = DonationFrom1LastName
    End If
    DonationFrom1LastName = SpecialNames(DonationFrom1LastName)
    Call UpdateChangeLog("DonorDonationFrom1LastName", DonationFrom1LastName)
End Sub

Private Sub DonationFrom2FirstName_AfterUpdate()
    If IsNull(RememberDonationFrom2FirstName) Or Len(RememberDonationFrom2FirstName) <= 0 Then
        DonationFrom2FirstName = CorrectProperNames(StrConv(DonationFrom2FirstName, vbProperCase))
        RememberDonationFrom2FirstName = DonationFrom2FirstName
    End If
    DonationFrom2FirstName = SpecialNames(DonationFrom2FirstName)
    Call UpdateChangeLog("DonorDonationFrom2FirstName", DonationFrom2FirstName)
End Sub

Private Sub DonationFrom2LastName_AfterUpdate()
    If IsNull(RememberDonationFrom2LastName) Or Len(RememberDonationFrom2LastName) <= 0 Then
        DonationFrom2LastName = CorrectProperNames(StrConv(DonationFrom2LastName, vbProperCase))
        RememberDonationFrom2LastName = DonationFrom2LastName
    End If
    DonationFrom2LastName = SpecialNames(DonationFrom2LastName)
    Call UpdateChangeLog("DonorDonationFrom2LastName", DonationFrom2LastName)
End Sub

Private Sub DonationFromCompany_AfterUpdate()
    If IsNull(RememberDonationFromCompany) Or Len(RememberDonationFromCompany) <= 0 Then
        DonationFromCompany = CorrectProperNames(StrConv(DonationFromCompany, vbProperCase))
        RememberDonationFromCompany = DonationFromCompany
    End If
    DonationFromCompany = SpecialNames(DonationFromCompany)
    Call UpdateChangeLog("DonorDonationFromCompany", DonationFromCompany)
End Sub

Private Sub FilterCriteria_AfterUpdate()
    Select Case FilterCriteria
        Case "<Show All>", Null
            Me.FilterOn = True
        Case Else
            Me.Filter = "IndexedName=""" & IndexedName & """" & " AND " & "Year(DateOfDonation) = " & FilterCriteria
            Me.FilterOn = True
    End Select
End Sub

Private Sub Inactive_Click()
    If Inactive Then
        AddNewDonation.Visible = False
        Me.Caption = Me.Caption & " (INACTIVE)"
        Call GreyAndNormal(Form_frmPeople.IsDonorLabel)
        Me.Repaint
    Else
        AddNewDonation.Visible = True
        Me.Caption = Left(Me.Caption, Len(Me.Caption) - 11)
        Call BlueAndBold(Form_frmPeople.IsDonorLabel)
        Me.Repaint
    End If
    Me.Repaint
    Call CheckPersonCompletelyInactive
End Sub

Private Sub DeleteThisRecord_Click()
On Error GoTo ShowMeError
    If MsgBox("Do you really want to delete this record?", vbYesNo, "Confirm Deletion") = vbNo Then Exit Sub
    TILLDataBase.Execute "DELETE * FROM tblPeopleDonors WHERE Index = " & Index, dbSeeChanges: Call BriefDelay
    Me.Requery
    Exit Sub
ShowMeError:
    MsgBox "Error # " & Str(Err.Number) & " was generated by " & Me.Name & Chr(13) & Err.Description, vbOKOnly, "Error", Err.HelpFile, Err.HelpContext
End Sub

Private Sub IsGrant_Click()
    If IsGrant Then Call UpdateChangeLog("DonorIsGrant", "TRUE") Else Call UpdateChangeLog("DonorIsGrant", "FALSE")
End Sub
