' Module Name: Form_frmPeopleClientsVendors
' Module Type: Document Module
' Lines of Code: 75
' Extracted: 2026-02-04 13:03:35

Option Compare Database
Option Explicit

Private Sub InitializeJumpButtons()
    With Form_frmPeopleClientsDemographics
        JumpClient1.Visible = True
        JumpPeople.Visible = True
        JumpAutism.Visible = .JumpAutism.Visible
        JumpCLO.Visible = .JumpCLO.Visible
'       JumpCC.Visible = .JumpCC.Visible
'       JumpCC = False
        JumpDay.Visible = .JumpDay.Visible
        JumpISS.Visible = .JumpISS.Visible
'       JumpPCA.Visible = .JumpPCA.Visible
'       JumpPCA.Visible = False
        JumpResidential.Visible = .JumpResidential.Visible
        JumpSharedLiving.Visible = .JumpSharedLiving.Visible
        JumpNHDay.Visible = .JumpNHDay.Visible
        JumpNHRes.Visible = .JumpNHRes.Visible
        JumpSpringboard.Visible = .JumpSpringboard.Visible
        JumpTrans.Visible = .JumpTrans.Visible
'       JumpTRASE.Visible = .JumpTRASE.Visible
'       JumpTRASE = False
        JumpVoc.Visible = .JumpVoc.Visible
    End With
End Sub

Private Sub OpenServiceForm(FieldLabel As Label, FormName As String, JumpButton As CommandButton)
    Call BlueAndBold(FieldLabel)
    DoCmd.OpenForm FormName, , , "IndexedName=""" & IndexedName & """"
    JumpButton.Visible = True
End Sub

Private Sub Form_Current()
    Call ProgressMessages("Append", "   Open providers form.")

    Me.Caption = "Client: Providers"
    Call InitializeJumpButtons
    
    If Left(ResidentialVendor, 4) = "TILL" Then
        ResVendorLocation.Visible = True
    Else
        ResVendorLocation.Visible = False
    End If
    
    If Left(DayVendor, 4) = "TILL" Then
        DayVendorLocation.Visible = True
    Else
        DayVendorLocation.Visible = False
    End If
End Sub

Private Sub GetFamily_Click()
    DoCmd.OpenForm "frmPeopleSelectPerson", , , , , , "ClientSelectFamily"
End Sub

Private Sub ServiceClick(PeopleToggle As Boolean, ServiceToggle As Boolean, AddQuery As String, ServiceForm As String, Dept As String)
    PeopleToggle = ServiceToggle
    If ServiceToggle Then
        DoCmd.OpenQuery AddQuery
        DoCmd.OpenForm ServiceForm, , , "IndexedName=""" & IndexedName & """"
        Form_frmPeople.DeptCriteria = Dept
    Else
        DoCmd.Close acForm, ServiceForm
        Form_frmPeople.DeptCriteria = ""
    End If
End Sub

Private Sub LivingIndependently_Click()
    If LivingIndependently Then Call UpdateChangeLog("LivingIndependently", "True") Else Call UpdateChangeLog("LivingIndependently", "False")
End Sub

Private Sub LivingWithParentOrGuardian_Click()
    If LivingWithParentOrGuardian Then Call UpdateChangeLog("LivingWithParentOrGuardian", "True") Else Call UpdateChangeLog("LivingWithParentOrGuardian", "False")
End Sub
