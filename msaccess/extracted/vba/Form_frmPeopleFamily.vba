' Module Name: Form_frmPeopleFamily
' Module Type: Document Module
' Lines of Code: 96
' Extracted: 1/29/2026 4:12:24 PM

Option Compare Database
Option Explicit

Private Sub Form_Current()
    Dim InvalidFamilyCount As Integer
    
    ' Display family member name.
    If Left(IndexedName, 3) = "///" Then
        CompanyOrganization.Visible = True:  FamilyMemberName.Visible = False
    Else
        CompanyOrganization.Visible = False: FamilyMemberName.Visible = True
    End If
    ' No relationship set.  Set it as "UNKNOWN".
    If Len(Relationship) <= 0 Then
        Relationship = "* UNKNOWN *": Call UpdateChangeLog("Relationship-AutoSet", Relationship)
    End If
    ' If UNKNOWN relationship, red-flag this.
    If Relationship = "* UNKNOWN *" Then Relationship.BackColor = RGB(255, 0, 0) Else Relationship.BackColor = RGB(255, 255, 255)
    Surrogate.Visible = True
    ' If these three relationships are set, then the toggles, by definition, should also be set.
    If Relationship = "Guardian of" And Guardian = False Then
        Guardian = True: Call UpdateChangeLog("Guardian-AutoSet", "True")
    End If
    If Relationship = "Surrogate of" And Surrogate = False Then
        Surrogate = True: Call UpdateChangeLog("Surrogate-AutoSet", "True")
    End If
    If Relationship = "Rep Payee of" And RepPayee = False Then
        RepPayee = True: Call UpdateChangeLog("Rep Payee-AutoSet", "True")
    End If
    ' Here, we check the "sanity" of the family records.
    If Relationship = "Former guardian of" And Guardian = True Then
        Guardian = False: Call UpdateChangeLog("Guardian-AutoSet", "False")
    End If
    ' Count the number of family members that do not have at least one toggle selected.
    InvalidFamilyCount = DCount("IndexedName", "tblPeopleFamily", _
        "IndexedName = """ & Form_frmPeopleFamily.IndexedName & """ AND IsDeceased = False AND PrimaryContact = False AND Guardian = False AND Surrogate = False AND RepPayee = False AND Inactive = False")
    If InvalidFamilyCount > 0 Then
        MsgBox "One or more associated clients are not flagged with Guardian, Contact, Surrogate, Rep Payee, or Inactive.", vbOKOnly, "Warning!"
        InvalidFamilyCountMessageDisplay = True
    End If
End Sub

Private Sub FindClient_Click()
    DoCmd.OpenForm "frmPeopleSelectPerson", , , , , , "FamilySelectClient"
End Sub

Private Sub Form_BeforeInsert(Cancel As Integer)
    RecordAddedDate = Format(Now(), "mm/dd/yyyy")
    RecordAddedBy = Form_frmMainMenu.UserName
End Sub

Private Function UpdateRelationshipChangeLog(BoolField As Boolean, StringField As String) As Boolean
    Dim LogRecord As String
    
    UpdateRelationshipChangeLog = True
    LogRecord = StringField & " of " & ClientFirstName & " " & ClientMiddleInitial & " " & ClientLastName
    If BoolField Then Call UpdateChangeLog(LogRecord, "True") Else Call UpdateChangeLog(LogRecord, "False")
End Function

Private Sub Inactive_Click()
    Dim LogRecord As String
    
    LogRecord = "Inactive for " & ClientFirstName & " " & ClientMiddleInitial & " " & ClientLastName
    If Inactive Then
        Call UpdateChangeLog(LogRecord, "True"):  Call GreyAndNormal(Form_frmPeople.IsFamilyGuardianLabel)
    Else
        Call UpdateChangeLog(LogRecord, "False"): Call BlueAndBold(Form_frmPeople.IsFamilyGuardianLabel)
    End If
    Call CheckPersonCompletelyInactive
End Sub

Private Sub RepPayee_Click()
    If RepPayee Then
        ' Apply the family information to the Rep Payee information for the associated client.
        DoCmd.SetWarnings False: DoCmd.OpenQuery "qryPeopleAddFamilyRepPayee": DoCmd.SetWarnings True
    Else
        ' Undo the family information to the Rep Payee information for the associated client.
        DoCmd.SetWarnings False: DoCmd.OpenQuery "qryPeopleRemoveFamilyRepPayee": DoCmd.SetWarnings True
    End If
    ' Log the change.
    Call UpdateRelationshipChangeLog(RepPayee, "Rep Payee Flag")
End Sub

Private Sub Relationship_AfterUpdate()
'   If Not (IsNull(ClientIndexedName)) Then Me.Dirty = False
    If Relationship = "Guardian of" And Guardian = False Then
        Guardian = True: Call UpdateChangeLog("Guardian-AutoSet", "True")
    End If
    If Relationship = "Surrogate of" And Surrogate = False Then
        Surrogate = True: Call UpdateChangeLog("Surrogate-AutoSet", "True")
    End If
    If Relationship = "Rep Payee of" And RepPayee = False Then
        RepPayee = True: Call UpdateChangeLog("Rep Payee-AutoSet", "True")
    End If
    Call UpdateChangeLog("Relationship of " & ClientFirstName & " " & ClientMiddleInitial & " " & ClientLastName, Relationship)
End Sub