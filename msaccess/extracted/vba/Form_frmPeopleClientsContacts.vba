' Module Name: Form_frmPeopleClientsContacts
' Module Type: Document Module
' Lines of Code: 11
' Extracted: 2026-02-04 13:03:35

Option Compare Database
Option Explicit

Private Sub JumpToFamilyMember_Click()
    With Form_frmPeople
        .ClientJumpName = FamilyIndexedName
        .ClientJumpFromFamily = 1
        .SetFocus
        .Requery
    End With
End Sub
