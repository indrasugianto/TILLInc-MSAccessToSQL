' Module Name: Form_frmPeopleClientsContacts
' Module Type: Document Module
' Lines of Code: 11
' Extracted: 1/29/2026 4:12:23 PM

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