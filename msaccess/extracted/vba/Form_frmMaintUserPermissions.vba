' Module Name: Form_frmMaintUserPermissions
' Module Type: Document Module
' Lines of Code: 14
' Extracted: 2026-02-04 13:03:36

Option Compare Database
Option Explicit

Private Sub Form_Open(Cancel As Integer)
    Me.RecordSource = "SELECT catUserPermissions.* FROM catUserPermissions ORDER BY catUserPermissions.Action": Me.Requery
End Sub

Private Sub Label23_Click()
    Me.RecordSource = "SELECT catUserPermissions.* FROM catUserPermissions ORDER BY catUserPermissions.Action": Me.Requery
End Sub

Private Sub Label24_Click()
    Me.RecordSource = "SELECT catUserPermissions.* FROM catUserPermissions ORDER BY catUserPermissions.User": Me.Requery
End Sub
