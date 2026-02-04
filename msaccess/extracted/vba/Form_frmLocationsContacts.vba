' Module Name: Form_frmLocationsContacts
' Module Type: Document Module
' Lines of Code: 12
' Extracted: 2026-02-04 13:03:36

Option Compare Database
Option Explicit

Private Sub DeleteThisRecord_Click()
    TILLDataBase.Execute "DELETE * FROM tblLocationsContacts WHERE tblLocationsContacts.IDX=" & IDX & ";", dbSeeChanges: Call BriefDelay
    Me.Requery: DoCmd.GoToRecord , , acFirst
End Sub

Private Function UpdateDepartment(Flag As Boolean) As Boolean
    UpdateDepartment = True
    If IsNull(Form_frmLocationsContacts.Department) Or Len(Form_frmLocationsContacts.Department) = 0 Then Form_frmLocationsContacts.Department = Form_frmLocations.Department
End Function
