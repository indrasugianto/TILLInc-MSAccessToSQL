' Module Name: Form_frmMaintEditReferenceTables
' Module Type: Document Module
' Lines of Code: 6
' Extracted: 1/29/2026 4:12:23 PM

Option Compare Database
Option Explicit

Private Function OpenReferenceTable(TableName As String) As Boolean
    OpenReferenceTable = True: DoCmd.OpenTable TableName, acViewNormal
End Function