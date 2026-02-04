' Module Name: Form_frmMaintEditReferenceTables
' Module Type: Document Module
' Lines of Code: 6
' Extracted: 2026-02-04 13:03:35

Option Compare Database
Option Explicit

Private Function OpenReferenceTable(TableName As String) As Boolean
    OpenReferenceTable = True: DoCmd.OpenTable TableName, acViewNormal
End Function
