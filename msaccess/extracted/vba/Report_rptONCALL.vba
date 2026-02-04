' Module Name: Report_rptONCALL
' Module Type: Document Module
' Lines of Code: 7
' Extracted: 2026-02-04 13:03:36

Option Compare Database
Option Explicit

Private Sub PageFooterSection_Format(Cancel As Integer, FormatCount As Integer)
    RespiteManagerOnly = DLookup("StaffPrimaryContactFirstName", "tblLocations", "CityTown=""Chelsea"" AND LocationName=""Nichols Street Respite""") & " " & _
        DLookup("StaffPrimaryContactLastName", "tblLocations", "CityTown=""Chelsea"" AND LocationName=""Nichols Street Respite""")
End Sub
