' Module Name: Report_rptONCALL
' Module Type: Document Module
' Lines of Code: 7
' Extracted: 1/29/2026 4:12:27 PM

Option Compare Database
Option Explicit

Private Sub PageFooterSection_Format(Cancel As Integer, FormatCount As Integer)
    RespiteManagerOnly = DLookup("StaffPrimaryContactFirstName", "tblLocations", "CityTown=""Chelsea"" AND LocationName=""Nichols Street Respite""") & " " & _
        DLookup("StaffPrimaryContactLastName", "tblLocations", "CityTown=""Chelsea"" AND LocationName=""Nichols Street Respite""")
End Sub