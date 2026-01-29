' Module Name: Report_rptDIRECTREPORTS
' Module Type: Document Module
' Lines of Code: 17
' Extracted: 1/29/2026 4:12:28 PM

Private Sub Detail_Format(Cancel As Integer, FormatCount As Integer)
    DirectReportFirstNameFixed = SpecialNames(StrConv(DirectReportFirstName, vbProperCase))
    DirectReportLastNameFixed = SpecialNames(StrConv(DirectReportLastName, vbProperCase))
End Sub

Private Sub GroupHeader0_Format(Cancel As Integer, FormatCount As Integer)
    CostCenter = DLookup("CostCenter", "tblLocations", "StaffPrimaryContactIndexedName=""" & SupervisorIndexedName & """")
    If IsNull(CostCenter) Or Len(CostCenter) = 0 Or CostCenter = "0010" Then
        SupervisorNameWithoutCC.Visible = True
        SupervisorNameWithCC.Visible = False
        If SigningAuthority Then SupervisorNameWithoutCC.ForeColor = RGB(0, 168, 0) Else SupervisorNameWithoutCC.ForeColor = RGB(0, 0, 0)
    Else
        SupervisorNameWithoutCC.Visible = False
        SupervisorNameWithCC.Visible = True
        If SigningAuthority Then SupervisorNameWithCC.ForeColor = RGB(0, 168, 0) Else SupervisorNameWithCC.ForeColor = RGB(0, 0, 0)
    End If
End Sub