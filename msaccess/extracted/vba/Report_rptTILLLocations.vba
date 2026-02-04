' Module Name: Report_rptTILLLocations
' Module Type: Document Module
' Lines of Code: 15
' Extracted: 2026-02-04 13:03:36

Option Compare Database

Private Sub Detail_Format(Cancel As Integer, FormatCount As Integer)
    If Department = "Residential Services" Then
        ResCapacity.Visible = True
        Cluster.Visible = True
        ABI.Visible = True
        ResTILLOwned.Visible = True
    Else
        ResCapacity.Visible = False
        Cluster.Visible = False
        ABI.Visible = False
        ResTILLOwned.Visible = False
    End If
End Sub
