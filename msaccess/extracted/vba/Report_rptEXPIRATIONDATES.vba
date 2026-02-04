' Module Name: Report_rptEXPIRATIONDATES
' Module Type: Document Module
' Lines of Code: 19
' Extracted: 2026-02-04 13:03:35

Option Compare Database
Option Explicit

Private Sub Detail_Format(Cancel As Integer, FormatCount As Integer)
On Error GoTo 0
    If Department = "Day Services" Or Department = "Vocational Services" Then
        Me.rptEXPIRATIONDATESclients.Visible = False: Me.rptEXPIRATIONDATESday.Visible = True: Me.rptEXPIRATIONDATEShouse.Visible = False
    ElseIf Left(GPName, 4) = "DED-" Then
        Me.rptEXPIRATIONDATESclients.Visible = False: Me.rptEXPIRATIONDATESday.Visible = False: Me.rptEXPIRATIONDATEShouse.Visible = False
    Else
        Me.rptEXPIRATIONDATESclients.Visible = True: Me.rptEXPIRATIONDATESday.Visible = False: Me.rptEXPIRATIONDATEShouse.Visible = True
    End If
End Sub

Private Sub PageHeaderSection_Format(Cancel As Integer, FormatCount As Integer)
On Error GoTo 0
    If Department = "Residential Services" And Cluster <= "90" Then ClusterFormatted.Visible = True Else ClusterFormatted.Visible = False
End Sub

