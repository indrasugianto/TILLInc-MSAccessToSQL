' Module Name: Form_frmPleaseWaitCustom
' Module Type: Document Module
' Lines of Code: 11
' Extracted: 2026-02-04 13:03:36

Option Compare Database
Option Explicit

Private Sub Form_Current()
    Me.Refresh
    Call BriefDelay(1)
End Sub

Private Sub Form_Load()
    BannerMessage = OpenArgs
End Sub
