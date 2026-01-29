' Module Name: Form_frmPleaseWaitCustom
' Module Type: Document Module
' Lines of Code: 11
' Extracted: 1/29/2026 4:12:28 PM

Option Compare Database
Option Explicit

Private Sub Form_Current()
    Me.Refresh
    Call BriefDelay(1)
End Sub

Private Sub Form_Load()
    BannerMessage = OpenArgs
End Sub