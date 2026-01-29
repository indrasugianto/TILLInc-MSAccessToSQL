' Module Name: Form_frmLocationsAddressMaintenance
' Module Type: Document Module
' Lines of Code: 17
' Extracted: 1/29/2026 4:12:23 PM

Option Compare Database
Option Explicit

Private Sub Form_Load()
    Me.Requery
End Sub

Private Sub Selected_Click()
    With Form_frmLocations
        .Address = CandidateAddress
        .City = CandidateCity
        .State = CandidateState
        .ZIP = CandidateZIP
        .AddressValidated = True: .NewRecordInProgress = False: .SetFocus
    End With
    DoCmd.Close
End Sub