' Module Name: Report_rptEXPIRATIONDATESclients
' Module Type: Document Module
' Lines of Code: 218
' Extracted: 2026-02-04 13:03:35

Option Compare Database
Option Explicit

'Private Sub Detail_Format(Cancel As Integer, FormatCount As Integer)
''On Error GoTo ShowMeError
'On Error GoTo 0
'    Dim ErrorField As Variant
'
'    ErrorField = "DateISP": NextISPTxt.Visible = False: DateISPFmt.Visible = False
'    Select Case DateISP
'        Case ExpMissingCalculated, ExpOptionalCalculated, ExpNACalculated: Call SetExpirationFieldProperties(NextISPTxt, , True, DateISP)
'        Case Else
'            DateISPFmt.Visible = Not IsEmpty(DateISP) And Not IsNull(DateISP) And (LengthN(8, DateISP) Or LengthN(10, DateISP))
'            If DateISPFmt.Visible Then
'                If (DateValue(DateISPFmt) - Int(Now)) < Trig_Indiv_ISP_Red Then
'                    DateISPFmt.ForeColor = RGB(255, 0, 0): DateISPFmt.FontWeight = 700: DateISPFmt.BorderColor = RGB(255, 0, 0): DateISPFmt.BorderStyle = 1
'                ElseIf (DateValue(DateISPFmt) - Int(Now)) <= Trig_Indiv_ISP_Green Then
'                    DateISPFmt.ForeColor = RGB(18, 94, 40): DateISPFmt.FontWeight = 700: DateISPFmt.BorderColor = RGB(18, 94, 40): DateISPFmt.BorderStyle = 1
'                Else
'                    DateISPFmt.ForeColor = RGB(0, 0, 0): DateISPFmt.FontWeight = 400: DateISPFmt.BorderStyle = 0
'                End If
'            End If
'    End Select
'
'    ErrorField = "PSDue": PSDueTxt.Visible = False: PSDueFmt.Visible = True: PSStrikeThru.Visible = False
'    Select Case DateISP
'        Case ExpMissingCalculated, ExpOptionalCalculated, ExpNACalculated: PSDue = Null
'        Case Else
'            PSDue = DateAdd("d", -182, [DateISP])
'            PSDueFmt = Format(PSDue, "mm/dd/yy")
'            If IsNull(PSDue) Then
'                PSDueTxt.Visible = True: PSDueFmt.Visible = False
'            ElseIf (PSDue - Int(Now)) <= Trig_Indiv_PSDue_Green Then
'                If Int(Now) > PSDue Then
'                    PSDueFmt.ForeColor = RGB(0, 0, 0): PSDueFmt.FontWeight = 400: PSDueFmt.BorderStyle = 0: PSStrikeThru.Visible = True
'                Else
'                    PSDueFmt.ForeColor = RGB(18, 94, 40): PSDueFmt.FontWeight = 700: PSDueFmt.BorderColor = RGB(18, 94, 40): PSDueFmt.BorderStyle = 1
'                End If
'            Else
'                PSDueFmt.ForeColor = RGB(0, 0, 0): PSDueFmt.FontWeight = 400: PSDueFmt.BorderStyle = 0
'            End If
'    End Select
'
'    ErrorField = "DateConsentFormsSigned": ConsentFormsTxt.Visible = False: DateConsentFormsSignedFmt.Visible = False
'    Select Case DateConsentFormsSigned
'        Case ExpMissingCalculated, ExpOptionalCalculated, ExpNACalculated: Call SetExpirationFieldProperties(ConsentFormsTxt, , True, DateConsentFormsSigned)
'        Case Else
'            DateConsentFormsSignedFmt.Visible = Not IsEmpty(DateConsentFormsSigned) And Not IsNull(DateConsentFormsSigned) And (LengthN(8, DateConsentFormsSigned) Or LengthN(10, DateConsentFormsSigned))
'            If DateConsentFormsSignedFmt.Visible Then
'                If (Int(Now) - Int(DateAdd("m", Trig_Indiv_CFS_Red, DateValue(DateConsentFormsSignedFmt)))) > 0 Then
'                    DateConsentFormsSignedFmt.ForeColor = RGB(255, 0, 0):  DateConsentFormsSignedFmt.FontWeight = 700: DateConsentFormsSignedFmt.BorderColor = RGB(255, 0, 0): DateConsentFormsSignedFmt.BorderStyle = 1
'                ElseIf (Int(Now) - Int(DateAdd("m", Trig_Indiv_CFS_Green, DateValue(DateConsentFormsSignedFmt)))) > 0 Then
'                    DateConsentFormsSignedFmt.ForeColor = RGB(18, 94, 40):  DateConsentFormsSignedFmt.FontWeight = 700: DateConsentFormsSignedFmt.BorderColor = RGB(18, 94, 40): DateConsentFormsSignedFmt.BorderStyle = 1
'                Else
'                    DateConsentFormsSignedFmt.ForeColor = RGB(0, 0, 0): DateConsentFormsSignedFmt.FontWeight = 400: DateConsentFormsSignedFmt.BorderStyle = 0
'                End If
'            End If
'    End Select
'
'    ErrorField = "DateBMMExpires": BMMExpiresTxt.Visible = False: DateBMMExpiresFmt.Visible = False
'    Select Case DateBMMExpires
'        Case ExpMissingCalculated, ExpOptionalCalculated, ExpNACalculated: Call SetExpirationFieldProperties(BMMExpiresTxt, , True, DateBMMExpires)
'        Case Else
'            DateBMMExpiresFmt.Visible = Not IsEmpty(DateBMMExpires) And Not IsNull(DateBMMExpires) And (LengthN(8, DateBMMExpires) Or LengthN(10, DateBMMExpires))
'            If DateBMMExpiresFmt.Visible Then
'                If (DateValue(DateBMMExpiresFmt) - Int(Now)) < Trig_Indiv_BMMX_Red Then
'                    DateBMMExpiresFmt.ForeColor = RGB(255, 0, 0): DateBMMExpiresFmt.FontWeight = 700: DateBMMExpiresFmt.BorderColor = RGB(255, 0, 0): DateBMMExpiresFmt.BorderStyle = 1
'                ElseIf (DateValue(DateBMMExpiresFmt) - Int(Now)) <= Trig_Indiv_BMMX_Green Then
'                    DateBMMExpiresFmt.ForeColor = RGB(18, 94, 40): DateBMMExpiresFmt.FontWeight = 700: DateBMMExpiresFmt.BorderColor = RGB(18, 94, 40): DateBMMExpiresFmt.BorderStyle = 1
'                Else
'                    DateBMMExpiresFmt.ForeColor = RGB(0, 0, 0): DateBMMExpiresFmt.FontWeight = 400: DateBMMExpiresFmt.BorderStyle = 0
'                End If
'            End If
'    End Select
'
'    ErrorField = "DateSPDAuthExpires": SPDAuthTxt.Visible = False: DateSPDAuthExpiresFmt.Visible = False
'    Select Case DateSPDAuthExpires
'        Case ExpMissingCalculated, ExpOptionalCalculated, ExpNACalculated: Call SetExpirationFieldProperties(SPDAuthTxt, , True, DateSPDAuthExpires)
'        Case Else
'            DateSPDAuthExpiresFmt.Visible = Not IsEmpty(DateSPDAuthExpires) And Not IsNull(DateSPDAuthExpires) And (LengthN(8, DateSPDAuthExpires) Or LengthN(10, DateSPDAuthExpires))
'            If DateSPDAuthExpiresFmt.Visible Then
'                If (DateValue(DateSPDAuthExpiresFmt) - Int(Now)) < Trig_Indiv_SPDX_Red Then
'                    DateSPDAuthExpiresFmt.ForeColor = RGB(255, 0, 0): DateSPDAuthExpiresFmt.FontWeight = 700
'                    DateSPDAuthExpiresFmt.BorderColor = RGB(255, 0, 0): DateSPDAuthExpiresFmt.BorderStyle = 1
'                ElseIf (DateValue(DateSPDAuthExpiresFmt) - Int(Now)) <= Trig_Indiv_SPDX_Green Then
'                    DateSPDAuthExpiresFmt.ForeColor = RGB(18, 94, 40): DateSPDAuthExpiresFmt.FontWeight = 700
'                    DateSPDAuthExpiresFmt.BorderColor = RGB(18, 94, 40): DateSPDAuthExpiresFmt.BorderStyle = 1
'                Else
'                    DateSPDAuthExpiresFmt.ForeColor = RGB(0, 0, 0): DateSPDAuthExpiresFmt.FontWeight = 400: DateSPDAuthExpiresFmt.BorderStyle = 0
'                End If
'            End If
'    End Select
'
'    ErrorField = "DateSignaturesDueBy": SPDAuthTxt.Visible = False: DateSignaturesDueByFmt.Visible = False
'    Select Case DateSignaturesDueBy
'        Case ExpMissingCalculated, ExpOptionalCalculated, ExpNACalculated: Call SetExpirationFieldProperties(DateSignaturesDueByTxt, , True, DateSignaturesDueBy)
'        Case Else
'' Start
'            DateSignaturesDueByFmt.Visible = Not IsEmpty(DateSignaturesDueBy) And Not IsNull(DateSignaturesDueBy) And (LengthN(8, DateSignaturesDueBy) Or LengthN(10, DateSignaturesDueBy))
'            If DateSignaturesDueByFmt.Visible Then
'                If (DateValue(DateSignaturesDueByFmt) - Int(Now)) < Trig_Indiv_SPDA_Red Then
'                    DateSignaturesDueByFmt.ForeColor = RGB(255, 0, 0): DateSignaturesDueByFmt.FontWeight = 700
'                    DateSignaturesDueByFmt.BorderColor = RGB(255, 0, 0): DateSignaturesDueByFmt.BorderStyle = 1
'                ElseIf (DateValue(DateSignaturesDueByFmt) - Int(Now)) <= Trig_Indiv_SPDA_Green Then
'                    DateSignaturesDueByFmt.ForeColor = RGB(18, 94, 40): DateSignaturesDueByFmt.FontWeight = 700
'                    DateSignaturesDueByFmt.BorderColor = RGB(18, 94, 40): DateSignaturesDueByFmt.BorderStyle = 1
'                Else
'                    DateSignaturesDueByFmt.ForeColor = RGB(0, 0, 0): DateSignaturesDueByFmt.FontWeight = 400: DateSignaturesDueByFmt.BorderStyle = 0
'                End If
'            End If
'    End Select
'
'    Exit Sub
'ShowMeError:
'    MsgBox "Error # " & Str(Err.Number) & " was generated by " & Me.Name & Chr(13) & Err.Description, vbOKOnly, "Error", Err.HelpFile, Err.HelpContext
'End Sub

Private Sub Detail_Format(Cancel As Integer, FormatCount As Integer)
On Error GoTo 0
    ' Record source must be vw_ExpirationsFormatted (or query from it) so _Display, _ShowDate, _Color columns exist.
    
    ' ========================================
    ' DateISP - binding to view columns
    ' ========================================
    If Nz(Me!DateISP_ShowDate, 0) = 1 Then
        NextISPTxt.Visible = False
        DateISPFmt.Visible = True
        Call ApplyColorFormatting(DateISPFmt, Nz(Me!DateISP_Color, "NORMAL"))
    Else
        NextISPTxt.Visible = True
        NextISPTxt.Caption = Nz(Me!DateISP_Display, "")
        NextISPTxt.ForeColor = IIf(Nz(Me!DateISP_Color, "") = "RED", RGB(255, 0, 0), RGB(0, 0, 0))
        NextISPTxt.FontWeight = IIf(Nz(Me!DateISP_Color, "") = "RED", 700, 400)
        NextISPTxt.BorderStyle = IIf(Nz(Me!DateISP_Color, "") = "RED", 1, 0)
        NextISPTxt.BorderColor = IIf(Nz(Me!DateISP_Color, "") = "RED", RGB(255, 0, 0), RGB(0, 0, 0))
        DateISPFmt.Visible = False
    End If
    
    ' ========================================
    ' PSDue - pre-calculated 182 days before DateISP
    ' ========================================
    If Nz(Me!PSDue_ShowText, 0) = 1 Then
        PSDueTxt.Visible = True
        PSDueTxt.Caption = Nz(Me!PSDue_Display, "")
        PSDueTxt.BorderStyle = 0
        PSDueTxt.BorderColor = RGB(0, 0, 0)
        PSDueFmt.Visible = False
        PSStrikeThru.Visible = False
    Else
        PSDueTxt.Visible = False
        PSDueFmt.Visible = True
        PSDueFmt.Value = Me!PSDue_Display
        Call ApplyColorFormatting(PSDueFmt, Nz(Me!PSDue_Color, "NORMAL"))
        PSStrikeThru.Visible = (Nz(Me!PSDue_Color, "") = "STRIKETHROUGH")
    End If
    
    ' ========================================
    ' Client-only fields (do NOT add Day/House fields like LastVehicleChecklistCompleted here)
    ' ========================================
    
    ' DateConsentFormsSigned
    Call FormatExpirationField(Nz(Me!DateConsentFormsSigned_ShowDate, 0), Nz(Me!DateConsentFormsSigned_Display, ""), _
                               Nz(Me!DateConsentFormsSigned_Color, "NORMAL"), ConsentFormsTxt, DateConsentFormsSignedFmt)
    
    ' DateBMMExpires
    Call FormatExpirationField(Nz(Me!DateBMMExpires_ShowDate, 0), Nz(Me!DateBMMExpires_Display, ""), _
                               Nz(Me!DateBMMExpires_Color, "NORMAL"), BMMExpiresTxt, DateBMMExpiresFmt)
    
    ' DateSPDAuthExpires
    Call FormatExpirationField(Nz(Me!DateSPDAuthExpires_ShowDate, 0), Nz(Me!DateSPDAuthExpires_Display, ""), _
                               Nz(Me!DateSPDAuthExpires_Color, "NORMAL"), SPDAuthTxt, DateSPDAuthExpiresFmt)
    
    ' DateSignaturesDueBy
    Call FormatExpirationField(Nz(Me!DateSignaturesDueBy_ShowDate, 0), Nz(Me!DateSignaturesDueBy_Display, ""), _
                               Nz(Me!DateSignaturesDueBy_Color, "NORMAL"), DateSignaturesDueByTxt, DateSignaturesDueByFmt)
    
End Sub

' ========================================
' Helper function to apply color formatting
' ========================================
Private Sub ApplyColorFormatting(ctl As Control, colorCode As String)
    Select Case colorCode
        Case "RED"
            ctl.ForeColor = RGB(255, 0, 0)
            ctl.FontWeight = 700
            ctl.BorderColor = RGB(255, 0, 0)
            ctl.BorderStyle = 1
        Case "GREEN"
            ctl.ForeColor = RGB(18, 94, 40)
            ctl.FontWeight = 700
            ctl.BorderColor = RGB(18, 94, 40)
            ctl.BorderStyle = 1
        Case "STRIKETHROUGH"
            ctl.ForeColor = RGB(0, 0, 0)
            ctl.FontWeight = 400
            ctl.BorderStyle = 0
            ctl.BorderColor = RGB(0, 0, 0)
        Case Else ' NORMAL
            ctl.ForeColor = RGB(0, 0, 0)
            ctl.FontWeight = 400
            ctl.BorderStyle = 0
            ctl.BorderColor = RGB(0, 0, 0)
    End Select
End Sub

' ========================================
' Helper function for standard field formatting
' displayValue: from view (e.g. "Missing", "Optional", "N/A", or "" for blank).
' ========================================
Private Sub FormatExpirationField(showDate As Integer, displayValue As String, _
                                  colorCode As String, txtLabel As Control, fmtDate As Control)
    If showDate = 1 Then
        txtLabel.Visible = False
        fmtDate.Visible = True
        Call ApplyColorFormatting(fmtDate, colorCode)
    Else
        txtLabel.Visible = True
        txtLabel.Caption = Nz(displayValue, "")
        txtLabel.ForeColor = IIf(colorCode = "RED", RGB(255, 0, 0), RGB(0, 0, 0))
        txtLabel.FontWeight = IIf(colorCode = "RED", 700, 400)
        txtLabel.BorderStyle = IIf(colorCode = "RED", 1, 0)
        txtLabel.BorderColor = IIf(colorCode = "RED", RGB(255, 0, 0), RGB(0, 0, 0))
        fmtDate.Visible = False
    End If
End Sub

