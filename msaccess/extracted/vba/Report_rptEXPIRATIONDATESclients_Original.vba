' Module Name: Report_rptEXPIRATIONDATESclients_Original
' Module Type: Document Module
' Lines of Code: 116
' Extracted: 2026-02-04 13:03:36

Option Compare Database
Option Explicit

Private Sub Detail_Format(Cancel As Integer, FormatCount As Integer)
'On Error GoTo ShowMeError
On Error GoTo 0
    Dim ErrorField As Variant
    
    ErrorField = "DateISP": NextISPTxt.Visible = False: DateISPFmt.Visible = False
    Select Case DateISP
        Case ExpMissingCalculated, ExpOptionalCalculated, ExpNACalculated: Call SetExpirationFieldProperties(NextISPTxt, , True, DateISP)
        Case Else
            DateISPFmt.Visible = Not IsEmpty(DateISP) And Not IsNull(DateISP) And (LengthN(8, DateISP) Or LengthN(10, DateISP))
            If DateISPFmt.Visible Then
                If (DateValue(DateISPFmt) - Int(Now)) < Trig_Indiv_ISP_Red Then
                    DateISPFmt.ForeColor = RGB(255, 0, 0): DateISPFmt.FontWeight = 700: DateISPFmt.BorderColor = RGB(255, 0, 0): DateISPFmt.BorderStyle = 1
                ElseIf (DateValue(DateISPFmt) - Int(Now)) <= Trig_Indiv_ISP_Green Then
                    DateISPFmt.ForeColor = RGB(18, 94, 40): DateISPFmt.FontWeight = 700: DateISPFmt.BorderColor = RGB(18, 94, 40): DateISPFmt.BorderStyle = 1
                Else
                    DateISPFmt.ForeColor = RGB(0, 0, 0): DateISPFmt.FontWeight = 400: DateISPFmt.BorderStyle = 0
                End If
            End If
    End Select
    
    ErrorField = "PSDue": PSDueTxt.Visible = False: PSDueFmt.Visible = True: PSStrikeThru.Visible = False
    Select Case DateISP
        Case ExpMissingCalculated, ExpOptionalCalculated, ExpNACalculated: PSDue = Null
        Case Else
            PSDue = DateAdd("d", -182, [DateISP])
            PSDueFmt = Format(PSDue, "mm/dd/yy")
            If IsNull(PSDue) Then
                PSDueTxt.Visible = True: PSDueFmt.Visible = False
            ElseIf (PSDue - Int(Now)) <= Trig_Indiv_PSDue_Green Then
                If Int(Now) > PSDue Then
                    PSDueFmt.ForeColor = RGB(0, 0, 0): PSDueFmt.FontWeight = 400: PSDueFmt.BorderStyle = 0: PSStrikeThru.Visible = True
                Else
                    PSDueFmt.ForeColor = RGB(18, 94, 40): PSDueFmt.FontWeight = 700: PSDueFmt.BorderColor = RGB(18, 94, 40): PSDueFmt.BorderStyle = 1
                End If
            Else
                PSDueFmt.ForeColor = RGB(0, 0, 0): PSDueFmt.FontWeight = 400: PSDueFmt.BorderStyle = 0
            End If
    End Select
    
    ErrorField = "DateConsentFormsSigned": ConsentFormsTxt.Visible = False: DateConsentFormsSignedFmt.Visible = False
    Select Case DateConsentFormsSigned
        Case ExpMissingCalculated, ExpOptionalCalculated, ExpNACalculated: Call SetExpirationFieldProperties(ConsentFormsTxt, , True, DateConsentFormsSigned)
        Case Else
            DateConsentFormsSignedFmt.Visible = Not IsEmpty(DateConsentFormsSigned) And Not IsNull(DateConsentFormsSigned) And (LengthN(8, DateConsentFormsSigned) Or LengthN(10, DateConsentFormsSigned))
            If DateConsentFormsSignedFmt.Visible Then
                If (Int(Now) - Int(DateAdd("m", Trig_Indiv_CFS_Red, DateValue(DateConsentFormsSignedFmt)))) > 0 Then
                    DateConsentFormsSignedFmt.ForeColor = RGB(255, 0, 0):  DateConsentFormsSignedFmt.FontWeight = 700: DateConsentFormsSignedFmt.BorderColor = RGB(255, 0, 0): DateConsentFormsSignedFmt.BorderStyle = 1
                ElseIf (Int(Now) - Int(DateAdd("m", Trig_Indiv_CFS_Green, DateValue(DateConsentFormsSignedFmt)))) > 0 Then
                    DateConsentFormsSignedFmt.ForeColor = RGB(18, 94, 40):  DateConsentFormsSignedFmt.FontWeight = 700: DateConsentFormsSignedFmt.BorderColor = RGB(18, 94, 40): DateConsentFormsSignedFmt.BorderStyle = 1
                Else
                    DateConsentFormsSignedFmt.ForeColor = RGB(0, 0, 0): DateConsentFormsSignedFmt.FontWeight = 400: DateConsentFormsSignedFmt.BorderStyle = 0
                End If
            End If
    End Select
    
    ErrorField = "DateBMMExpires": BMMExpiresTxt.Visible = False: DateBMMExpiresFmt.Visible = False
    Select Case DateBMMExpires
        Case ExpMissingCalculated, ExpOptionalCalculated, ExpNACalculated: Call SetExpirationFieldProperties(BMMExpiresTxt, , True, DateBMMExpires)
        Case Else
            DateBMMExpiresFmt.Visible = Not IsEmpty(DateBMMExpires) And Not IsNull(DateBMMExpires) And (LengthN(8, DateBMMExpires) Or LengthN(10, DateBMMExpires))
            If DateBMMExpiresFmt.Visible Then
                If (DateValue(DateBMMExpiresFmt) - Int(Now)) < Trig_Indiv_BMMX_Red Then
                    DateBMMExpiresFmt.ForeColor = RGB(255, 0, 0): DateBMMExpiresFmt.FontWeight = 700: DateBMMExpiresFmt.BorderColor = RGB(255, 0, 0): DateBMMExpiresFmt.BorderStyle = 1
                ElseIf (DateValue(DateBMMExpiresFmt) - Int(Now)) <= Trig_Indiv_BMMX_Green Then
                    DateBMMExpiresFmt.ForeColor = RGB(18, 94, 40): DateBMMExpiresFmt.FontWeight = 700: DateBMMExpiresFmt.BorderColor = RGB(18, 94, 40): DateBMMExpiresFmt.BorderStyle = 1
                Else
                    DateBMMExpiresFmt.ForeColor = RGB(0, 0, 0): DateBMMExpiresFmt.FontWeight = 400: DateBMMExpiresFmt.BorderStyle = 0
                End If
            End If
    End Select
   
    ErrorField = "DateSPDAuthExpires": SPDAuthTxt.Visible = False: DateSPDAuthExpiresFmt.Visible = False
    Select Case DateSPDAuthExpires
        Case ExpMissingCalculated, ExpOptionalCalculated, ExpNACalculated: Call SetExpirationFieldProperties(SPDAuthTxt, , True, DateSPDAuthExpires)
        Case Else
            DateSPDAuthExpiresFmt.Visible = Not IsEmpty(DateSPDAuthExpires) And Not IsNull(DateSPDAuthExpires) And (LengthN(8, DateSPDAuthExpires) Or LengthN(10, DateSPDAuthExpires))
            If DateSPDAuthExpiresFmt.Visible Then
                If (DateValue(DateSPDAuthExpiresFmt) - Int(Now)) < Trig_Indiv_SPDX_Red Then
                    DateSPDAuthExpiresFmt.ForeColor = RGB(255, 0, 0): DateSPDAuthExpiresFmt.FontWeight = 700
                    DateSPDAuthExpiresFmt.BorderColor = RGB(255, 0, 0): DateSPDAuthExpiresFmt.BorderStyle = 1
                ElseIf (DateValue(DateSPDAuthExpiresFmt) - Int(Now)) <= Trig_Indiv_SPDX_Green Then
                    DateSPDAuthExpiresFmt.ForeColor = RGB(18, 94, 40): DateSPDAuthExpiresFmt.FontWeight = 700
                    DateSPDAuthExpiresFmt.BorderColor = RGB(18, 94, 40): DateSPDAuthExpiresFmt.BorderStyle = 1
                Else
                    DateSPDAuthExpiresFmt.ForeColor = RGB(0, 0, 0): DateSPDAuthExpiresFmt.FontWeight = 400: DateSPDAuthExpiresFmt.BorderStyle = 0
                End If
            End If
    End Select
    
    ErrorField = "DateSignaturesDueBy": SPDAuthTxt.Visible = False: DateSignaturesDueByFmt.Visible = False
    Select Case DateSignaturesDueBy
        Case ExpMissingCalculated, ExpOptionalCalculated, ExpNACalculated: Call SetExpirationFieldProperties(DateSignaturesDueByTxt, , True, DateSignaturesDueBy)
        Case Else
' Start
            DateSignaturesDueByFmt.Visible = Not IsEmpty(DateSignaturesDueBy) And Not IsNull(DateSignaturesDueBy) And (LengthN(8, DateSignaturesDueBy) Or LengthN(10, DateSignaturesDueBy))
            If DateSignaturesDueByFmt.Visible Then
                If (DateValue(DateSignaturesDueByFmt) - Int(Now)) < Trig_Indiv_SPDA_Red Then
                    DateSignaturesDueByFmt.ForeColor = RGB(255, 0, 0): DateSignaturesDueByFmt.FontWeight = 700
                    DateSignaturesDueByFmt.BorderColor = RGB(255, 0, 0): DateSignaturesDueByFmt.BorderStyle = 1
                ElseIf (DateValue(DateSignaturesDueByFmt) - Int(Now)) <= Trig_Indiv_SPDA_Green Then
                    DateSignaturesDueByFmt.ForeColor = RGB(18, 94, 40): DateSignaturesDueByFmt.FontWeight = 700
                    DateSignaturesDueByFmt.BorderColor = RGB(18, 94, 40): DateSignaturesDueByFmt.BorderStyle = 1
                Else
                    DateSignaturesDueByFmt.ForeColor = RGB(0, 0, 0): DateSignaturesDueByFmt.FontWeight = 400: DateSignaturesDueByFmt.BorderStyle = 0
                End If
            End If
    End Select

    Exit Sub
ShowMeError:
    MsgBox "Error # " & Str(Err.Number) & " was generated by " & Me.Name & Chr(13) & Err.Description, vbOKOnly, "Error", Err.HelpFile, Err.HelpContext
End Sub
