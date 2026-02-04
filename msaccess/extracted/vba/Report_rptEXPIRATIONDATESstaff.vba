' Module Name: Report_rptEXPIRATIONDATESstaff
' Module Type: Document Module
' Lines of Code: 105
' Extracted: 2026-02-04 13:03:35

Option Compare Database
Option Explicit

Public DateThreeMonthsOut As Date, ThisYear As Date, NextYear As Date

' ----- Helper: do not set .Value on bound report controls -----
' displayValue: from view (e.g. "Missing", "Optional", "N/A", "Pending", or "" for blank).
' When blank (empty string), show nothing to match original behavior for null dates.
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
        Case Else
            ctl.ForeColor = RGB(0, 0, 0)
            ctl.FontWeight = 400
            ctl.BorderStyle = 0
            ctl.BorderColor = RGB(0, 0, 0)
    End Select
End Sub

Private Sub Detail_Format(Cancel As Integer, FormatCount As Integer)
On Error GoTo 0
    Dim FindDash As Long

    ' ----- FullName (FirstName + LastName with dash handling) -----
    FullName = RTrim(StrConv(FirstName, 3)) & " " & RTrim(StrConv(LastName, 3))
    FindDash = InStr(1, FullName, "-", vbTextCompare)
    If FindDash > 0 Then FullName = Left(FullName, FindDash) & StrConv(Mid(FullName, FindDash + 1, 1), vbUpperCase) & Mid(FullName, FindDash + 2, 256)

    ' ----- Standard staff fields (bound controls – do NOT set .Value) -----
    Call FormatExpirationField(BBP_ShowDate, BBP_Display, BBP_Color, BBPTxt, BBPFmt)
    Call FormatExpirationField(BackInjuryPrevention_ShowDate, BackInjuryPrevention_Display, BackInjuryPrevention_Color, BIPTxt, BIPFmt)
    Call FormatExpirationField(CPR_ShowDate, CPR_Display, CPR_Color, CPRTxt, CPRFmt)
    Call FormatExpirationField(DefensiveDriving_ShowDate, DefensiveDriving_Display, DefensiveDriving_Color, DefensiveDrivingTxt, DefensiveDrivingFmt)
    Call FormatExpirationField(DriversLicense_ShowDate, DriversLicense_Display, DriversLicense_Color, DriversLicenseTxt, DriversLicenseFmt)
    Call FormatExpirationField(FirstAid_ShowDate, FirstAid_Display, FirstAid_Color, FirstAidTxt, FirstAidFmt)
    Call FormatExpirationField(PBS_ShowDate, PBS_Display, PBS_Color, PBSTxt, PBSFmt)
    Call FormatExpirationField(SafetyCares_ShowDate, SafetyCares_Display, SafetyCares_Color, SafetyCaresTxt, SafetyCaresFmt)
    Call FormatExpirationField(WheelchairSafety_ShowDate, WheelchairSafety_Display, WheelchairSafety_Color, WheelchairSafetyTxt, WheelchairSafetyFmt)
    Call FormatExpirationField(WorkplaceViolence_ShowDate, WorkplaceViolence_Display, WorkplaceViolence_Color, WorkplaceViolenceTxt, WorkplaceViolenceFmt)
    Call FormatExpirationField(ProfessionalLicenses_ShowDate, ProfessionalLicenses_Display, ProfessionalLicenses_Color, ProfessionalLicensesTxt, ProfessionalLicensesFmt)

    ' ----- TB (with Pending + location-specific red) -----
    Call FormatExpirationField(TB_ShowDate, TB_Display, TB_Color, TBTxt, TBFmt)
    If TB_ShowDate = 0 And TB_Display = "Pending" Then
        If Location = "Hollis" Or Nz(DLookup("ABI", "tblLocations", "GPName='" & Replace(Location, "'", "''") & "'"), False) = True Or Nz(DLookup("Department", "tblLocations", "GPName='" & Replace(Location, "'", "''") & "'"), "") = "Day Services" Then
            TBTxt.Visible = True
            TBTxt.Caption = "Pending"
            TBTxt.ForeColor = RGB(255, 0, 0)
            TBTxt.FontWeight = 700
            TBTxt.BorderStyle = 1
            TBTxt.BorderColor = RGB(255, 0, 0)
            TBFmt.Visible = False
        End If
    End If

    ' ----- MAPCert (standard binding; N/A by location / 120-day purple stay in VBA if needed) -----
    MapCertLabel.Visible = True
    MapCertLine.Visible = True
    Call FormatExpirationField(MAPCert_ShowDate, MAPCert_Display, MAPCert_Color, MAPCertTxt, MAPCertFmt)
    ' Optional: N/A for certain DED locations and 120-day-from-AdjustedStartDate purple – keep existing logic here if you did not move it into the view.

    ' ----- EvalDueBy (no Txt control in original – show Fmt only when ShowDate=1, then 3-month eval blue) -----
    EvalDueByFmt.Visible = (EvalDueBy_ShowDate = 1)
    If EvalDueByFmt.Visible Then
        Call ApplyColorFormatting(EvalDueByFmt, EvalDueBy_Color)
        If Nz(ThreeMonthEvaluation, False) Then
            If (DateValue(EvalDueByFmt) - Int(Now)) < Nz(DLookup("Red", "catExpirationTriggers", "FieldName='3MoEval' AND Section='Staff' AND Program IS NULL"), -1) Then
                EvalDueByFmt.ForeColor = RGB(255, 0, 0)
                EvalDueByFmt.FontWeight = 700
                EvalDueByFmt.BorderColor = RGB(0, 0, 255)
                EvalDueByFmt.BorderStyle = 1
            Else
                EvalDueByFmt.ForeColor = RGB(0, 0, 255)
                EvalDueByFmt.FontWeight = 700
                EvalDueByFmt.BorderColor = RGB(0, 0, 255)
                EvalDueByFmt.BorderStyle = 1
            End If
        End If
    End If

    ' ----- LastSupervision (currently hidden in original; optional) -----
    LastSupervisionFmt.Visible = False
End Sub

