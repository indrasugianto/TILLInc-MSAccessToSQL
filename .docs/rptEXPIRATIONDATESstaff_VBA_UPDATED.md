# rptEXPIRATIONDATESstaff – Updated VBA for vw_ExpirationsFormatted

**Last updated:** February 2, 2026

Use this code **after** the view has Staff columns and hidden controls are added (see **Prerequisites** below).

---

## Prerequisites

1. **Record Source**  
   Set **rptEXPIRATIONDATESstaff** Record Source to **vw_ExpirationsFormatted**.

2. **View has Staff columns**  
   **vw_ExpirationsFormatted** includes Staff columns (`BBP_Display`, `BBP_Color`, `BBP_ShowDate`, etc.).  
   Redeploy **msaccess/extracted/sql/vw_ExpirationsFormatted.sql** on Azure SQL, then refresh the linked table in Access (**External Data → Linked Table Manager**).

3. **Hidden controls**  
   Run **FixStaffReport()** or **FixAllReports()** in **ModReportFieldManager** to add the hidden text boxes to the Staff report.

4. **Linked table**  
   Refresh the **vw_ExpirationsFormatted** link in **External Data → Linked Table Manager** after changing the view.

---

## Updated VBA – Detail_Format and helpers

**Report:** rptEXPIRATIONDATESstaff  
**Event:** Detail_Format  
**Rule:** Do **not** set `.Value` on bound controls; only visibility, Caption, and formatting.

```vba
Option Compare Database
Option Explicit

Public DateThreeMonthsOut As Date, ThisYear As Date, NextYear As Date

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
            If (DateValue(EvalDueByFmt) - Int(Now)) < Nz(DLookup("Red", "catExpirationTriggers", "FieldName='3MoEval' AND Section IS NULL AND Program IS NULL"), -1) Then
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

' ----- Helper: do not set .Value on bound report controls -----
Private Sub FormatExpirationField(showDate As Integer, displayValue As String, _
                                  colorCode As String, txtLabel As Control, fmtDate As Control)
    If showDate = 1 Then
        txtLabel.Visible = False
        fmtDate.Visible = True
        Call ApplyColorFormatting(fmtDate, colorCode)
    Else
        txtLabel.Visible = True
        txtLabel.Caption = displayValue
        txtLabel.ForeColor = IIf(colorCode = "RED", RGB(255, 0, 0), RGB(0, 0, 0))
        txtLabel.FontWeight = IIf(colorCode = "RED", 700, 400)
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
    End Select
End Sub
```

---

## Control names (must match your report)

- **BBP:** BBPTxt, BBPFmt  
- **BackInjuryPrevention:** BIPTxt, BIPFmt  
- **CPR:** CPRTxt, CPRFmt  
- **DefensiveDriving:** DefensiveDrivingTxt, DefensiveDrivingFmt  
- **DriversLicense:** DriversLicenseTxt, DriversLicenseFmt  
- **FirstAid:** FirstAidTxt, FirstAidFmt  
- **PBS:** PBSTxt, PBSFmt  
- **SafetyCares:** SafetyCaresTxt, SafetyCaresFmt  
- **TB:** TBTxt, TBFmt  
- **WheelchairSafety:** WheelchairSafetyTxt, WheelchairSafetyFmt  
- **WorkplaceViolence:** WorkplaceViolenceTxt, WorkplaceViolenceFmt  
- **ProfessionalLicenses:** ProfessionalLicensesTxt, ProfessionalLicensesFmt  
- **MAPCert:** MAPCertTxt, MAPCertFmt (and MapCertLabel, MapCertLine)  
- **EvalDueBy:** EvalDueByFmt only (no Txt in original – code shows Fmt when `EvalDueBy_ShowDate = 1`)  
- **LastSupervision:** LastSupervisionFmt (kept hidden)

If your report uses different control names, adjust the calls accordingly.

---

## If you get “Can't find the field”

Redeploy **msaccess/extracted/sql/vw_ExpirationsFormatted.sql** (it now includes the Staff section), then in Access run **External Data → Linked Table Manager** and refresh the **vw_ExpirationsFormatted** link.
