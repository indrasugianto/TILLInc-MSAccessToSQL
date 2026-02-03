# vw_ExpirationsFormatted - Field Reference

Quick reference for developers working with the pre-calculated Expirations view.  
**Last updated:** February 2, 2026

---

## Column Naming Convention

For each date field, the view provides **3 calculated columns:**

| Suffix | Type | Values | Purpose |
|--------|------|--------|---------|
| `_Display` | VARCHAR | Date string, "Missing", "Optional", "N/A", "Pending" | What to show on the report |
| `_Color` | VARCHAR | "RED", "GREEN", "NORMAL", "STRIKETHROUGH" | Color formatting to apply |
| `_ShowDate` | BIT | 0 or 1 | 1 = show date field, 0 = show text label |

**Program-Specific Fields:**  
Some fields have separate calculations for Day and Residential programs:
- `_Day` suffix = Day program calculations
- `_Res` suffix = Residential program calculations

---

## CLIENT FIELDS (RecordType = 'Client')

### DateISP (Individual Service Plan Expiration)

| Column | Type | Description | Example Values |
|--------|------|-------------|----------------|
| `DateISP` | DATE | Original date from tblExpirations | 2026-03-15, 1900-01-01 |
| `DateISP_Display` | VARCHAR | Formatted for display | "03/15/2026", "Missing" |
| `DateISP_Color` | VARCHAR | Color code | "RED", "GREEN", "NORMAL" |
| `DateISP_ShowDate` | BIT | Show date (1) or text (0) | 1, 0 |

**VBA Usage:** (do not set .Value on bound report controls)
```vba
If DateISP_ShowDate = 1 Then
    DateISPFmt.Visible = True
    Call ApplyColorFormatting(DateISPFmt, DateISP_Color)
Else
    NextISPTxt.Visible = True
    NextISPTxt.Caption = DateISP_Display  ' "Missing", "Optional", "N/A"
End If
```

### PSDue (Person Served Due - Calculated as DateISP - 182 days)

| Column | Type | Description | Example Values |
|--------|------|-------------|----------------|
| `PSDue_Calculated` | DATE | DateISP - 182 days | 2025-09-15, NULL |
| `PSDue_Display` | VARCHAR | Formatted for display | "09/15/2025", "N/A" |
| `PSDue_Color` | VARCHAR | Color code (includes STRIKETHROUGH) | "GREEN", "STRIKETHROUGH", "NORMAL" |
| `PSDue_ShowText` | BIT | Show text (1) or date (0) | 0, 1 |

**VBA Usage:** (do not set .Value on bound report controls)
```vba
If PSDue_ShowText = 1 Then
    PSDueTxt.Visible = True
    PSDueTxt.Caption = PSDue_Display
Else
    PSDueFmt.Visible = True
    Call ApplyColorFormatting(PSDueFmt, PSDue_Color)
    PSStrikeThru.Visible = (PSDue_Color = "STRIKETHROUGH")
End If
```

### DateConsentFormsSigned

| Column | Type | Description |
|--------|------|-------------|
| `DateConsentFormsSigned` | DATE | Original date |
| `DateConsentFormsSigned_Display` | VARCHAR | Formatted for display |
| `DateConsentFormsSigned_Color` | VARCHAR | Color code |
| `DateConsentFormsSigned_ShowDate` | BIT | Show date (1) or text (0) |

**Threshold:** Red if > 12 months, Green if > 11 months

### DateBMMExpires

| Column | Type | Description |
|--------|------|-------------|
| `DateBMMExpires` | DATE | Original date |
| `DateBMMExpires_Display` | VARCHAR | Formatted for display |
| `DateBMMExpires_Color` | VARCHAR | Color code |
| `DateBMMExpires_ShowDate` | BIT | Show date (1) or text (0) |

### DateSPDAuthExpires

| Column | Type | Description |
|--------|------|-------------|
| `DateSPDAuthExpires` | DATE | Original date |
| `DateSPDAuthExpires_Display` | VARCHAR | Formatted for display |
| `DateSPDAuthExpires_Color` | VARCHAR | Color code |
| `DateSPDAuthExpires_ShowDate` | BIT | Show date (1) or text (0) |

### DateSignaturesDueBy

| Column | Type | Description |
|--------|------|-------------|
| `DateSignaturesDueBy` | DATE | Original date |
| `DateSignaturesDueBy_Display` | VARCHAR | Formatted for display |
| `DateSignaturesDueBy_Color` | VARCHAR | Color code |
| `DateSignaturesDueBy_ShowDate` | BIT | Show date (1) or text (0) |

---

## DAY & HOUSE FIELDS (RecordType = 'House')

### LastVehicleChecklistCompleted

**Two versions:** One for Day program, one for Residential

| Column | Program | Type | Description |
|--------|---------|------|-------------|
| `LastVehicleChecklistCompleted_Display_Day` | Day | VARCHAR | Formatted for Day program |
| `LastVehicleChecklistCompleted_Color_Day` | Day | VARCHAR | Color for Day program |
| `LastVehicleChecklistCompleted_ShowDate_Day` | Day | BIT | Show date (1) or text (0) |
| `LastVehicleChecklistCompleted_Display_Res` | Res | VARCHAR | Formatted for Residential |
| `LastVehicleChecklistCompleted_Color_Res` | Res | VARCHAR | Color for Residential |
| `LastVehicleChecklistCompleted_ShowDate_Res` | Res | BIT | Show date (1) or text (0) |

**VBA Usage (Day):**
```vba
Call FormatExpirationField(LastVehicleChecklistCompleted_ShowDate_Day, _
                           LastVehicleChecklistCompleted_Display_Day, _
                           LastVehicleChecklistCompleted_Color_Day, _
                           LastVehicleChecklistCompletedTxt, _
                           LastVehicleChecklistCompletedFmt)
```

**VBA Usage (Residential):**
```vba
Call FormatExpirationField(LastVehicleChecklistCompleted_ShowDate_Res, _
                           LastVehicleChecklistCompleted_Display_Res, _
                           LastVehicleChecklistCompleted_Color_Res, _
                           LastVehicleChecklistCompletedTxt, _
                           LastVehicleChecklistCompletedFmt)
```

### DAYStaffTrainedInPrivacyBefore

| Column | Type | Description |
|--------|------|-------------|
| `DAYStaffTrainedInPrivacyBefore_Display` | VARCHAR | Formatted for display |
| `DAYStaffTrainedInPrivacyBefore_Color` | VARCHAR | Color code |
| `DAYStaffTrainedInPrivacyBefore_ShowDate` | BIT | Show date (1) or text (0) |

### DAYAllPlansReviewedByStaffBefore

| Column | Type | Description |
|--------|------|-------------|
| `DAYAllPlansReviewedByStaffBefore_Display` | VARCHAR | Formatted for display |
| `DAYAllPlansReviewedByStaffBefore_Color` | VARCHAR | Color code |
| `DAYAllPlansReviewedByStaffBefore_ShowDate` | BIT | Show date (1) or text (0) |

### DAYQtrlySafetyChecklistDueBy

| Column | Type | Description |
|--------|------|-------------|
| `DAYQtrlySafetyChecklistDueBy_Display` | VARCHAR | Formatted for display |
| `DAYQtrlySafetyChecklistDueBy_Color` | VARCHAR | Color code |
| `DAYQtrlySafetyChecklistDueBy_ShowDate` | BIT | Show date (1) or text (0) |

### MostRecentAsleepFireDrill (Residential only)

| Column | Type | Description |
|--------|------|-------------|
| `MostRecentAsleepFireDrill_Display` | VARCHAR | Formatted for display |
| `MostRecentAsleepFireDrill_Color` | VARCHAR | Color based on 14-month expiry |
| `MostRecentAsleepFireDrill_ShowDate` | BIT | Show date (1) or text (0) |

**Note:** Calculates based on `DATEADD(month, 14, MostRecentAsleepFireDrill)`

### HousePlansReviewedByStaffBefore (Residential)

| Column | Type | Description |
|--------|------|-------------|
| `HousePlansReviewedByStaffBefore_Display` | VARCHAR | Formatted for display |
| `HousePlansReviewedByStaffBefore_Color` | VARCHAR | Color code |
| `HousePlansReviewedByStaffBefore_ShowDate` | BIT | Show date (1) or text (0) |

### HouseSafetyPlanExpires (Residential)

| Column | Type | Description |
|--------|------|-------------|
| `HouseSafetyPlanExpires_Display` | VARCHAR | Formatted for display |
| `HouseSafetyPlanExpires_Color` | VARCHAR | Color code |
| `HouseSafetyPlanExpires_ShowDate` | BIT | Show date (1) or text (0) |

### MAPChecklistCompleted (Residential)

| Column | Type | Description |
|--------|------|-------------|
| `MAPChecklistCompleted_Display` | VARCHAR | Formatted for display |
| `MAPChecklistCompleted_Color` | VARCHAR | Color code |
| `MAPChecklistCompleted_ShowDate` | BIT | Show date (1) or text (0) |

### HRO Training Fields

**Two versions:** One for Day program, one for Residential

#### HROTrainsStaffBefore

| Column | Program | Description |
|--------|---------|-------------|
| `HROTrainsStaffBefore_Display_Day` | Day | Display value |
| `HROTrainsStaffBefore_Color_Day` | Day | Color code |
| `HROTrainsStaffBefore_ShowDate_Day` | Day | Show flag |
| `HROTrainsStaffBefore_Display_Res` | Res | Display value |
| `HROTrainsStaffBefore_Color_Res` | Res | Color code |
| `HROTrainsStaffBefore_ShowDate_Res` | Res | Show flag |

#### HROTrainsIndividualsBefore

| Column | Program | Description |
|--------|---------|-------------|
| `HROTrainsIndividualsBefore_Display_Day` | Day | Display value |
| `HROTrainsIndividualsBefore_Color_Day` | Day | Color code |
| `HROTrainsIndividualsBefore_ShowDate_Day` | Day | Show flag |
| `HROTrainsIndividualsBefore_Display_Res` | Res | Display value |
| `HROTrainsIndividualsBefore_Color_Res` | Res | Color code |
| `HROTrainsIndividualsBefore_ShowDate_Res` | Res | Show flag |

### FSO Training Fields

**Two versions:** One for Day program, one for Residential

#### FSOTrainsStaffBefore

| Column | Program | Description |
|--------|---------|-------------|
| `FSOTrainsStaffBefore_Display_Day` | Day | Display value |
| `FSOTrainsStaffBefore_Color_Day` | Day | Color code |
| `FSOTrainsStaffBefore_ShowDate_Day` | Day | Show flag |
| `FSOTrainsStaffBefore_Display_Res` | Res | Display value |
| `FSOTrainsStaffBefore_Color_Res` | Res | Color code |
| `FSOTrainsStaffBefore_ShowDate_Res` | Res | Show flag |

#### FSOTrainsIndividualsBefore

| Column | Program | Description |
|--------|---------|-------------|
| `FSOTrainsIndividualsBefore_Display_Day` | Day | Display value |
| `FSOTrainsIndividualsBefore_Color_Day` | Day | Color code |
| `FSOTrainsIndividualsBefore_ShowDate_Day` | Day | Show flag |
| `FSOTrainsIndividualsBefore_Display_Res` | Res | Display value |
| `FSOTrainsIndividualsBefore_Color_Res` | Res | Color code |
| `FSOTrainsIndividualsBefore_ShowDate_Res` | Res | Show flag |

---

## NAME FORMATTING FIELDS

### HumanRightsOfficer

| Column | Type | Description | Example |
|--------|------|-------------|---------|
| `HumanRightsOfficer` | VARCHAR | Original (LastName, FirstName) | "Smith, John" |
| `HumanRightsOfficer_Formatted` | VARCHAR | Formatted (FirstName LastName) | "John Smith" |
| `HumanRightsOfficer_IsBlank` | BIT | 1 = blank/invalid, 0 = has value | 0, 1 |

**VBA Usage:**
```vba
HumanRightsOfficerFmt.Value = HumanRightsOfficer_Formatted
If HumanRightsOfficer_IsBlank = 1 Then
    HumanRightsOfficerFmt.BackColor = RGB(255, 0, 0)  ' Red background
Else
    HumanRightsOfficerFmt.BackColor = RGB(255, 255, 255)  ' White
End If
```

### FireSafetyOfficer

| Column | Type | Description | Example |
|--------|------|-------------|---------|
| `FireSafetyOfficer` | VARCHAR | Original (LastName, FirstName) | "Doe, Jane" |
| `FireSafetyOfficer_Formatted` | VARCHAR | Formatted (FirstName LastName) | "Jane Doe" |
| `FireSafetyOfficer_IsBlank` | BIT | 1 = blank/invalid, 0 = has value | 0, 1 |

**VBA Usage:**
```vba
FireSafetyOfficerFmt.Value = FireSafetyOfficer_Formatted
If FireSafetyOfficer_IsBlank = 1 Then
    FireSafetyOfficerFmt.BackColor = RGB(255, 0, 0)  ' Red background
Else
    FireSafetyOfficerFmt.BackColor = RGB(255, 255, 255)  ' White
End If
```

---

## SPECIAL DATE VALUES

These special dates indicate non-standard states:

| Constant | Value | Meaning | Display As |
|----------|-------|---------|------------|
| `ExpMissing` | 1900-01-01 | Required data is missing | "Missing" (RED) |
| `ExpOptional` | 1900-01-02 | Field is optional | "Optional" (NORMAL) |
| `ExpNA` | 1900-01-03 | Field is not applicable | "N/A" (NORMAL) |
| `ExpCompleted` | 1900-01-04 | Task is completed | "Completed" (NORMAL) |
| `ExpPending` | 1900-01-05 | Task is pending | "Pending" (NORMAL) |

**The view automatically handles these special values** - you don't need to check for them in VBA.

---

## COLOR CODES

| Color Code | RGB Value | Font Weight | Border | Use Case |
|------------|-----------|-------------|--------|----------|
| `RED` | RGB(255, 0, 0) | 700 (Bold) | 1 (Solid) | Expired or critical |
| `GREEN` | RGB(18, 94, 40) | 700 (Bold) | 1 (Solid) | Warning threshold |
| `NORMAL` | RGB(0, 0, 0) | 400 (Normal) | 0 (None) | Normal/OK |
| `STRIKETHROUGH` | RGB(0, 0, 0) | 400 (Normal) | 0 (None) | Past due (show strikethrough line) |

---

## REUSABLE VBA HELPER FUNCTIONS

### ApplyColorFormatting

```vba
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
        Case Else ' NORMAL or STRIKETHROUGH
            ctl.ForeColor = RGB(0, 0, 0)
            ctl.FontWeight = 400
            ctl.BorderStyle = 0
    End Select
End Sub
```

### FormatExpirationField

```vba
Private Sub FormatExpirationField(showDate As Integer, displayValue As String, _
                                  colorCode As String, txtLabel As Control, fmtDate As Control)
    If showDate = 1 Then
        ' Show date: fmtDate is bound (ControlSource = field); do NOT set .Value (causes error in reports)
        txtLabel.Visible = False
        fmtDate.Visible = True
        Call ApplyColorFormatting(fmtDate, colorCode)
    Else
        ' Show the text label (Missing/Optional/N/A/Pending)
        txtLabel.Visible = True
        txtLabel.Caption = displayValue
        txtLabel.ForeColor = IIf(colorCode = "RED", RGB(255, 0, 0), RGB(0, 0, 0))
        txtLabel.FontWeight = IIf(colorCode = "RED", 700, 400)
        fmtDate.Visible = False
    End If
End Sub
```

---

## QUERY EXAMPLES

### Get all expired ISPs for active clients
```sql
SELECT 
    LastName,
    FirstName,
    DateISP_Display,
    DateISP_Color
FROM vw_ExpirationsFormatted
WHERE RecordType = 'Client' 
    AND DateISP_Color = 'RED'
    AND DateISP_ShowDate = 1
ORDER BY DateISP;
```

### Get houses with missing Human Rights Officers
```sql
SELECT 
    Location,
    HumanRightsOfficer_Formatted,
    HumanRightsOfficer_IsBlank
FROM vw_ExpirationsFormatted
WHERE RecordType = 'House' 
    AND HumanRightsOfficer_IsBlank = 1;
```

### Get all day program fields needing attention (RED)
```sql
SELECT 
    Location,
    LastName,
    FirstName,
    CASE WHEN LastVehicleChecklistCompleted_Color_Day = 'RED' THEN 'LastVehicleChecklist' END AS Field1,
    CASE WHEN DAYStaffTrainedInPrivacyBefore_Color = 'RED' THEN 'StaffPrivacyTraining' END AS Field2,
    CASE WHEN DAYAllPlansReviewedByStaffBefore_Color = 'RED' THEN 'PlansReview' END AS Field3
FROM vw_ExpirationsFormatted
WHERE RecordType = 'House'
    AND (LastVehicleChecklistCompleted_Color_Day = 'RED'
         OR DAYStaffTrainedInPrivacyBefore_Color = 'RED'
         OR DAYAllPlansReviewedByStaffBefore_Color = 'RED');
```

---

## TESTING CHECKLIST

When testing the view, verify:

- [ ] Client fields (DateISP, PSDue, DateConsentFormsSigned, DateBMMExpires, DateSPDAuthExpires, DateSignaturesDueBy)
- [ ] Day program fields (all _Day suffixed fields)
- [ ] Residential program fields (all _Res suffixed fields)
- [ ] Name formatting (HumanRightsOfficer_Formatted, FireSafetyOfficer_Formatted)
- [ ] Special dates ("Missing", "Optional", "N/A", "Pending") display correctly
- [ ] Red threshold triggers correctly
- [ ] Green threshold triggers correctly
- [ ] PSDue calculation (182 days before DateISP)
- [ ] Fire drill calculation (14 months after MostRecentAsleepFireDrill)
- [ ] Consent forms calculation (months based, not days)

---

## PERFORMANCE TIPS

1. **Always filter on RecordType first:**
   ```sql
   WHERE RecordType = 'Client' AND ...
   ```

2. **Use indexes on tblExpirations:**
   ```sql
   CREATE INDEX IX_tblExpirations_RecordType ON tblExpirations(RecordType);
   ```

3. **Avoid SELECT *:** Only select the columns you need

4. **Cache trigger values:** The view uses a CTE to cache trigger lookups - don't bypass the view

---

## MAINTENANCE

### Updating Trigger Values

```sql
-- Example: Change DateISP Green warning from 60 to 90 days
UPDATE catExpirationTriggers 
SET Green = 90 
WHERE Section = 'Individuals' AND FieldName = 'DateISP';
```

### Adding New Fields

To add a new expiration field:

1. Add trigger configuration:
   ```sql
   INSERT INTO catExpirationTriggers (Section, Program, FieldName, Red, Green, Description)
   VALUES ('Individuals', NULL, 'NewField', -1, 60, 'New field description');
   ```

2. Update the view to add calculations for the new field

3. Update VBA to bind to the new calculated columns

---

## SUPPORT

For questions or issues:
1. Check the implementation guide: `vw_ExpirationsFormatted_IMPLEMENTATION_GUIDE.md`
2. Check the performance analysis: `rptEXPIRATIONDATES_PERFORMANCE_ANALYSIS.md`
3. Review trigger configurations: `SELECT * FROM catExpirationTriggers;`
4. Test the view directly: `SELECT TOP 10 * FROM vw_ExpirationsFormatted;`
