# Expiration Report - Performance Optimization Implementation Guide

**Date:** January 30, 2026  
**Purpose:** Implement SQL-based pre-calculation to eliminate VBA performance bottlenecks  
**VBA Module:** `ModReportFieldManager.vba` - Automates field addition

> **Quick Start:** See `EXPIRATION_REPORT_OPTIMIZATION_README.md` for overview

---

## Overview

This guide walks you through implementing the pre-calculated view `vw_ExpirationsFormatted` to dramatically improve the performance of the Expirations report (`rptEXPIRATIONDATES`).

**Key Changes:**
- Move ALL date calculations from VBA to SQL Server
- Move ALL name formatting from VBA to SQL Server
- Simplify VBA to just read pre-calculated values and apply formatting
- Reduce network calls from thousands to dozens

**Expected Result:** 10x - 40x performance improvement (5-20 minutes → 30 seconds - 2 minutes)

---

## Step 1: Deploy SQL Objects to Azure SQL Server

### 1.1 Create the Trigger Configuration Table

Run this script first to create the `catExpirationTriggers` table:

```bash
# Connect to your Azure SQL Database and run:
sqlcmd -S tillsqlserver.database.windows.net -d TILLDBWEB_Prod -U tillsqladmin -P <password> -i create_catExpirationTriggers.sql
```

**Or via Azure Data Studio / SSMS:**
1. Connect to `tillsqlserver.database.windows.net`
2. Select database `TILLDBWEB_Prod`
3. Open file: `msaccess/extracted/sql/create_catExpirationTriggers.sql`
4. Execute script

**Verify:**
```sql
SELECT COUNT(*) FROM catExpirationTriggers;
-- Should return 31 rows (or adjust based on your configuration)
```

### 1.2 Create the Formatted View

Run this script to create the view:

```bash
sqlcmd -S tillsqlserver.database.windows.net -d TILLDBWEB_Prod -U tillsqladmin -P <password> -i vw_ExpirationsFormatted.sql
```

**Or via Azure Data Studio / SSMS:**
1. Open file: `msaccess/extracted/sql/vw_ExpirationsFormatted.sql`
2. Execute script

**Verify:**
```sql
-- Test the view
SELECT TOP 10 
    RecordType,
    LastName,
    FirstName,
    DateISP_Display,
    DateISP_Color,
    DateISP_ShowDate,
    HumanRightsOfficer_Formatted
FROM vw_ExpirationsFormatted
ORDER BY RecordType, LastName;
```

You should see:
- `DateISP_Display` shows either "Missing", "Optional", "N/A", or a formatted date
- `DateISP_Color` shows "RED", "GREEN", or "NORMAL"
- `DateISP_ShowDate` is 0 or 1
- `HumanRightsOfficer_Formatted` shows "FirstName LastName" format

---

## Step 2: Add Fields to Reports (VBA Automation)

### 2.1 Import VBA Module (2 minutes)

1. **Open MS Access database**
2. **Press Alt+F11** (VBA Editor)
3. **File → Import File...**
4. **Navigate to:** `msaccess/vba/`
5. **Select:** `ModReportFieldManager.vba`
6. **Click Open**

### 2.2 Run Implementation (1-2 minutes)

```vba
' In VBA Editor:
' 1. Find function: ImplementCompleteOptimization
' 2. Press F5 to run
' 3. Follow prompts
```

This automatically:
- ✅ Updates RecordSource to `vw_ExpirationsFormatted`
- ✅ Adds all required hidden fields (~77 total)
- ✅ Updates all 3 subreports

**Note:** The VBA module sets `ControlSource = fieldName` (not `="fieldName"`) - this is correct for programmatic assignment.

### 2.3 Verify

Open main report `rptEXPIRATIONDATES` - should open without "Can't find field" errors.

---

## Step 3: Review and Adjust Trigger Values

The default trigger values are samples. Review them with your business stakeholders:

```sql
-- View all trigger configurations
SELECT 
    ISNULL(Section, 'Staff') AS Area,
    ISNULL(Program, '--') AS Program,
    FieldName,
    Red AS RedThreshold,
    Green AS GreenThreshold,
    Description
FROM catExpirationTriggers
ORDER BY Area, Program, FieldName;
```

**Common Adjustments:**
- Change `Red = -1` to `Red = 0` if you want warnings on the expiration day (not after)
- Change `Green = 60` to `Green = 90` if you want longer warning periods
- Change `Red = 365` for checklists if annual reviews are acceptable

**Example - Modify DateISP Green threshold to 90 days:**
```sql
UPDATE catExpirationTriggers 
SET Green = 90 
WHERE Section = 'Individuals' AND FieldName = 'DateISP';
```

---

---

## Step 4: Simplify VBA Code in Subreports

**Note:** The VBA module (`ImplementCompleteOptimization`) already updated RecordSource for all subreports. You can skip manual RecordSource updates unless you need to verify or customize.

Now we need to replace complex VBA calculations with simple field binding.

### 4.1 Update Report_rptEXPIRATIONDATESclients

**BEFORE (Complex VBA - 121 lines):**

```vba
Private Sub Detail_Format(Cancel As Integer, FormatCount As Integer)
On Error GoTo 0
    Dim ErrorField As Variant
    
    ' Lots of complex date calculations...
    ErrorField = "DateISP": NextISPTxt.Visible = False: DateISPFmt.Visible = False
    Select Case DateISP
        Case ExpMissingCalculated, ExpOptionalCalculated, ExpNACalculated: 
            Call SetExpirationFieldProperties(NextISPTxt, , True, DateISP)
        Case Else
            DateISPFmt.Visible = Not IsEmpty(DateISP) And Not IsNull(DateISP) And (LengthN(8, DateISP) Or LengthN(10, DateISP))
            If DateISPFmt.Visible Then
                If (DateValue(DateISPFmt) - Int(Now)) < Trig_Indiv_ISP_Red Then
                    DateISPFmt.ForeColor = RGB(255, 0, 0): DateISPFmt.FontWeight = 700
                    ' ... more complex logic
                End If
            End If
    End Select
    ' ... repeat for 5 more fields (100+ more lines)
End Sub
```

**AFTER (Simple VBA - ~30 lines):**

```vba
Private Sub Detail_Format(Cancel As Integer, FormatCount As Integer)
On Error GoTo 0
    
    ' ========================================
    ' DateISP - Simple binding to pre-calculated view
    ' ========================================
    If DateISP_ShowDate = 1 Then
        ' Show the date field
        NextISPTxt.Visible = False
        DateISPFmt.Visible = True
        DateISPFmt.Value = DateISP_Display
        Call ApplyColorFormatting(DateISPFmt, DateISP_Color)
    Else
        ' Show the text label (Missing/Optional/N/A)
        NextISPTxt.Visible = True
        NextISPTxt.Caption = DateISP_Display
        NextISPTxt.ForeColor = IIf(DateISP_Color = "RED", RGB(255, 0, 0), RGB(0, 0, 0))
        NextISPTxt.FontWeight = IIf(DateISP_Color = "RED", 700, 400)
        DateISPFmt.Visible = False
    End If
    
    ' ========================================
    ' PSDue - Pre-calculated 182 days before DateISP
    ' ========================================
    If PSDue_ShowText = 1 Then
        PSDueTxt.Visible = True
        PSDueTxt.Caption = PSDue_Display
        PSDueFmt.Visible = False
        PSStrikeThru.Visible = False
    Else
        PSDueTxt.Visible = False
        PSDueFmt.Visible = True
        PSDueFmt.Value = PSDue_Display
        Call ApplyColorFormatting(PSDueFmt, PSDue_Color)
        PSStrikeThru.Visible = (PSDue_Color = "STRIKETHROUGH")
    End If
    
    ' ========================================
    ' Repeat for other fields (DateConsentFormsSigned, DateBMMExpires, etc.)
    ' Total: ~30 lines instead of 121
    ' ========================================
    
    ' DateConsentFormsSigned
    Call FormatExpirationField(DateConsentFormsSigned_ShowDate, DateConsentFormsSigned_Display, _
                               DateConsentFormsSigned_Color, ConsentFormsTxt, DateConsentFormsSignedFmt)
    
    ' DateBMMExpires
    Call FormatExpirationField(DateBMMExpires_ShowDate, DateBMMExpires_Display, _
                               DateBMMExpires_Color, BMMExpiresTxt, DateBMMExpiresFmt)
    
    ' DateSPDAuthExpires
    Call FormatExpirationField(DateSPDAuthExpires_ShowDate, DateSPDAuthExpires_Display, _
                               DateSPDAuthExpires_Color, SPDAuthTxt, DateSPDAuthExpiresFmt)
    
    ' DateSignaturesDueBy
    Call FormatExpirationField(DateSignaturesDueBy_ShowDate, DateSignaturesDueBy_Display, _
                               DateSignaturesDueBy_Color, DateSignaturesDueByTxt, DateSignaturesDueByFmt)
    
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
        Case Else ' NORMAL
            ctl.ForeColor = RGB(0, 0, 0)
            ctl.FontWeight = 400
            ctl.BorderStyle = 0
    End Select
End Sub

' ========================================
' Helper function for standard field formatting
' ========================================
Private Sub FormatExpirationField(showDate As Integer, displayValue As String, _
                                  colorCode As String, txtLabel As Control, fmtDate As Control)
    If showDate = 1 Then
        txtLabel.Visible = False
        fmtDate.Visible = True
        fmtDate.Value = displayValue
        Call ApplyColorFormatting(fmtDate, colorCode)
    Else
        txtLabel.Visible = True
        txtLabel.Caption = displayValue
        txtLabel.ForeColor = IIf(colorCode = "RED", RGB(255, 0, 0), RGB(0, 0, 0))
        txtLabel.FontWeight = IIf(colorCode = "RED", 700, 400)
        fmtDate.Visible = False
    End If
End Sub
```

**Key Changes:**
1. No more date calculations (`DateValue()`, `DateAdd()`, `DateDiff()`)
2. No more conditional logic for special dates
3. No more trigger constant comparisons
4. Just read pre-calculated fields and apply simple formatting
5. Reusable helper functions reduce code duplication

### 4.2 Update Report_rptEXPIRATIONDATESday

**BEFORE (200 lines with complex name formatting):**

```vba
Private Sub Detail_Format(Cancel As Integer, FormatCount As Integer)
    ' ... 9 date field calculations (similar to clients)
    
    ' Complex name formatting
    If Not IsNull(HumanRightsOfficer) And Not IsEmpty(HumanRightsOfficer) Then
        FindComma = InStr(1, HumanRightsOfficer, ",", vbTextCompare)
        If FindComma > 0 Then
            HumanRightsOfficerFmt = CorrectProperNames(StrConv(HumanRightsOfficer, vbProperCase))
            LN = Left(HumanRightsOfficer, FindComma - 1): FN = Mid(HumanRightsOfficer, FindComma + 2, 255)
            FN = SpecialNames(FN): LN = SpecialNames(LN)
            HumanRightsOfficerFmt = FN & " " & LN
            ' ... more processing
        End If
    End If
    ' ... repeat for FireSafetyOfficer
End Sub
```

**AFTER (~50 lines):**

```vba
Private Sub Detail_Format(Cancel As Integer, FormatCount As Integer)
On Error GoTo 0
    
    ' Use program-specific fields (Day vs Res)
    ' All calculations already done in SQL!
    
    ' LastVehicleChecklistCompleted (Day program)
    Call FormatExpirationField(LastVehicleChecklistCompleted_ShowDate_Day, _
                               LastVehicleChecklistCompleted_Display_Day, _
                               LastVehicleChecklistCompleted_Color_Day, _
                               LastVehicleChecklistCompletedTxt, _
                               LastVehicleChecklistCompletedFmt)
    
    ' DAYStaffTrainedInPrivacyBefore
    Call FormatExpirationField(DAYStaffTrainedInPrivacyBefore_ShowDate, _
                               DAYStaffTrainedInPrivacyBefore_Display, _
                               DAYStaffTrainedInPrivacyBefore_Color, _
                               DAYStaffTrainedInPrivacyBeforeTxt, _
                               DAYStaffTrainedInPrivacyBeforeFmt)
    
    ' ... repeat for other Day fields
    
    ' HumanRightsOfficer - SIMPLE!
    HumanRightsOfficerFmt.Visible = True
    HumanRightsOfficerFmt.Value = HumanRightsOfficer_Formatted
    HumanRightsOfficerFmt.FontWeight = 400
    If HumanRightsOfficer_IsBlank = 1 Then
        HumanRightsOfficerFmt.BackColor = RGB(255, 0, 0) ' Red background
    Else
        HumanRightsOfficerFmt.BackColor = RGB(255, 255, 255) ' White
    End If
    
    ' FireSafetyOfficer - SIMPLE!
    FireSafetyOfficerFmt.Visible = True
    FireSafetyOfficerFmt.Value = FireSafetyOfficer_Formatted
    FireSafetyOfficerFmt.FontWeight = 400
    If FireSafetyOfficer_IsBlank = 1 Then
        FireSafetyOfficerFmt.BackColor = RGB(255, 0, 0) ' Red background
    Else
        FireSafetyOfficerFmt.BackColor = RGB(255, 255, 255) ' White
    End If
    
    ' ... HRO and FSO training dates (use helper function)
    
End Sub
```

### 4.3 Update Report_rptEXPIRATIONDATEShouse

Similar approach - use the `_Res` (Residential) suffixed fields from the view:

```vba
Private Sub Detail_Format(Cancel As Integer, FormatCount As Integer)
On Error GoTo 0
    
    ' LastVehicleChecklistCompleted (Residential program)
    Call FormatExpirationField(LastVehicleChecklistCompleted_ShowDate_Res, _
                               LastVehicleChecklistCompleted_Display_Res, _
                               LastVehicleChecklistCompleted_Color_Res, _
                               LastVehicleChecklistCompletedTxt, _
                               LastVehicleChecklistCompletedFmt)
    
    ' MostRecentAsleepFireDrill (Residential only)
    Call FormatExpirationField(MostRecentAsleepFireDrill_ShowDate, _
                               MostRecentAsleepFireDrill_Display, _
                               MostRecentAsleepFireDrill_Color, _
                               MostRecentAsleepFireDrillTxt, _
                               MostRecentAsleepFireDrillFmt)
    
    ' ... repeat for other Residential fields
    
    ' Names (same as Day program)
    HumanRightsOfficerFmt.Value = HumanRightsOfficer_Formatted
    FireSafetyOfficerFmt.Value = FireSafetyOfficer_Formatted
    ' ... apply blank field formatting
    
End Sub
```

---

## Step 5: Test the Changes

### 5.1 Test with Small Dataset

1. Run the stored procedure `spApp_RunExpirationReport` to populate `tblExpirations`
2. Query the view to verify data:
   ```sql
   SELECT * FROM vw_ExpirationsFormatted WHERE RecordType = 'Client' AND LastName = 'Smith';
   ```
3. Open the report in MS Access - **Print Preview**
4. Verify:
   - Dates display correctly
   - Colors are correct (Red/Green/Normal)
   - "Missing", "Optional", "N/A" display correctly
   - Names are formatted (FirstName LastName)

### 5.2 Performance Test

**Measure BEFORE (baseline):**
1. Time the report generation: `Note start time → Run report → Note end time`
2. Record the time taken

**Measure AFTER (with new view):**
1. Time the report generation again
2. Compare times

**Expected:**
- **Before:** 5-20 minutes
- **After:** 30 seconds - 2 minutes
- **Improvement:** 10x - 40x faster

### 5.3 Validate Business Logic

Spot-check several records to ensure:
- Red flags appear at the correct thresholds
- Green warnings appear at the correct thresholds
- Calculated fields (like PSDue = DateISP - 182 days) are correct
- Special dates ("Missing", "Optional") display correctly

---

## Step 6: Add Indexes for Further Optimization

After verifying the solution works, add indexes to speed up the view queries:

```sql
-- Index on tblExpirations for faster filtering
CREATE NONCLUSTERED INDEX IX_tblExpirations_RecordType_Location
ON tblExpirations (RecordType, Location)
INCLUDE (LastName, FirstName, DateISP, DateConsentFormsSigned, 
         DateBMMExpires, DateSPDAuthExpires, DateSignaturesDueBy,
         LastVehicleChecklistCompleted, MostRecentAsleepFireDrill,
         HumanRightsOfficer, FireSafetyOfficer);

-- Index on tblExpirations for report sorting
CREATE NONCLUSTERED INDEX IX_tblExpirations_Location_RecordType_Name
ON tblExpirations (Location, RecordType, LastName, FirstName);
```

**Verify index usage:**
```sql
-- Check if indexes are being used
SET STATISTICS IO ON;
SELECT * FROM vw_ExpirationsFormatted WHERE RecordType = 'Client';
SET STATISTICS IO OFF;
```

---

## Step 7: Update Documentation

### 7.1 Document the Changes

Update your internal documentation to note:
- The report now uses `vw_ExpirationsFormatted` instead of `tblExpirations`
- Trigger values are managed in `catExpirationTriggers` table (not VBA constants)
- To change warning thresholds, update the table (not VBA code)
- VBA code is now simpler and easier to maintain

### 7.2 Train Administrators

Show administrators how to adjust trigger values:

```sql
-- Example: Change DateISP Green warning from 60 to 90 days
UPDATE catExpirationTriggers 
SET Green = 90 
WHERE Section = 'Individuals' AND FieldName = 'DateISP';

-- View current configuration
SELECT * FROM catExpirationTriggers ORDER BY Section, Program, FieldName;
```

---

## Troubleshooting

### Issue: View returns no data

**Cause:** `tblExpirations` is empty or `catExpirationTriggers` is not populated

**Fix:**
```sql
-- Check if tblExpirations has data
SELECT COUNT(*) FROM tblExpirations;

-- Check if catExpirationTriggers has data
SELECT COUNT(*) FROM catExpirationTriggers;

-- If empty, run:
EXEC spApp_RunExpirationReport;
-- And re-run create_catExpirationTriggers.sql
```

### Issue: Colors are wrong

**Cause:** Trigger values are incorrect or date calculations are off

**Fix:**
1. Verify trigger values match your business rules:
   ```sql
   SELECT * FROM catExpirationTriggers WHERE FieldName = 'DateISP';
   ```
2. Test the date calculation manually:
   ```sql
   SELECT 
       DateISP,
       DATEDIFF(day, CAST(GETDATE() AS DATE), DateISP) AS DaysUntilExpiration,
       CASE 
           WHEN DATEDIFF(day, CAST(GETDATE() AS DATE), DateISP) < -1 THEN 'RED'
           WHEN DATEDIFF(day, CAST(GETDATE() AS DATE), DateISP) <= 60 THEN 'GREEN'
           ELSE 'NORMAL'
       END AS ExpectedColor,
       DateISP_Color AS ActualColor
   FROM vw_ExpirationsFormatted
   WHERE RecordType = 'Client' AND DateISP_ShowDate = 1;
   ```

### Issue: Names not formatting correctly

**Cause:** Names in the database don't have a comma, or have unexpected format

**Fix:**
1. Check the raw data:
   ```sql
   SELECT HumanRightsOfficer, HumanRightsOfficer_Formatted 
   FROM vw_ExpirationsFormatted 
   WHERE HumanRightsOfficer IS NOT NULL;
   ```
2. If names are in different formats, update the view's name formatting logic

### Issue: VBA compile error after changes

**Cause:** Field names in VBA don't match view column names

**Fix:**
1. Verify view column names:
   ```sql
   SELECT TOP 1 * FROM vw_ExpirationsFormatted;
   ```
2. Update VBA field references to match exactly (case-sensitive)

### Issue: Report is still slow

**Cause:** View not being used, or indexes not in place

**Fix:**
1. Verify RecordSource is set to `vw_ExpirationsFormatted`:
   - Open report in Design View
   - Check Property Sheet → Record Source
2. Add indexes (see Step 6)
3. Check if VBA still has complex calculations (should be simple binding only)

---

## Rollback Plan

If you need to rollback:

1. **Revert Report RecordSource:**
   - Change from `vw_ExpirationsFormatted` back to `qrytblExpirations`
   
2. **Restore Original VBA:**
   - Use backup copy of VBA modules (from before changes)
   - Or retrieve from source control
   
3. **Remove SQL Objects (optional):**
   ```sql
   DROP VIEW IF EXISTS vw_ExpirationsFormatted;
   -- Keep catExpirationTriggers for future use
   ```

---

## Next Steps

After successful implementation:

1. **Apply same pattern to other slow reports:**
   - `rptSTAFFWITHNOSKILLS`
   - Staff evaluation reports
   - Any reports with VBA calculations in Detail_Format

2. **Monitor performance:**
   - Track report generation times
   - Monitor Azure SQL Server DTU usage during report runs
   - Adjust indexes if needed

3. **Consider migrating to modern reporting:**
   - Power BI for interactive dashboards
   - SSRS for server-side report generation
   - Azure Data Studio for ad-hoc queries

---

## Summary

**What We Did:**
1. Created `catExpirationTriggers` table to store threshold values
2. Created `vw_ExpirationsFormatted` view to pre-calculate all formatting
3. Updated report RecordSource to use the view
4. Simplified VBA from 100+ lines to ~30 lines per subreport
5. Added indexes for faster queries

**Benefits:**
- ✅ 10x - 40x faster report generation
- ✅ Minimal network calls (bulk data transfer)
- ✅ Server-side processing (Azure SQL does the work)
- ✅ Easier to maintain (SQL is easier to debug than VBA)
- ✅ Configurable thresholds (database table, not hardcoded VBA)
- ✅ Consistent formatting logic across all reports

**Result:**
Your Expirations report should now run in 30 seconds to 2 minutes instead of 5-20 minutes!
