# Expiration Dates Report - Performance Analysis

**Report:** `rptEXPIRATIONDATES` with subreports  
**Date:** January 30, 2026 · **Last updated:** February 2, 2026  
**Issue:** Report takes excessively long to run when connected to Azure SQL  

---

## Executive Summary

The report is slow because **MS Access is performing thousands of calculations in VBA for every detail record**, while pulling data over the network from Azure SQL. Each calculation may trigger additional network round trips, creating a massive performance bottleneck.

**Root Cause:** The Detail_Format event in the subreports executes complex VBA code (date calculations, string formatting, conditional logic) for every single record displayed in the report. With Azure SQL, this means:
- Network latency for each field access
- Client-side processing instead of server-side
- No optimization or caching

**Impact:** A report with 100 records and four subreports could trigger **thousands of VBA operations and network calls**.

---

## Specific Performance Issues

### 1. **Detail_Format Event Processing (CRITICAL)**

All four subreports (clients, day, house, staff) execute extensive VBA code in their `Detail_Format` event:

#### `rptEXPIRATIONDATESclients.vba` (121 lines)
- **6 date fields** with complex calculations per record
- Each field checks special values (Missing/Optional/N/A)
- Date arithmetic: `DateValue()`, `DateAdd()`, `Int(Now)`, `DateDiff()`
- Color/formatting: RGB values, font weights, border styles
- **Example for ONE field (DateISP):**
  ```vba
  Select Case DateISP
      Case ExpMissingCalculated, ExpOptionalCalculated, ExpNACalculated: 
          Call SetExpirationFieldProperties(NextISPTxt, , True, DateISP)
      Case Else
          DateISPFmt.Visible = Not IsEmpty(DateISP) And Not IsNull(DateISP) And (LengthN(8, DateISP) Or LengthN(10, DateISP))
          If DateISPFmt.Visible Then
              If (DateValue(DateISPFmt) - Int(Now)) < Trig_Indiv_ISP_Red Then
                  DateISPFmt.ForeColor = RGB(255, 0, 0): DateISPFmt.FontWeight = 700
              ElseIf (DateValue(DateISPFmt) - Int(Now)) <= Trig_Indiv_ISP_Green Then
                  DateISPFmt.ForeColor = RGB(18, 94, 40): DateISPFmt.FontWeight = 700
              Else
                  DateISPFmt.ForeColor = RGB(0, 0, 0): DateISPFmt.FontWeight = 400
              End If
          End If
  End Select
  ```
- **This repeats for EACH of 6 fields, for EVERY client record!**

#### `rptEXPIRATIONDATESday.vba` (200 lines)
- **9 date fields** with calculations per record
- **Heavy string manipulation** for name formatting:
  - `InStr()`, `StrConv()`, `Left()`, `Mid()` for HumanRightsOfficer
  - `InStr()`, `StrConv()`, `Left()`, `Mid()` for FireSafetyOfficer
  - Calls to `SpecialNames()`, `CorrectProperNames()`, `CheckBlankField()`
- Date arithmetic and conditional formatting for 9 fields
- **Example name processing (repeated twice per record):**
  ```vba
  If Not IsNull(HumanRightsOfficer) And Not IsEmpty(HumanRightsOfficer) Then
      FindComma = InStr(1, HumanRightsOfficer, ",", vbTextCompare)
      If FindComma > 0 Then
          HumanRightsOfficerFmt = CorrectProperNames(StrConv(HumanRightsOfficer, vbProperCase))
          LN = Left(HumanRightsOfficer, FindComma - 1)
          FN = Mid(HumanRightsOfficer, FindComma + 2, 255)
          FN = SpecialNames(FN)
          LN = SpecialNames(LN)
          HumanRightsOfficerFmt = FN & " " & LN
          ' ... more processing
      End If
  End If
  ```

#### `rptEXPIRATIONDATEShouse.vba` (218 lines)
- **10 date fields** with calculations per record
- **Heavy string manipulation** for name formatting (same as day)
- Duplicate processing for HumanRightsOfficer and FireSafetyOfficer
- Date arithmetic and conditional formatting for 10 fields

### 2. **Network Latency Issues**

**Current Architecture:**
```
Azure SQL Server (data) 
    ↓ Network
MS Access (processing)
    ↓ For each record...
    ↓ VBA reads field → Network call
    ↓ VBA calculates → Client CPU
    ↓ VBA formats → Client CPU
    ↓ Repeat for next field...
```

**Problems:**
- Each field reference in VBA may trigger a network fetch from Azure SQL
- No bulk operations or caching
- All processing happens client-side
- Latency amplified by record count × field count × calculations

### 3. **No Pre-calculated Columns**

The stored procedure `spApp_RunExpirationReport` correctly populates `tblExpirations` with **raw data only**:
- Date fields are just dates
- No color/formatting flags
- No visibility indicators
- No calculated text (e.g., "Missing", "Optional", "N/A")

**Result:** The report must calculate everything at runtime in VBA.

### 4. **Trigger Constants (Global Variables)**

The VBA code references global trigger constants like:
- `Trig_Indiv_ISP_Red`
- `Trig_Indiv_ISP_Green`
- `Trig_Day_LVC_Red`
- `Trig_Res_HROTS_Red`
- etc.

These are accessed repeatedly for every field calculation on every record.

---

## Performance Impact Calculation

**Example Scenario:**
- 50 locations in report
- Average 3 subreport types per location
- Average 20 client records per location
- Average 6-10 date fields per subreport

**Operations:**
- `rptEXPIRATIONDATESclients`: 20 records × 6 fields × 5 operations = **600 operations**
- `rptEXPIRATIONDATESday`: 50 records × 9 fields × 6 operations = **2,700 operations**
- `rptEXPIRATIONDATEShouse`: 50 records × 10 fields × 6 operations = **3,000 operations**

**Total: ~6,300 VBA operations**, each potentially causing:
- Network round trip (10-100ms each)
- VBA function call overhead
- String manipulation/date calculation

**Estimated time:** 6,300 × 50ms = **315 seconds (5+ minutes)**

With more records or slower network, this could easily be 10-20 minutes or more.

---

## Recommended Solutions

### Solution 1: Move Calculations to SQL Server (RECOMMENDED)

**Create a new view or modify the stored procedure** to pre-calculate all formatting data:

```sql
CREATE OR ALTER VIEW vw_ExpirationsFormatted
AS
SELECT 
    e.*,
    
    -- Pre-calculate visibility and display values for DateISP
    CASE 
        WHEN e.DateISP = '1899-12-30' THEN 'Missing'
        WHEN e.DateISP = '1900-01-01' THEN 'Optional'
        WHEN e.DateISP = '1900-01-02' THEN 'N/A'
        WHEN e.DateISP IS NULL THEN 'Missing'
        ELSE FORMAT(e.DateISP, 'MM/dd/yy')
    END AS DateISP_Display,
    
    CASE
        WHEN e.DateISP IN ('1899-12-30', NULL) THEN 'RED'
        WHEN e.DateISP IN ('1900-01-01', '1900-01-02') THEN 'NORMAL'
        WHEN DATEDIFF(day, GETDATE(), e.DateISP) < @Trig_Indiv_ISP_Red THEN 'RED'
        WHEN DATEDIFF(day, GETDATE(), e.DateISP) <= @Trig_Indiv_ISP_Green THEN 'GREEN'
        ELSE 'NORMAL'
    END AS DateISP_Color,
    
    CASE 
        WHEN e.DateISP IN ('1899-12-30', '1900-01-01', '1900-01-02', NULL) THEN 0
        ELSE 1
    END AS DateISP_ShowDate,
    
    -- Pre-calculate PSDue
    CASE 
        WHEN e.DateISP NOT IN ('1899-12-30', '1900-01-01', '1900-01-02') 
            AND e.DateISP IS NOT NULL
        THEN DATEADD(day, -182, e.DateISP)
        ELSE NULL
    END AS PSDue_Calculated,
    
    -- Pre-format names (HumanRightsOfficer, FireSafetyOfficer)
    CASE 
        WHEN CHARINDEX(',', e.HumanRightsOfficer) > 0
        THEN LTRIM(RTRIM(SUBSTRING(e.HumanRightsOfficer, CHARINDEX(',', e.HumanRightsOfficer) + 1, 255))) + ' ' +
             LTRIM(RTRIM(SUBSTRING(e.HumanRightsOfficer, 1, CHARINDEX(',', e.HumanRightsOfficer) - 1)))
        ELSE NULL
    END AS HumanRightsOfficer_Formatted,
    
    -- Repeat for ALL fields...
    -- (Similar logic for DateConsentFormsSigned, DateBMMExpires, etc.)
    
FROM tblExpirations e
WHERE e.RecordType IN ('Client', 'Staff', 'House')
```

**Benefits:**
- All calculations done once on the SQL Server (fast)
- Results cached and transferred in bulk
- No VBA processing needed
- Report just binds to pre-calculated columns

**VBA Changes:**
```vba
Private Sub Detail_Format(Cancel As Integer, FormatCount As Integer)
On Error GoTo 0
    ' Simple binding - no calculations!
    If DateISP_ShowDate = 1 Then
        DateISPFmt.Visible = True
        DateISPFmt.Caption = DateISP_Display
        Select Case DateISP_Color
            Case "RED": DateISPFmt.ForeColor = RGB(255, 0, 0): DateISPFmt.FontWeight = 700
            Case "GREEN": DateISPFmt.ForeColor = RGB(18, 94, 40): DateISPFmt.FontWeight = 700
            Case Else: DateISPFmt.ForeColor = RGB(0, 0, 0): DateISPFmt.FontWeight = 400
        End Select
    Else
        DateISPFmt.Visible = False
        NextISPTxt.Visible = True
        NextISPTxt.Caption = DateISP_Display
    End If
End Sub
```

**Estimated Performance Improvement: 90-95% faster**

### Solution 2: Use Snapshot/Local Table (INTERIM FIX)

If you cannot modify SQL immediately, create a local Access table as a snapshot:

```vba
Public Function RunExpirationsReport(ExpDatesReportInitiatedFromReportsMenu As Boolean) As Boolean
    ' ... existing code ...
    
    ' AFTER stored procedure runs and populates tblExpirations:
    Call AppendProgressMessages("Creating local snapshot for faster rendering...")
    
    ' Delete existing local snapshot
    If IsTableQuery("tblExpirations_Local") Then
        DoCmd.DeleteObject acTable, "tblExpirations_Local"
    End If
    
    ' Create local copy
    DoCmd.RunSQL "SELECT * INTO tblExpirations_Local FROM tblExpirations"
    
    ' Update report RecordSource to use local table
    ' (This minimizes network calls during Detail_Format)
    
    ' ... continue with report generation ...
End Function
```

**Benefits:**
- Reduces network calls during rendering
- Data loaded once in bulk
- VBA still slow, but field access is local

**Estimated Performance Improvement: 50-70% faster**

### Solution 3: Create Indexes on Azure SQL (SUPPORTING FIX)

Add indexes to speed up the stored procedure:

```sql
-- Index on tblExpirations for report filtering
CREATE NONCLUSTERED INDEX IX_tblExpirations_RecordType_Location
ON tblExpirations (RecordType, Location)
INCLUDE (LastName, FirstName, DateISP, DateConsentFormsSigned);

-- Index on tblStaff for lookups
CREATE NONCLUSTERED INDEX IX_tblStaff_EMPLOYID_DEPRTMNT
ON tblStaff (EMPLOYID)
INCLUDE (LASTNAME, FRSTNAME, DEPRTMNT, SUPERVISORCODE_I);

-- Index on tblStaffSkills
CREATE NONCLUSTERED INDEX IX_tblStaffSkills_EMPID_SKILLNUMBER
ON tblStaffSkills (EMPID_I, SKILLNUMBER_I)
INCLUDE (EXPIREDSKILL_I);
```

**Benefits:**
- Faster stored procedure execution
- Faster queries during report binding

**Estimated Performance Improvement: 10-20% faster (on stored procedure execution)**

---

## Recommended Implementation Plan

### Phase 1 (Immediate - 1-2 hours)
1. **Create `vw_ExpirationsFormatted` view** with pre-calculated columns for:
   - All date formatting/visibility flags
   - Color codes (RED/GREEN/NORMAL)
   - Calculated fields (PSDue, formatted names)
   - Display text ("Missing", "Optional", "N/A")

2. **Update stored procedure** to store trigger thresholds in a config table:
   ```sql
   CREATE TABLE tblExpirationsConfig (
       ConfigKey VARCHAR(50) PRIMARY KEY,
       ConfigValue INT
   );
   
   INSERT INTO tblExpirationsConfig VALUES 
       ('Trig_Indiv_ISP_Red', -1),
       ('Trig_Indiv_ISP_Green', 60),
       ('Trig_Day_LVC_Red', 365),
       -- ... etc
   ```

3. **Simplify subreport VBA** to just read pre-calculated fields and apply formatting

### Phase 2 (Next Day - 2-4 hours)
1. **Update report RecordSource** to use `vw_ExpirationsFormatted`
2. **Test thoroughly** with production data
3. **Add indexes** on Azure SQL tables

### Phase 3 (Following Week - ongoing)
1. **Monitor performance** and adjust indexes
2. **Document changes** for future maintenance
3. **Apply same pattern** to other slow reports

---

## Expected Results

**Before:**
- Report generation: 5-20 minutes
- Network calls: Thousands
- CPU usage: High (client-side VBA)

**After (with Solution 1):**
- Report generation: 30 seconds - 2 minutes
- Network calls: Minimal (bulk data transfer)
- CPU usage: Low (server-side SQL)

**Performance Gain: 10x - 40x faster**

---

## Additional Notes

### Why the Stored Procedure Alone Isn't Enough

The stored procedure `spApp_RunExpirationReport` correctly populates `tblExpirations`, but it only stores **raw data**. The formatting logic (colors, visibility, special text) still happens in VBA at runtime.

**The fix is to move the formatting logic into SQL** so the report just displays pre-calculated values.

### Alternative: Power BI or SSRS

For long-term maintainability, consider migrating this report to:
- **Power BI**: Modern, fast, cloud-native, great for Azure SQL
- **SSRS (SQL Server Reporting Services)**: Server-side rendering, no client VBA
- **Azure Data Studio Reports**: Lightweight alternative

These tools process data server-side and eliminate network latency issues entirely.

---

## Conclusion

The report is slow because MS Access is doing thousands of VBA calculations over a network connection to Azure SQL. The solution is to **move all calculations to SQL Server** using views or computed columns, and simplify the VBA to just apply pre-calculated formatting.

Implementing Solution 1 (SQL-based pre-calculation) will provide immediate and dramatic performance improvements with minimal code changes.
