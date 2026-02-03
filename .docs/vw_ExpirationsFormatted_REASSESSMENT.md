# vw_ExpirationsFormatted and Reports – Reassessment

**Date:** February 2, 2026  
**Scope:** View `vw_ExpirationsFormatted`, table `catExpirationTriggers`, and the four expiration subreports (clients, day, house, staff) – original VBA vs. updated implementation using the view.

---

## 1. Executive Summary

| Item | Status | Notes |
|------|--------|--------|
| **vw_ExpirationsFormatted** | ✅ Covers all report fields | Client, Day, House, and Staff sections implemented. One SQL bug: `IN (..., null)` does not match NULL in T-SQL – fix applied. |
| **catExpirationTriggers** | ✅ Aligned | All trigger names and Section/Program match view and original GlobalVariables. |
| **rptEXPIRATIONDATESclients** | ✅ Replaceable | View provides all _Display/_Color/_ShowDate columns; updated VBA can bind only (no calculations). |
| **rptEXPIRATIONDATESday** | ✅ Replaceable | Day-specific columns + HumanRightsOfficer_Formatted, FireSafetyOfficer_Formatted. |
| **rptEXPIRATIONDATEShouse** | ✅ Replaceable | Res-specific columns + officer names. NextRecentAsleepFireDrill is raw in view via `e.*` (no _Display needed). |
| **rptEXPIRATIONDATESstaff** | ⚠️ Partial | View replaces most logic; 3MoEval and MAPCert N/A/120-day still in VBA (by design). |

**Conclusion:** The view correctly replaces the VBA formatting logic for clients, day, and house, and for most staff fields. Remaining VBA handles location-specific rules (TB Pending, MAPCert N/A/120-day, ThreeMonthEvaluation + 3MoEval) and FullName/name formatting.

---

## 2. View and catExpirationTriggers

### 2.1 vw_ExpirationsFormatted

- **Source:** `msaccess/extracted/sql/vw_ExpirationsFormatted.sql`
- **Depends on:** `tblExpirations`, `tblLocations`, `catExpirationTriggers`
- **Output:** All base columns from `tblExpirations` plus:
  - **Client:** DateISP, PSDue, DateConsentFormsSigned, DateBMMExpires, DateSPDAuthExpires, DateSignaturesDueBy → each with _Display, _Color, _ShowDate (PSDue uses _ShowText).
  - **Day/House:** LastVehicleChecklistCompleted (_Day/_Res), DAYStaffTrainedInPrivacyBefore, DAYAllPlansReviewedByStaffBefore, DAYQtrlySafetyChecklistDueBy, MostRecentAsleepFireDrill, HousePlansReviewedByStaffBefore, HouseSafetyPlanExpires, MAPChecklistCompleted, HRO/FSO training (_Day/_Res), HumanRightsOfficer_Formatted, FireSafetyOfficer_Formatted, _IsBlank.
  - **Staff:** BBP, BackInjuryPrevention, CPR, DefensiveDriving, DriversLicense, FirstAid, PBS, SafetyCares, TB, WheelchairSafety, WorkplaceViolence, ProfessionalLicenses, MAPCert, EvalDueBy, LastSupervision → _Display, _Color, _ShowDate.

**Fix applied:** In T-SQL, `column IN ('1900-01-01', null)` does **not** match rows where `column IS NULL`. The view was updated to use explicit `OR column IS NULL` (or `column = '1900-01-01' OR column IS NULL`) so Missing/NULL is handled the same as the original VBA.

### 2.2 catExpirationTriggers

- **Source:** `msaccess/extracted/sql/create_catExpirationTriggers.sql`
- **Usage:** View’s `TriggerValues` CTE reads Red/Green (and Section/Program) per field. Original VBA used `GlobalVariables` filled via `DLookup("Red", "catExpirationTriggers", ...)`. The view replaces those lookups with a single CTE; reports no longer need DLookup for standard fields.
- **Staff 3MoEval:** Still used in staff report VBA via `DLookup("Red", "catExpirationTriggers", "FieldName='3MoEval' AND Section IS NULL AND Program IS NULL")` for the ThreeMonthEvaluation blue/red logic. Optional future improvement: add EvalDueBy_3MoEval_Color (or similar) to the view to remove this DLookup.

---

## 3. Report-by-Report: Original vs. Updated

### 3.1 rptEXPIRATIONDATESclients

| Original VBA | View replacement | Binding (updated) |
|--------------|------------------|-------------------|
| DateISP: Case on ExpMissing/Optional/N/A, else show date; Red if (DateValue − Now) &lt; Trig_Indiv_ISP_Red, Green ≤ Trig_Indiv_ISP_Green | DateISP_Display, DateISP_Color, DateISP_ShowDate | Show DateISPFmt when ShowDate=1 and apply Color; else show NextISPTxt with Caption = Display |
| PSDue: DateAdd("d", -182, DateISP); Red/Green/Strikethrough by Trig_Indiv_PSDue_Green and past-due | PSDue_Display, PSDue_Color, PSDue_ShowText, PSDue_Calculated | Show PSDueFmt when ShowText=0, apply Color and PSStrikeThru when Color=STRIKETHROUGH |
| DateConsentFormsSigned | _Display, _Color, _ShowDate | Same pattern as DateISP |
| DateBMMExpires | _Display, _Color, _ShowDate | Same pattern |
| DateSPDAuthExpires | _Display, _Color, _ShowDate | Same pattern |
| DateSignaturesDueBy | _Display, _Color, _ShowDate | Same pattern |

**VBA to remove:** All Select Case blocks and Trig_* comparisons for these six fields.  
**VBA to keep:** Simple visibility and ApplyColorFormatting based on _Display/_Color/_ShowDate (and PSDue_ShowText, PSStrikeThru for PSDue).  
**ModReportFieldManager:** FixClientsReport() adds the hidden view columns; RecordSource = vw_ExpirationsFormatted.

---

### 3.2 rptEXPIRATIONDATESday

| Original VBA | View replacement | Binding (updated) |
|--------------|------------------|-------------------|
| LastVehicleChecklistCompleted: Case on Exp*; Red if (Now − DateValue) ≥ Trig_Day_LVC_Red | LastVehicleChecklistCompleted_Display_Day, _Color_Day, _ShowDate_Day | FormatExpirationField( _ShowDate_Day, _Display_Day, _Color_Day, Txt, Fmt ) |
| DAYStaffTrainedInPrivacyBefore, DAYAllPlansReviewedByStaffBefore, DAYQtrlySafetyChecklistDueBy | _Display, _Color, _ShowDate | Same pattern |
| HumanRightsOfficer: Parse "LastName, FirstName" → "FirstName LastName", CorrectProperNames, SpecialNames, CheckBlankField | HumanRightsOfficer_Formatted, HumanRightsOfficer_IsBlank | Bind Fmt to HumanRightsOfficer_Formatted; if IsBlank=1 then red backcolor (or hide) |
| FireSafetyOfficer | FireSafetyOfficer_Formatted, FireSafetyOfficer_IsBlank | Same |
| HROTrainsStaffBefore, HROTrainsIndividualsBefore, FSOTrainsStaffBefore, FSOTrainsIndividualsBefore | _Display_Day, _Color_Day, _ShowDate_Day | FormatExpirationField |

**Note:** View does not replicate VBA’s CorrectProperNames/SpecialNames; it only flips "LastName, FirstName" to "FirstName LastName". If exact casing/hyphen rules are required, that can stay in VBA or be extended in the view later.

**ModReportFieldManager:** FixDayReport() adds the Day and officer fields.

---

### 3.3 rptEXPIRATIONDATEShouse

| Original VBA | View replacement | Binding (updated) |
|-------------|------------------|-------------------|
| LastVehicleChecklistCompleted | _Display_Res, _Color_Res, _ShowDate_Res | Same pattern as Day |
| MostRecentAsleepFireDrill: Red if (DateAdd("m",14,...) − Now) &lt; Trig_Res_MRFD_Red, Green ≤ Trig_Res_MRFD_Green | MostRecentAsleepFireDrill_Display, _Color, _ShowDate | FormatExpirationField |
| NextRecentAsleepFireDrill | No _Display in view | Raw column in view via `e.*`; house report only sets NextRecentAsleepFireDrillFmt.Visible = Not IsEmpty(NextRecentAsleepFireDrill). No change needed. |
| HousePlansReviewedByStaffBefore, HouseSafetyPlanExpires, MAPChecklistCompleted | _Display, _Color, _ShowDate | FormatExpirationField |
| HumanRightsOfficer, FireSafetyOfficer | _Formatted, _IsBlank | Same as Day |
| HRO/FSO training | _Display_Res, _Color_Res, _ShowDate_Res | FormatExpirationField |

**ModReportFieldManager:** FixHouseReport() adds the Res and officer fields.

---

### 3.4 rptEXPIRATIONDATESstaff

| Original VBA | View replacement | Binding (updated) |
|-------------|------------------|-------------------|
| FullName = FirstName + LastName, dash-uppercase | Not in view | Keep in VBA |
| BBP, BackInjuryPrevention, CPR, DriversLicense, FirstAid, SafetyCares, TB, WorkplaceViolence, ProfessionalLicenses | _Display, _Color, _ShowDate | FormatExpirationField (updated VBA in rptEXPIRATIONDATESstaff_VBA_UPDATED.md) |
| DefensiveDriving, PBS, WheelchairSafety: show "Done" or date | _Display, _Color, _ShowDate | FormatExpirationField (Color always NORMAL in view) |
| TB Pending + Location = Hollis / ABI / Day Services → red TBTxt | View gives TB_Display = "Pending", TB_Color | VBA still applies location rule: if TB_ShowDate=0 and TB_Display="Pending" and (Hollis or ABI or Day Services) then show red TBTxt |
| MAPCert: N/A for certain DED locations; 120-day-from-AdjustedStartDate purple | View gives MAPCert_Display/_Color/_ShowDate | VBA keeps N/A and 120-day purple logic (location/AdjustedStartDate) |
| EvalDueBy: Red/Green by Trig_Staff_EVL_*; if ThreeMonthEvaluation then blue/red by Trig_Staff_3MO_Red | EvalDueBy_Display, EvalDueBy_Color, EvalDueBy_ShowDate | VBA applies EvalDueBy_Color; for ThreeMonthEvaluation still uses DLookup("Red", "catExpirationTriggers", "FieldName='3MoEval' ...") and blue formatting |
| LastSupervision: Red if (Now − LastSupervision) > Trig_Staff_SUP_Red | LastSupervision_Display, _Color, _ShowDate | View computes correctly; report keeps LastSupervisionFmt hidden in current design |

**ModReportFieldManager:** FixStaffReport() adds all Staff _Display/_Color/_ShowDate fields.

---

## 4. Gaps and Optional Improvements

1. **NULL handling in view**  
   All `WHEN col IN (..., null)` conditions were updated so NULL is explicitly handled (e.g. `OR col IS NULL` or equivalent). This matches original VBA behavior for Missing/null dates.

2. **3MoEval in view**  
   Staff report still uses DLookup on catExpirationTriggers for 3MoEval when ThreeMonthEvaluation is true. Optional: add a view column (e.g. EvalDueBy_3MoEval_Color or EvalDueBy_Color_3Mo) so VBA does not need DLookup.

3. **MAPCert N/A and 120-day purple**  
   Location-based N/A and 120-day-from-AdjustedStartDate are left in VBA; moving them into the view would require location and AdjustedStartDate in the view (already there via e.*) and business rules in SQL.

4. **HumanRightsOfficer_Formatted / FireSafetyOfficer_Formatted**  
   View does simple "LastName, FirstName" → "FirstName LastName". VBA also uses CorrectProperNames and SpecialNames; if those must stay consistent, keep that logic in VBA or document the small difference.

5. **NextRecentAsleepFireDrill**  
   No _Display/_Color in view; house report only needs the raw date for visibility. Already satisfied by `e.*`.

---

## 5. Checklist: Ensuring the View Replaces VBA Correctly

- [x] **catExpirationTriggers** created and populated; view uses it in TriggerValues CTE.
- [x] **RecordSource** of all four subreports set to vw_ExpirationsFormatted (via FixClientsReport, FixDayReport, FixHouseReport, FixStaffReport).
- [x] **Hidden controls** added for all _Display/_Color/_ShowDate (and PSDue_ShowText, officer _Formatted/_IsBlank) so report code can bind and format without recalculating.
- [x] **Clients:** No Trig_* or date math in VBA for DateISP, PSDue, DateConsentFormsSigned, DateBMMExpires, DateSPDAuthExpires, DateSignaturesDueBy; only visibility and ApplyColorFormatting (and PSDue STRIKETHROUGH).
- [x] **Day:** No Trig_* or date math for LVC, DAY*, HRO*, FSO*; officer names from _Formatted/_IsBlank (or VBA if keeping SpecialNames).
- [x] **House:** Same for Res fields and MostRecentAsleepFireDrill; NextRecentAsleepFireDrill from e.*.
- [x] **Staff:** All standard staff fields from view; VBA keeps FullName, TB Pending location rule, MAPCert N/A/120-day, EvalDueBy + ThreeMonthEvaluation/3MoEval DLookup, LastSupervision hidden.
- [x] **View NULL handling** fixed so Missing/null dates are treated like original VBA (explicit IS NULL in CASE conditions).

---

## 6. Files Reference

| File | Purpose |
|------|---------|
| `msaccess/extracted/sql/vw_ExpirationsFormatted.sql` | View definition (with NULL fix) |
| `msaccess/extracted/sql/create_catExpirationTriggers.sql` | Trigger table and defaults |
| `msaccess/vba/ModReportFieldManager.vba` | Fix*Report(), UpdateRecordSource(), AddHiddenFieldsToReport() |
| `.docs/vw_ExpirationsFormatted_IMPLEMENTATION_GUIDE.md` | Step-by-step deployment (clients, day, house) |
| `.docs/vw_ExpirationsFormatted_FIELD_REFERENCE.md` | Column list and VBA usage |
| `.docs/rptEXPIRATIONDATESstaff_VBA_UPDATED.md` | Staff report updated Detail_Format and helpers |
| `msaccess/extracted/vba/Report_rptEXPIRATIONDATESclients.vba` | Original clients report VBA |
| `msaccess/extracted/vba/Report_rptEXPIRATIONDATESday.vba` | Original day report VBA |
| `msaccess/extracted/vba/Report_rptEXPIRATIONDATEShouse.vba` | Original house report VBA |
| `msaccess/extracted/vba/Report_rptEXPIRATIONDATESstaff.vba` | Original staff report VBA |

---

**Summary:** vw_ExpirationsFormatted and catExpirationTriggers are correctly designed to replace the VBA formatting logic in the four reports. After applying the view NULL fix and using the updated VBA patterns (bind to _Display/_Color/_ShowDate, ApplyColorFormatting, FormatExpirationField), the view properly replaces the original report VBA except for the few staff-specific rules (TB Pending, MAPCert N/A/120-day, 3MoEval, FullName) which remain in VBA by design.
