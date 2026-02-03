# RunExpirationsReport: VBA vs spApp_RunExpirationReport Assessment

This document maps every SQL step in `RunExpirationsReport()` (Expirations_Original.vba) to the stored procedure `spApp_RunExpirationReport.sql` and records any discrepancies and fixes.

---

## Pre-report (outside SP)

| Step | VBA | Notes |
|------|-----|--------|
| **rptSTAFFWITHNOSKILLS** | `Call ExecReport("rptSTAFFWITHNOSKILLS")` | Runs in VBA before calling SP. Not part of SP. |
| **DontRunExpirations** | Checked in VBA; `DCount("Inactive", "tblStaff") = 0 Or DCount("EMPID_I", "tblStaffSkills") = 0` | SP has its own validation (STEP 1). Aligned: SP now uses `NOT EXISTS (SELECT 1 FROM tblStaff) OR NOT EXISTS (SELECT 1 FROM tblStaffSkills)`. |

---

## Step-by-step mapping

### STEP 1: Validation (no data = don’t run)

| VBA | SP | Status |
|-----|----|--------|
| Handled in VBA via `DontRunExpirations`. Form sets it when `DCount("Inactive", "tblStaff") = 0 Or DCount("EMPID_I", "tblStaffSkills") = 0`. | STEP 1: Return 1 if no staff or no staff skills. | **Fixed.** SP now uses `NOT EXISTS (SELECT 1 FROM tblStaff) OR NOT EXISTS (SELECT 1 FROM tblStaffSkills)` to match “no records in Staff or Staff Skills tables.” |

---

### STEP 2: Temporary staff table (qryEXPIRATIONS01–03, 02A)

| Query / Logic | VBA | SP | Status |
|---------------|-----|----|--------|
| **qryEXPIRATIONS01** | `SELECT tblStaff.* INTO tempstaff FROM tblStaff ORDER BY LASTNAME, FRSTNAME` | `SELECT` specific columns `INTO #tempstaff FROM tblStaff ORDER BY LASTNAME, FRSTNAME` | OK. SP only needs columns used later; same row set. |
| **qryEXPIRATIONS02** | `UPDATE qrytblStaffDedhamManagers INNER JOIN tempstaff ON ... SET tempstaff.DIVISIONCODE_I = 'DEDHAM', tempstaff.DEPRTMNT = qrytblStaffDedhamManagers.NewLocation` | `UPDATE #tempstaff` from `tblStaffDedhamManagers` (qry is `SELECT * FROM tblStaffDedhamManagers`) SET `DIVISIONCODE_I = 'DEDHAM', DEPRTMNT = dm.NewLocation` | OK. Same logic; SP uses table directly. |
| **qryEXPIRATIONS02A** | `ALTER TABLE tempstaff ADD CONSTRAINT ... PRIMARY KEY (EMPLOYID)` | `ALTER TABLE #tempstaff ADD CONSTRAINT PK_tempstaff PRIMARY KEY (EMPLOYID)` | OK. |
| **qryEXPIRATIONS03** | `DELETE tempstaff.* WHERE LastName = 'EXAMPLE'` | `DELETE FROM #tempstaff WHERE LASTNAME = 'EXAMPLE'` | OK. |

---

### STEP 3: Temporary GP Supervisors (qryEXPIRATIONS03A, 04)

| Query / Logic | VBA | SP | Status |
|---------------|-----|----|--------|
| **qryEXPIRATIONS03A** | `DELETE [~TempSuperCodes].*` | SP creates new `#TempSuperCodes` (no delete of persistent table). | OK. SP uses session temp table. |
| **qryEXPIRATIONS04** | `INSERT INTO [~TempSuperCodes] (GPCode, GPSuperCode, JobTitle) SELECT tblStaff.DEPRTMNT, SUPERVISORCODE_I, JOBTITLE FROM tblStaff WHERE (JobTitle IN (...)) OR (DEPRTMNT='CHELSE' AND JobTitle='PRGMGR') OR ... ORDER BY DEPRTMNT` | Same filter and columns; `INSERT INTO #TempSuperCodes` from `tblStaff`. | OK. |

---

### STEP 4: Temporary staff skills (qryEXPIRATIONS05, 05A)

| Query / Logic | VBA | SP | Status |
|---------------|-----|----|--------|
| **qryEXPIRATIONS05** | `SELECT tblStaffSkills.* INTO tempstaffskills FROM tblStaff LEFT JOIN tblStaffSkills ON ... WHERE SKILLNUMBER_I IN (1,2,3,15,22,30,31,32,33,34,35,36,39)` | `SELECT tss.* INTO #tempstaffskills FROM tblStaff ts INNER JOIN tblStaffSkills tss ON ... WHERE tss.SKILLNUMBER_I IN (1,2,3,15,22,30,31,32,33,34,35,36,39)` | OK. Same skill set; INNER vs LEFT+WHERE gives same rows when filtering on skills. |
| **qryEXPIRATIONS05A** | `ALTER TABLE tempstaffskills ADD CONSTRAINT ... PRIMARY KEY (EMPID_I, SKILLNUMBER_I)` | Same for `#tempstaffskills`. | OK. |

---

### STEP 5: Empty expirations (qryEXPIRATIONS05B)

| VBA | SP | Status |
|-----|----|--------|
| `DELETE tblExpirations.*` | `DELETE FROM tblExpirations` | OK. |

---

### STEP 6: Program lookup temptbl (qryEXPIRATIONS06–08)

| Query / Logic | VBA | SP | Status |
|---------------|-----|----|--------|
| **qryEXPIRATIONS06** | `SELECT [CityTown] & ' - ' & [LocationName] AS Location, ..., tblLocations.GPName, tblPeople.GPSuperCode INTO temptbl FROM tblLocations INNER JOIN tblPeople ON (LocationName = OfficeLocationName) AND (CityTown = OfficeCityTown) WHERE GPName IS NOT NULL AND IsStaff = True` | Same join and filters; `#temptbl` with `Location, CityTown, LocationName, GPName, GPSuperCode`. | OK. |
| **qryEXPIRATIONS07** | `UPDATE temptbl SET GPSuperCode = DLookUp(...'~TempSuperCodes'...) WHERE GPSuperCode IS NULL` | `UPDATE #temptbl` from `#TempSuperCodes` where `GPSuperCode IS NULL`. | OK. |
| **qryEXPIRATIONS08** | `INSERT INTO temptbl ... FROM tblLocations WHERE CityTown <> 'Dedham' AND GPName IS NOT NULL AND Department = 'Individualized Support Options'`; GPSuperCode from `DLookUp('tblPeople', FirstName/LastName)` | `INSERT INTO #temptbl` same filters; GPSuperCode via subquery on `tblPeople` by `StaffPrimaryContactFirstName/LastName`. | OK. |

---

### STEP 7: Client lookup temptbl0 (qryEXPIRATIONS09–13)

| Query / Logic | VBA | SP | Status |
|---------------|-----|----|--------|
| **qryEXPIRATIONS09** | Build temptbl0 from tblPeople, tblPeopleClientsDemographics, CLO/Day/Res/Voc services; LocRes, LocCLO, LocDay, LocVoc via IIf(IsNull(...),'', ...); WHERE (Active* + Inactive) and IsDeceased = False | Same joins and conditions; Loc* as `ISNULL(... , '')`. | OK. |
| **qryEXPIRATIONS10–13** | UPDATE temptbl0 set LocCLO/LocRes/LocDay/LocVoc = Null where inactive or not active for that program. Uses `qrytblPeopleClientsDemographics` (same as table). | SP UPDATEs set Loc* = `''` using `tblPeopleClientsDemographics` and service tables. | OK. Using `''` vs Null is equivalent for “clear location” and subsequent EXISTS checks. |

---

### STEP 8: Populate house info (qryEXPIRATIONS14)

| Query / Logic | VBA | SP | Status |
|---------------|-----|----|--------|
| **qryEXPIRATIONS14** | INSERT house rows from tblLocations where `GPName IS NOT NULL AND DLookUp('GPSuperCode','temptbl','GPName=...') IS NOT NULL AND Department <> 'Clinical and Support Services'` | Same INSERT list and filters. | **Fixed.** SP now requires `(SELECT TOP 1 t.GPSuperCode FROM #temptbl t WHERE t.GPName = loc.GPName) IS NOT NULL` so only locations with a non-null GPSuperCode in the lookup are inserted, matching VBA. |

---

### STEP 9: Populate client info – CLO, Residential, Vocational (qryEXPIRATIONS15–17)

| Query / Logic | VBA | SP | Status |
|---------------|-----|----|--------|
| **qryEXPIRATIONS15** | INSERT from temptbl0 for CLO: Location/Supervisor via DLookUp on temptbl by LocCLO; WHERE DLookUp GPName not null and LastName/FirstName not null. | INSERT from #temptbl0; Location/Supervisor via `SELECT TOP 1 ... FROM #temptbl WHERE Location = t0.LocCLO`; WHERE EXISTS on temptbl and LastName/FirstName not null. | OK. |
| **qryEXPIRATIONS16** | Same pattern for LocRes. | Same for LocRes. | OK. |
| **qryEXPIRATIONS17** | Same pattern for LocVoc. | Same for LocVoc. | OK. |
| **Day clients** | No INSERT for “Day” clients (only CLO, Res, Voc). | SP also has no Day client INSERT. | OK. |

---

### STEP 10: Populate staff (qryEXPIRATIONS18)

| VBA | SP | Status |
|-----|----|--------|
| INSERT into tblExpirations (Location, RecordType, LastName, FirstName, JobTitle, Supervisor, AdjustedStartDate) from tempstaff INNER JOIN tempstaffskills, WHERE DEPRTMNT/LastName/FRSTNAME not null, ORDER BY LASTNAME, FRSTNAME | Same columns and logic from #tempstaff and #tempstaffskills. | OK. |

---

### STEP 11: Staff skills crosstab and update (qryEXPIRATIONS19–22)

| Query / Logic | VBA | SP | Status |
|---------------|-----|----|--------|
| **qryEXPIRATIONS19** | temptbl1 = tempstaffskills JOIN tblStaff, skill numbers 1,2,3,15,22,30,31,32,33,34,35,36,39 | #temptbl1 from #tempstaffskills JOIN tblStaff, same skill list. | OK. |
| **qryEXPIRATIONS20** | temptbl2 = temptbl1 + DLookUp("Skill","catSkills","SkillID=" & SKILLNUMBER_I) AS SkillDesc | #temptbl2 = #temptbl1 + subquery on catSkills.Skill. | OK. |
| **qryEXPIRATIONS21** | temptbl3 = qryExpirationsStaffBySkills (TRANSFORM/PIVOT on SkillDesc) | #temptbl3 = manual crosstab with CASE WHEN SKILLNUMBER_I = 1 THEN ... (CPR, FirstAid, MAPCert, DriversLicense, BBP, SafetyCares, etc.) | OK. Same skill IDs → same column names. |
| **qryEXPIRATIONS22** | UPDATE qrytblExpirations ( = tblExpirations) SET CPR, FirstAid, MAPCert, DriversLicense, BBP, BackInjuryPrevention, SafetyCares, TB, WorkplaceViolence, DefensiveDriving, WheelchairSafety, PBS, ProfLic FROM temptbl3 WHERE RecordType='Staff' | UPDATE tblExpirations from #temptbl3 same columns, WHERE RecordType = 'Staff'. | OK. |

---

### STEP 12: Staff evals and supervisions (qryEXPIRATIONS23)

| VBA | SP | Status |
|-----|----|--------|
| UPDATE qrytblExpirations JOIN qrytblStaffEvalsAndSupervisions ON FirstName, LastName SET ThreeMonthEvaluation, EvalDueBy, LastSupervision, OnLeave WHERE RecordType='Staff'. Query is SELECT * FROM tblStaffEvalsAndSupervisions. | UPDATE tblExpirations from tblStaffEvalsAndSupervisions on FirstName, LastName, same columns, WHERE RecordType = 'Staff'. | OK. |

---

### STEP 13: Supervisor expirations (qryEXPIRATIONS24)

| VBA | SP | Status |
|-----|----|--------|
| INSERT into tblExpirations (full column list) from tblLocations JOIN tblPeople ON StaffPrimaryContactIndexedName JOIN tblExpirations on FirstName/LastName. WHERE (GPName/ GPSuperCode not null and Department in list and Right(StaffPrimaryContactIndexedName,5) <> 'TBD//') OR (same plus CityTown <> 'Dedham' and Department = 'Individualized Support Options'). ORDER BY CityTown. | Same INSERT and column list. Two branches implemented as UNION: first branch Department IN ('Residential Services', 'Day Services', ...); second branch CityTown <> 'Dedham' AND Department = 'Individualized Support Options'. Both exclude Right(StaffPrimaryContactIndexedName, 5) = 'TBD//'. | OK. |

---

### Cleanup (qryEXPIRATIONS25–26)

| VBA | SP | Status |
|-----|----|--------|
| qryEXPIRATIONS25: DELETE tblExpirations (done at end in VBA after report). qryEXPIRATIONS26: DELETE [~TempSuperCodes]. | SP does not clear tblExpirations (caller/Export step may do that). Temp tables #tempstaff, #TempSuperCodes, etc. dropped at end of SP. | OK. VBA cleanup is outside the “build data” scope; SP correctly drops its own temp tables. |

---

## Summary of changes made to the SP

1. **STEP 1 – Validation**  
   - **Before:** `NOT EXISTS (SELECT 1 FROM tblStaff WHERE INACTIVE = 0) OR NOT EXISTS (SELECT 1 FROM tblStaffSkills)`  
   - **After:** `NOT EXISTS (SELECT 1 FROM tblStaff) OR NOT EXISTS (SELECT 1 FROM tblStaffSkills)`  
   - Aligns with VBA “no records in the Staff or Staff Skills tables” (DCount on tblStaff/tblStaffSkills = 0).

2. **STEP 8 – House insert**  
   - **Before:** `AND EXISTS (SELECT 1 FROM #temptbl t WHERE t.GPName = loc.GPName)`  
   - **After:** `AND (SELECT TOP 1 t.GPSuperCode FROM #temptbl t WHERE t.GPName = loc.GPName) IS NOT NULL`  
   - Matches VBA: only insert houses when the program lookup has a non-null GPSuperCode for that GPName.

---

## Queries not in the SP (by design)

- **rptSTAFFWITHNOSKILLS** – Run in VBA before the SP; not part of report data build.
- **qryEXPIRATIONS25/26** – Table cleanup and ~TempSuperCodes cleanup; in VBA after export. SP only drops its own # temp tables.

---

## Conclusion

All SQL steps in `RunExpirationsReport()` that build expiration data are reflected in `spApp_RunExpirationReport` with two corrections applied:

1. Validation (STEP 1) now matches the VBA “no data” condition (empty tblStaff or tblStaffSkills).  
2. House insert (STEP 8) now only includes locations whose GPName has a non-null GPSuperCode in the program lookup, matching the VBA DLookUp condition.

No Day client insert exists in VBA; the SP correctly omits it and only inserts CLO, Residential, and Vocational clients.
