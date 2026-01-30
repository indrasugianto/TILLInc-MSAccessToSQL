-- =============================================
-- Stored Procedure: spApp_RunExpirationReport
-- Description: Generates the Expiration Dates report data by aggregating
--              staff, client, and location information into tblExpirations
-- 
-- This replaces the complex SQL operations in the VBA RunExpirationsReport() function
-- 
-- Returns:
--   0 = Success
--   1 = No staff/skills data available (DontRunExpirations condition)
--   -1 = Error occurred
--
-- Version: 1.0
-- Created: 2026-01-30
-- Modified: 2026-01-30
-- 
-- Notes:
--   - Converted from MS Access inline SQL to SQL Server stored procedure
--   - Uses temp tables instead of Access "permanent temp tables"
--   - Direct table references instead of Access query objects
--   - Transaction-wrapped for data consistency
-- =============================================

CREATE OR ALTER PROCEDURE spApp_RunExpirationReport
AS
BEGIN
    SET NOCOUNT ON;
    SET XACT_ABORT ON;
    
    DECLARE @ErrorMessage NVARCHAR(4000);
    DECLARE @ErrorSeverity INT;
    DECLARE @ErrorState INT;
    
    BEGIN TRY
        BEGIN TRANSACTION;
        
        -- ========================================
        -- STEP 1: Validation Check
        -- ========================================
        -- Check if staff and staff skills tables have data
        IF NOT EXISTS (SELECT 1 FROM tblStaff WHERE INACTIVE = 0)
           OR NOT EXISTS (SELECT 1 FROM tblStaffSkills)
        BEGIN
            ROLLBACK TRANSACTION;
            RETURN 1; -- Signal: No data available
        END
        
        -- ========================================
        -- STEP 2: Create Temporary Staff Table
        -- ========================================
        IF OBJECT_ID('tempdb..#tempstaff') IS NOT NULL DROP TABLE #tempstaff;
        
        SELECT 
            EMPLOYID, EMPLCLAS, INACTIVE, LASTNAME, FRSTNAME, MIDLNAME,
            DIVISIONCODE_I, DEPRTMNT, JOBTITLE, SUPERVISORCODE_I, 
            STRTDATE, BENADJDATE
        INTO #tempstaff
        FROM tblStaff
        ORDER BY LASTNAME, FRSTNAME;
        
        -- Update Dedham Managers
        UPDATE ts
        SET 
            ts.DIVISIONCODE_I = 'DEDHAM',
            ts.DEPRTMNT = dm.NewLocation
        FROM #tempstaff ts
        INNER JOIN tblStaffDedhamManagers dm 
            ON ts.SUPERVISORCODE_I = dm.SUPERVISORCODE_I;
        
        -- Add primary key constraint
        ALTER TABLE #tempstaff ADD CONSTRAINT PK_tempstaff PRIMARY KEY (EMPLOYID);
        
        -- Delete example records
        DELETE FROM #tempstaff WHERE LASTNAME = 'EXAMPLE';
        
        -- ========================================
        -- STEP 3: Create Temporary GP Supervisors Table
        -- ========================================
        IF OBJECT_ID('tempdb..#TempSuperCodes') IS NOT NULL DROP TABLE #TempSuperCodes;
        
        CREATE TABLE #TempSuperCodes (
            GPCode VARCHAR(10),
            GPSuperCode VARCHAR(10),
            JobTitle VARCHAR(10)
        );
        
        -- Insert GP Supervisor codes
        INSERT INTO #TempSuperCodes (GPCode, GPSuperCode, JobTitle)
        SELECT 
            DEPRTMNT, 
            SUPERVISORCODE_I, 
            JOBTITLE
        FROM tblStaff
        WHERE 
            (JobTitle IN ('RESUNT', 'RESUPR', 'ASDRRE', 'DASUPR', 'SENDPM'))
            OR (DEPRTMNT = 'CHELSE' AND JobTitle = 'PRGMGR')
            OR (DEPRTMNT = 'NEWTON' AND JobTitle = 'PRGMGR')
            OR (JobTitle IN ('RESMGR', 'SITECO'))
        ORDER BY DEPRTMNT;
        
        -- ========================================
        -- STEP 4: Create Temporary Staff Skills Table
        -- ========================================
        IF OBJECT_ID('tempdb..#tempstaffskills') IS NOT NULL DROP TABLE #tempstaffskills;
        
        SELECT tss.*
        INTO #tempstaffskills
        FROM tblStaff ts
        INNER JOIN tblStaffSkills tss ON ts.EMPLOYID = tss.EMPID_I
        WHERE tss.SKILLNUMBER_I IN (1, 2, 3, 15, 22, 30, 31, 32, 33, 34, 35, 36, 39);
        
        ALTER TABLE #tempstaffskills ADD CONSTRAINT PK_tempstaffskills 
            PRIMARY KEY (EMPID_I, SKILLNUMBER_I);
        
        -- ========================================
        -- STEP 5: Empty the Expirations Table
        -- ========================================
        DELETE FROM tblExpirations;
        
        -- ========================================
        -- STEP 6: Build Program Lookup Table
        -- ========================================
        IF OBJECT_ID('tempdb..#temptbl') IS NOT NULL DROP TABLE #temptbl;
        
        SELECT 
            loc.CityTown + ' - ' + loc.LocationName AS Location,
            loc.CityTown,
            loc.LocationName,
            loc.GPName,
            p.GPSuperCode
        INTO #temptbl
        FROM tblLocations loc
        INNER JOIN tblPeople p 
            ON loc.LocationName = p.OfficeLocationName 
            AND loc.CityTown = p.OfficeCityTown
        WHERE loc.GPName IS NOT NULL 
            AND p.IsStaff = 1
        ORDER BY loc.CityTown + ' - ' + loc.LocationName;
        
        -- Fix blank supercodes using lookup
        UPDATE t
        SET t.GPSuperCode = sc.GPSuperCode
        FROM #temptbl t
        INNER JOIN #TempSuperCodes sc ON t.GPName = sc.GPCode
        WHERE t.GPSuperCode IS NULL;
        
        -- Insert additional ISO locations
        INSERT INTO #temptbl (Location, CityTown, LocationName, GPName, GPSuperCode)
        SELECT 
            loc.CityTown + ' - ' + loc.LocationName AS Location,
            loc.CityTown,
            loc.LocationName,
            loc.GPName,
            (SELECT TOP 1 p.GPSuperCode 
             FROM tblPeople p 
             WHERE p.FirstName = loc.StaffPrimaryContactFirstName 
                AND p.LastName = loc.StaffPrimaryContactLastName) AS GPSuperCode
        FROM tblLocations loc
        WHERE loc.CityTown <> 'Dedham'
            AND loc.GPName IS NOT NULL
            AND loc.Department = 'Individualized Support Options'
        ORDER BY loc.CityTown + ' - ' + loc.LocationName;
        
        -- ========================================
        -- STEP 7: Build Client Services Lookup (temptbl0)
        -- ========================================
        IF OBJECT_ID('tempdb..#temptbl0') IS NOT NULL DROP TABLE #temptbl0;
        
        SELECT 
            ISNULL(res.CityTown + ' - ' + res.Location, '') AS LocRes,
            ISNULL(clo.CityTown + ' - ' + clo.Location, '') AS LocCLO,
            ISNULL(day.CityTown + ' - ' + day.LocationName, '') AS LocDay,
            ISNULL(voc.CityTown + ' - ' + voc.Location, '') AS LocVoc,
            NULL AS Supervisor,
            p.IndexedName,
            p.LastName,
            p.FirstName,
            p.MiddleInitial,
            dem.DateISP,
            dem.DateConsentFormsSigned,
            dem.DateBMMExpires,
            dem.DateBMMAccessSignedHRC,
            dem.DateBMMAccessSigned,
            dem.DateSPDAuthExpires,
            dem.DateSignaturesDueBy,
            dem.AllSPDSignaturesReceived
        INTO #temptbl0
        FROM tblPeople p
        RIGHT JOIN tblPeopleClientsDemographics dem ON p.IndexedName = dem.IndexedName
        LEFT JOIN tblPeopleClientsCLOServices clo ON dem.IndexedName = clo.IndexedName
        LEFT JOIN tblPeopleClientsDayServices day ON dem.IndexedName = day.IndexedName
        LEFT JOIN tblPeopleClientsResidentialServices res ON dem.IndexedName = res.IndexedName
        LEFT JOIN tblPeopleClientsVocationalServices voc ON dem.IndexedName = voc.IndexedName
        WHERE 
            ((dem.ActiveDayServices = 1 AND day.Inactive = 0)
            OR (dem.ActiveResidentialServices = 1 AND res.Inactive = 0)
            OR (dem.ActiveCLO = 1 AND clo.Inactive = 0)
            OR (dem.ActiveVocationalServices = 1 AND voc.Inactive = 0))
            AND p.IsDeceased = 0;
        
        -- Clean up inactive service locations
        -- CLO
        UPDATE t0
        SET t0.LocCLO = ''
        FROM #temptbl0 t0
        LEFT JOIN tblPeopleClientsDemographics qdem ON t0.IndexedName = qdem.IndexedName
        LEFT JOIN tblPeopleClientsCLOServices clo ON t0.IndexedName = clo.IndexedName
        WHERE (LEN(t0.LocCLO) > 0 AND qdem.ActiveCLO = 0)
            OR (LEN(t0.LocCLO) > 0 AND qdem.ActiveCLO = 1 AND clo.Inactive = 1);
        
        -- Residential
        UPDATE t0
        SET t0.LocRes = ''
        FROM #temptbl0 t0
        LEFT JOIN tblPeopleClientsDemographics qdem ON t0.IndexedName = qdem.IndexedName
        LEFT JOIN tblPeopleClientsResidentialServices res ON t0.IndexedName = res.IndexedName
        WHERE (LEN(t0.LocRes) > 0 AND qdem.ActiveResidentialServices = 0)
            OR (LEN(t0.LocRes) > 0 AND qdem.ActiveResidentialServices = 1 AND res.Inactive = 1);
        
        -- Day
        UPDATE t0
        SET t0.LocDay = ''
        FROM #temptbl0 t0
        LEFT JOIN tblPeopleClientsDemographics qdem ON t0.IndexedName = qdem.IndexedName
        LEFT JOIN tblPeopleClientsDayServices day ON t0.IndexedName = day.IndexedName
        WHERE (LEN(t0.LocDay) > 0 AND qdem.ActiveDayServices = 0)
            OR (LEN(t0.LocDay) > 0 AND qdem.ActiveDayServices = 1 AND day.Inactive = 1);
        
        -- Vocational
        UPDATE t0
        SET t0.LocVoc = ''
        FROM #temptbl0 t0
        LEFT JOIN tblPeopleClientsDemographics qdem ON t0.IndexedName = qdem.IndexedName
        LEFT JOIN tblPeopleClientsVocationalServices voc ON t0.IndexedName = voc.IndexedName
        WHERE (LEN(t0.LocVoc) > 0 AND qdem.ActiveVocationalServices = 0)
            OR (LEN(t0.LocVoc) > 0 AND qdem.ActiveVocationalServices = 1 AND voc.Inactive = 1);
        
        -- ========================================
        -- STEP 8: Populate House Information
        -- ========================================
        INSERT INTO tblExpirations (
            Location, RecordType, LastName, FirstName, Supervisor, 
            LastVehicleChecklistCompleted, MostRecentAsleepFireDrill, 
            NextRecentAsleepFireDrill, HouseSafetyPlanExpires, 
            HousePlansReviewedByStaffBefore, DAYStaffTrainedInPrivacyBefore, 
            DAYAllPlansReviewedByStaffBefore, DAYQtrlySafetyChecklistDueBy, 
            MAPChecklistCompleted, HumanRightsOfficer, HROTrainsStaffBefore, 
            HROTrainsIndividualsBefore, FireSafetyOfficer, FSOTrainsStaffBefore, 
            FSOTrainsIndividualsBefore
        )
        SELECT 
            loc.GPName AS Location,
            'House' AS RecordType,
            '*' AS LastName,
            '*' AS FirstName,
            (SELECT TOP 1 t.GPSuperCode FROM #temptbl t WHERE t.GPName = loc.GPName) AS Supervisor,
            loc.LastVehicleChecklistCompleted,
            loc.MostRecentAsleepFireDrill,
            loc.NextRecentAsleepFireDrill,
            loc.HouseSafetyPlanExpires,
            loc.HousePlansReviewedByStaffBefore,
            loc.DAYStaffTrainedInPrivacyBefore,
            loc.DAYAllPlansReviewedByStaffBefore,
            loc.DAYQtrlySafetyChecklistDueBy,
            loc.MAPChecklistCompleted,
            loc.HumanRightsOfficer,
            loc.HROTrainsStaffBefore,
            loc.HROTrainsIndividualsBefore,
            loc.FireSafetyOfficer,
            loc.FSOTrainsStaffBefore,
            loc.FSOTrainsIndividualsBefore
        FROM tblLocations loc
        WHERE loc.GPName IS NOT NULL
            AND EXISTS (SELECT 1 FROM #temptbl t WHERE t.GPName = loc.GPName)
            AND loc.Department <> 'Clinical and Support Services'
        ORDER BY loc.GPName;
        
        -- ========================================
        -- STEP 9: Populate Client Information (CLO, Residential, Vocational)
        -- ========================================
        -- CLO Clients
        INSERT INTO tblExpirations (
            Location, RecordType, LastName, FirstName, Supervisor, 
            DateISP, DateConsentFormsSigned, DateBMMExpires, 
            DateBMMAccessSignedHRC, DateBMMAccessSigned, DateSPDAuthExpires, 
            DateSignaturesDueBy, AllSPDSignaturesReceived
        )
        SELECT 
            (SELECT TOP 1 t.GPName FROM #temptbl t WHERE t.Location = t0.LocCLO) AS Location,
            'Client' AS RecordType,
            t0.LastName,
            t0.FirstName,
            (SELECT TOP 1 t.GPSuperCode FROM #temptbl t WHERE t.Location = t0.LocCLO) AS Supervisor,
            t0.DateISP,
            t0.DateConsentFormsSigned,
            t0.DateBMMExpires,
            t0.DateBMMAccessSignedHRC,
            t0.DateBMMAccessSigned,
            t0.DateSPDAuthExpires,
            t0.DateSignaturesDueBy,
            t0.AllSPDSignaturesReceived
        FROM #temptbl0 t0
        WHERE EXISTS (SELECT 1 FROM #temptbl t WHERE t.Location = t0.LocCLO)
            AND t0.LastName IS NOT NULL
            AND t0.FirstName IS NOT NULL;
        
        -- Residential Clients
        INSERT INTO tblExpirations (
            Location, RecordType, LastName, FirstName, Supervisor, 
            DateISP, DateConsentFormsSigned, DateBMMExpires, 
            DateBMMAccessSignedHRC, DateBMMAccessSigned, DateSPDAuthExpires, 
            DateSignaturesDueBy, AllSPDSignaturesReceived
        )
        SELECT 
            (SELECT TOP 1 t.GPName FROM #temptbl t WHERE t.Location = t0.LocRes) AS Location,
            'Client' AS RecordType,
            t0.LastName,
            t0.FirstName,
            (SELECT TOP 1 t.GPSuperCode FROM #temptbl t WHERE t.Location = t0.LocRes) AS Supervisor,
            t0.DateISP,
            t0.DateConsentFormsSigned,
            t0.DateBMMExpires,
            t0.DateBMMAccessSignedHRC,
            t0.DateBMMAccessSigned,
            t0.DateSPDAuthExpires,
            t0.DateSignaturesDueBy,
            t0.AllSPDSignaturesReceived
        FROM #temptbl0 t0
        WHERE EXISTS (SELECT 1 FROM #temptbl t WHERE t.Location = t0.LocRes)
            AND t0.LastName IS NOT NULL
            AND t0.FirstName IS NOT NULL;
        
        -- Vocational Clients
        INSERT INTO tblExpirations (
            Location, RecordType, LastName, FirstName, Supervisor, 
            DateISP, DateConsentFormsSigned, DateBMMExpires, 
            DateBMMAccessSignedHRC, DateBMMAccessSigned, DateSPDAuthExpires, 
            DateSignaturesDueBy, AllSPDSignaturesReceived
        )
        SELECT 
            (SELECT TOP 1 t.GPName FROM #temptbl t WHERE t.Location = t0.LocVoc) AS Location,
            'Client' AS RecordType,
            t0.LastName,
            t0.FirstName,
            (SELECT TOP 1 t.GPSuperCode FROM #temptbl t WHERE t.Location = t0.LocVoc) AS Supervisor,
            t0.DateISP,
            t0.DateConsentFormsSigned,
            t0.DateBMMExpires,
            t0.DateBMMAccessSignedHRC,
            t0.DateBMMAccessSigned,
            t0.DateSPDAuthExpires,
            t0.DateSignaturesDueBy,
            t0.AllSPDSignaturesReceived
        FROM #temptbl0 t0
        WHERE EXISTS (SELECT 1 FROM #temptbl t WHERE t.Location = t0.LocVoc)
            AND t0.LastName IS NOT NULL
            AND t0.FirstName IS NOT NULL;
        
        -- ========================================
        -- STEP 10: Populate Staff Information
        -- ========================================
        INSERT INTO tblExpirations (
            Location, RecordType, LastName, FirstName, JobTitle, 
            Supervisor, AdjustedStartDate
        )
        SELECT 
            ts.DEPRTMNT AS Location,
            'Staff' AS RecordType,
            ts.LASTNAME,
            ts.FRSTNAME,
            ts.JOBTITLE,
            ts.SUPERVISORCODE_I,
            ts.BENADJDATE AS AdjustedStartDate
        FROM #tempstaff ts
        INNER JOIN #tempstaffskills tss ON ts.EMPLOYID = tss.EMPID_I
        WHERE ts.DEPRTMNT IS NOT NULL
            AND ts.LASTNAME IS NOT NULL
            AND ts.FRSTNAME IS NOT NULL
        ORDER BY ts.LASTNAME, ts.FRSTNAME;
        
        -- ========================================
        -- STEP 11: Update Staff Skills Information
        -- ========================================
        IF OBJECT_ID('tempdb..#temptbl1') IS NOT NULL DROP TABLE #temptbl1;
        IF OBJECT_ID('tempdb..#temptbl2') IS NOT NULL DROP TABLE #temptbl2;
        IF OBJECT_ID('tempdb..#temptbl3') IS NOT NULL DROP TABLE #temptbl3;
        
        -- Get staff skills
        SELECT 
            tss.EMPID_I,
            tss.SKILLNUMBER_I,
            tss.EXPIREDSKILL_I
        INTO #temptbl1
        FROM #tempstaffskills tss
        INNER JOIN tblStaff ts ON tss.EMPID_I = ts.EMPLOYID
        WHERE tss.SKILLNUMBER_I IN (1, 2, 3, 15, 22, 30, 31, 32, 33, 34, 35, 36, 39);
        
        -- Add skill descriptions
        SELECT 
            t1.*,
            (SELECT TOP 1 cs.Skill FROM catSkills cs WHERE cs.SkillID = t1.SKILLNUMBER_I) AS SkillDesc
        INTO #temptbl2
        FROM #temptbl1 t1;
        
        -- Crosstab skills by employee (using PIVOT or similar logic)
        SELECT 
            ts.EMPLOYID AS EMPID_I,
            ts.LASTNAME,
            ts.FRSTNAME,
            ts.DEPRTMNT,
            MAX(CASE WHEN t2.SKILLNUMBER_I = 1 THEN t2.EXPIREDSKILL_I END) AS CPR,
            MAX(CASE WHEN t2.SKILLNUMBER_I = 2 THEN t2.EXPIREDSKILL_I END) AS FirstAid,
            MAX(CASE WHEN t2.SKILLNUMBER_I = 3 THEN t2.EXPIREDSKILL_I END) AS MAPCert,
            MAX(CASE WHEN t2.SKILLNUMBER_I = 15 THEN t2.EXPIREDSKILL_I END) AS DriversLicense,
            MAX(CASE WHEN t2.SKILLNUMBER_I = 22 THEN t2.EXPIREDSKILL_I END) AS BBP,
            MAX(CASE WHEN t2.SKILLNUMBER_I = 30 THEN t2.EXPIREDSKILL_I END) AS SafetyCares,
            MAX(CASE WHEN t2.SKILLNUMBER_I = 31 THEN t2.EXPIREDSKILL_I END) AS WorkplaceViolence,
            MAX(CASE WHEN t2.SKILLNUMBER_I = 32 THEN t2.EXPIREDSKILL_I END) AS DefensiveDriving,
            MAX(CASE WHEN t2.SKILLNUMBER_I = 33 THEN t2.EXPIREDSKILL_I END) AS WheelchairSafety,
            MAX(CASE WHEN t2.SKILLNUMBER_I = 34 THEN t2.EXPIREDSKILL_I END) AS PBS,
            MAX(CASE WHEN t2.SKILLNUMBER_I = 35 THEN t2.EXPIREDSKILL_I END) AS ProfLic,
            MAX(CASE WHEN t2.SKILLNUMBER_I = 36 THEN t2.EXPIREDSKILL_I END) AS TB,
            MAX(CASE WHEN t2.SKILLNUMBER_I = 39 THEN t2.EXPIREDSKILL_I END) AS BackInjuryPrevention
        INTO #temptbl3
        FROM tblStaff ts
        LEFT JOIN #temptbl2 t2 ON ts.EMPLOYID = t2.EMPID_I
        GROUP BY ts.EMPLOYID, ts.LASTNAME, ts.FRSTNAME, ts.DEPRTMNT;
        
        -- Update tblExpirations with staff skills
        UPDATE te
        SET 
            te.CPR = t3.CPR,
            te.FirstAid = t3.FirstAid,
            te.MAPCert = t3.MAPCert,
            te.DriversLicense = t3.DriversLicense,
            te.BBP = t3.BBP,
            te.BackInjuryPrevention = t3.BackInjuryPrevention,
            te.SafetyCares = t3.SafetyCares,
            te.TB = t3.TB,
            te.WorkplaceViolence = t3.WorkplaceViolence,
            te.DefensiveDriving = t3.DefensiveDriving,
            te.WheelchairSafety = t3.WheelchairSafety,
            te.PBS = t3.PBS,
            te.ProfessionalLicenses = t3.ProfLic
        FROM tblExpirations te
        INNER JOIN #temptbl3 t3 
            ON te.FirstName = t3.FRSTNAME 
            AND te.LastName = t3.LASTNAME 
            AND te.Location = t3.DEPRTMNT
        WHERE te.RecordType = 'Staff';
        
        -- ========================================
        -- STEP 12: Update Staff Evaluations and Supervisions
        -- ========================================
        UPDATE te
        SET 
            te.ThreeMonthEvaluation = ev.ThreeMonthEval,
            te.EvalDueBy = ev.EvalDueBy,
            te.LastSupervision = ev.LastSupervision,
            te.OnLeave = ev.OnLeave
        FROM tblExpirations te
        INNER JOIN tblStaffEvalsAndSupervisions ev 
            ON te.FirstName = ev.FirstName 
            AND te.LastName = ev.LastName
        WHERE te.RecordType = 'Staff';
        
        -- ========================================
        -- STEP 13: Add Supervisor's Expirations to Their Own Report
        -- ========================================
        INSERT INTO tblExpirations (
            Location, RecordType, LastName, FirstName, Supervisor, JobTitle, 
            AdjustedStartDate, LastVehicleChecklistCompleted, 
            MostRecentAsleepFireDrill, NextRecentAsleepFireDrill, 
            HouseSafetyPlanExpires, HousePlansReviewedByStaffBefore, 
            DAYStaffTrainedInPrivacyBefore, DAYAllPlansReviewedByStaffBefore, 
            DAYQtrlySafetyChecklistDueBy, MAPChecklistCompleted, 
            HumanRightsOfficer, HROTrainsStaffBefore, HROTrainsIndividualsBefore, 
            FireSafetyOfficer, FSOTrainsStaffBefore, FSOTrainsIndividualsBefore, 
            DateISP, DateConsentFormsSigned, DateBMMExpires, DateBMMAccessSignedHRC, 
            DateBMMAccessSigned, DateSPDAuthExpires, DateSignaturesDueBy, 
            AllSPDSignaturesReceived, BBP, BackInjuryPrevention, CPR, 
            DefensiveDriving, DriversLicense, FirstAid, MAPCert, PBS, 
            SafetyCares, TB, WheelchairSafety, WorkplaceViolence, 
            ProfessionalLicenses, ThreeMonthEvaluation, EvalDueBy, 
            LastSupervision, OnLeave
        )
        SELECT 
            loc.GPName,
            'Staff' AS RecordType,
            te.LastName,
            te.FirstName,
            p.GPSuperCode AS Supervisor,
            te.JobTitle,
            te.AdjustedStartDate,
            te.LastVehicleChecklistCompleted,
            te.MostRecentAsleepFireDrill,
            te.NextRecentAsleepFireDrill,
            te.HouseSafetyPlanExpires,
            te.HousePlansReviewedByStaffBefore,
            te.DAYStaffTrainedInPrivacyBefore,
            te.DAYAllPlansReviewedByStaffBefore,
            te.DAYQtrlySafetyChecklistDueBy,
            te.MAPChecklistCompleted,
            te.HumanRightsOfficer,
            te.HROTrainsStaffBefore,
            te.HROTrainsIndividualsBefore,
            te.FireSafetyOfficer,
            te.FSOTrainsStaffBefore,
            te.FSOTrainsIndividualsBefore,
            te.DateISP,
            te.DateConsentFormsSigned,
            te.DateBMMExpires,
            te.DateBMMAccessSignedHRC,
            te.DateBMMAccessSigned,
            te.DateSPDAuthExpires,
            te.DateSignaturesDueBy,
            te.AllSPDSignaturesReceived,
            te.BBP,
            te.BackInjuryPrevention,
            te.CPR,
            te.DefensiveDriving,
            te.DriversLicense,
            te.FirstAid,
            te.MAPCert,
            te.PBS,
            te.SafetyCares,
            te.TB,
            te.WheelchairSafety,
            te.WorkplaceViolence,
            te.ProfessionalLicenses,
            te.ThreeMonthEvaluation,
            te.EvalDueBy,
            te.LastSupervision,
            te.OnLeave
        FROM tblLocations loc
        INNER JOIN tblPeople p ON loc.StaffPrimaryContactIndexedName = p.IndexedName
        INNER JOIN tblExpirations te ON p.FirstName = te.FirstName AND p.LastName = te.LastName
        WHERE loc.GPName IS NOT NULL
            AND p.GPSuperCode IS NOT NULL
            AND loc.Department IN ('Residential Services', 'Day Services', 'Vocational Services', 'TILL NH', 'Expirations Reporting')
            AND RIGHT(loc.StaffPrimaryContactIndexedName, 5) <> 'TBD//'
        
        UNION
        
        SELECT 
            loc.GPName,
            'Staff' AS RecordType,
            te.LastName,
            te.FirstName,
            p.GPSuperCode AS Supervisor,
            te.JobTitle,
            te.AdjustedStartDate,
            te.LastVehicleChecklistCompleted,
            te.MostRecentAsleepFireDrill,
            te.NextRecentAsleepFireDrill,
            te.HouseSafetyPlanExpires,
            te.HousePlansReviewedByStaffBefore,
            te.DAYStaffTrainedInPrivacyBefore,
            te.DAYAllPlansReviewedByStaffBefore,
            te.DAYQtrlySafetyChecklistDueBy,
            te.MAPChecklistCompleted,
            te.HumanRightsOfficer,
            te.HROTrainsStaffBefore,
            te.HROTrainsIndividualsBefore,
            te.FireSafetyOfficer,
            te.FSOTrainsStaffBefore,
            te.FSOTrainsIndividualsBefore,
            te.DateISP,
            te.DateConsentFormsSigned,
            te.DateBMMExpires,
            te.DateBMMAccessSignedHRC,
            te.DateBMMAccessSigned,
            te.DateSPDAuthExpires,
            te.DateSignaturesDueBy,
            te.AllSPDSignaturesReceived,
            te.BBP,
            te.BackInjuryPrevention,
            te.CPR,
            te.DefensiveDriving,
            te.DriversLicense,
            te.FirstAid,
            te.MAPCert,
            te.PBS,
            te.SafetyCares,
            te.TB,
            te.WheelchairSafety,
            te.WorkplaceViolence,
            te.ProfessionalLicenses,
            te.ThreeMonthEvaluation,
            te.EvalDueBy,
            te.LastSupervision,
            te.OnLeave
        FROM tblLocations loc
        INNER JOIN tblPeople p ON loc.StaffPrimaryContactIndexedName = p.IndexedName
        INNER JOIN tblExpirations te ON p.FirstName = te.FirstName AND p.LastName = te.LastName
        WHERE loc.GPName IS NOT NULL
            AND p.GPSuperCode IS NOT NULL
            AND loc.CityTown <> 'Dedham'
            AND loc.Department = 'Individualized Support Options'
            AND RIGHT(loc.StaffPrimaryContactIndexedName, 5) <> 'TBD//';
        
        -- ========================================
        -- Cleanup: Drop temporary tables
        -- ========================================
        IF OBJECT_ID('tempdb..#tempstaff') IS NOT NULL DROP TABLE #tempstaff;
        IF OBJECT_ID('tempdb..#tempstaffskills') IS NOT NULL DROP TABLE #tempstaffskills;
        IF OBJECT_ID('tempdb..#TempSuperCodes') IS NOT NULL DROP TABLE #TempSuperCodes;
        IF OBJECT_ID('tempdb..#temptbl') IS NOT NULL DROP TABLE #temptbl;
        IF OBJECT_ID('tempdb..#temptbl0') IS NOT NULL DROP TABLE #temptbl0;
        IF OBJECT_ID('tempdb..#temptbl1') IS NOT NULL DROP TABLE #temptbl1;
        IF OBJECT_ID('tempdb..#temptbl2') IS NOT NULL DROP TABLE #temptbl2;
        IF OBJECT_ID('tempdb..#temptbl3') IS NOT NULL DROP TABLE #temptbl3;
        
        COMMIT TRANSACTION;
        RETURN 0; -- Success
        
    END TRY
    BEGIN CATCH
        IF @@TRANCOUNT > 0
            ROLLBACK TRANSACTION;
        
        -- Capture error information
        SELECT 
            @ErrorMessage = ERROR_MESSAGE(),
            @ErrorSeverity = ERROR_SEVERITY(),
            @ErrorState = ERROR_STATE();
        
        -- Log the error (optional - if you have an error log table)
        -- INSERT INTO ErrorLog (ErrorMessage, ErrorSeverity, ErrorState, ErrorDate)
        -- VALUES (@ErrorMessage, @ErrorSeverity, @ErrorState, GETDATE());
        
        -- Re-throw the error
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
        
        RETURN -1; -- Error
    END CATCH
END
GO
