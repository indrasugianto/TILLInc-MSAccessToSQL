-- =============================================
-- Table: catExpirationTriggers
-- Description: Stores threshold values (Red/Green) for expiration date warnings
--              Replaces VBA global variables with database configuration
--
-- This table is required by vw_ExpirationsFormatted
--
-- Version: 1.0
-- Created: 2026-01-30
-- =============================================
drop table if exists catExpirationTriggers;

-- Create the table if it doesn't exist
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'catExpirationTriggers') AND type = 'U')
BEGIN
    CREATE TABLE catExpirationTriggers (
        TriggerID INT IDENTITY(1,1) PRIMARY KEY,
        Section VARCHAR(50) NULL,         -- 'Individuals', NULL for Staff
        Program VARCHAR(50) NULL,         -- 'Day', 'Res' (Residential), NULL for Individuals/Staff
        FieldName VARCHAR(100) NOT NULL,  -- Field name this trigger applies to
        Red INT  NULL,                 -- Red threshold (days or months depending on field)
        Green INT NULL,                   -- Green threshold (days or months)
        [Description] VARCHAR(255) NULL,  -- Human-readable description (escaped - SQL keyword)
        LastModified DATETIME DEFAULT GETDATE(),
        CONSTRAINT UQ_TriggerConfig UNIQUE (Section, Program, FieldName)
    );
    
    PRINT 'Table catExpirationTriggers created successfully.';
END
ELSE
BEGIN
    PRINT 'Table catExpirationTriggers already exists.';
END
GO

-- =============================================
-- Populate with default values
-- These are typical values - adjust based on your business rules
-- =============================================

-- Clear existing data (if re-running)
TRUNCATE TABLE catExpirationTriggers;

-- ========================================
-- INDIVIDUALS (Client) Triggers
-- ========================================
INSERT INTO catExpirationTriggers (Section, Program, FieldName, Red, Green, [Description]) VALUES
('Individuals', NULL, 'DateISP', -1, 60, 'ISP expiration date - Red if expired, Green if within 60 days'),
('Individuals', NULL, 'PSDue', NULL, 60, 'Person Served Due date - Green if within 60 days of due date'),
('Individuals', NULL, 'DateConsentFormsSigned', 12, 11, 'Consent Forms - Red if > 12 months, Green if > 11 months (in months)'),
('Individuals', NULL, 'DateBMMExpires', -1, 60, 'BMM Expiration - Red if expired, Green if within 60 days'),
('Individuals', NULL, 'DateBMMAccessSignedHRC', NULL, NULL, 'BMM Access Signed HRC'),
('Individuals', NULL, 'DateBMMAccessSigned', NULL, NULL, 'BMM Access Signed'),
('Individuals', NULL, 'DateSPDAuthExpires', -1, 60, 'SPD Authorization Expires - Red if expired, Green if within 60 days'),
('Individuals', NULL, 'DateSignaturesDueBy', -1, 60, 'Signatures Due By - Red if expired, Green if within 60 days');

-- ========================================
-- DAY PROGRAM Triggers
-- ========================================
INSERT INTO catExpirationTriggers (Section, Program, FieldName, Red, Green, [Description]) VALUES
(NULL, 'Day', 'LastVehicleChecklistCompleted', 365, NULL, 'Last Vehicle Checklist - Red if >= 365 days old'),
(NULL, 'Day', 'DAYStaffTrainedInPrivacyBefore', -1, 60, 'Staff Privacy Training Due - Red if expired, Green if within 60 days'),
(NULL, 'Day', 'DAYAllPlansReviewedByStaffBefore', -1, 60, 'Plans Review Due - Red if expired, Green if within 60 days'),
(NULL, 'Day', 'DAYQtrlySafetyChecklistDueBy', -1, 30, 'Quarterly Safety Checklist - Red if expired, Green if within 30 days'),
(NULL, 'Day', 'HROTrainsStaffBefore', -1, 60, 'HRO Staff Training Due - Red if expired, Green if within 60 days'),
(NULL, 'Day', 'HROTrainsIndividualsBefore', -1, 60, 'HRO Individuals Training Due - Red if expired, Green if within 60 days'),
(NULL, 'Day', 'FSOTrainsStaffBefore', -1, 60, 'Fire Safety Officer Staff Training - Red if expired, Green if within 60 days'),
(NULL, 'Day', 'FSOTrainsIndividualsBefore', -1, 60, 'Fire Safety Officer Individuals Training - Red if expired, Green if within 60 days');

-- ========================================
-- RESIDENTIAL PROGRAM Triggers
-- ========================================
INSERT INTO catExpirationTriggers (Section, Program, FieldName, Red, Green, [Description]) VALUES
(NULL, 'Res', 'LastVehicleChecklistCompleted', 365, NULL, 'Last Vehicle Checklist - Red if >= 365 days old'),
(NULL, 'Res', 'MostRecentAsleepFireDrill', -1, 30, 'Asleep Fire Drill - Red if 14-month expiry passed, Green if within 30 days'),
(NULL, 'Res', 'HousePlansReviewedByStaffBefore', -1, 60, 'House Plans Review Due - Red if expired, Green if within 60 days'),
(NULL, 'Res', 'HouseSafetyPlanExpires', -1, 60, 'House Safety Plan Expires - Red if expired, Green if within 60 days'),
(NULL, 'Res', 'MAPChecklistCompleted', 365, NULL, 'MAP Checklist - Red if >= 365 days old'),
(NULL, 'Res', 'HROTrainsStaffBefore', -1, 60, 'HRO Staff Training Due - Red if expired, Green if within 60 days'),
(NULL, 'Res', 'HROTrainsIndividualsBefore', -1, 60, 'HRO Individuals Training Due - Red if expired, Green if within 60 days'),
(NULL, 'Res', 'FSOTrainsStaffBefore', -1, 60, 'Fire Safety Officer Staff Training - Red if expired, Green if within 60 days'),
(NULL, 'Res', 'FSOTrainsIndividualsBefore', -1, 60, 'Fire Safety Officer Individuals Training - Red if expired, Green if within 60 days');

-- ========================================
-- STAFF Triggers (for future use if needed)
-- ========================================
INSERT INTO catExpirationTriggers (Section, Program, FieldName, Red, Green, [Description]) VALUES
(NULL, NULL, 'CPR', -1, 30, 'CPR Certification - Red if expired, Green if within 30 days'),
(NULL, NULL, 'FirstAid', -1, 30, 'First Aid Certification - Red if expired, Green if within 30 days'),
(NULL, NULL, 'MAPCert', -1, 30, 'MAP Certification - Red if expired, Green if within 30 days'),
(NULL, NULL, 'DriversLicense', -1, 60, 'Drivers License - Red if expired, Green if within 60 days'),
(NULL, NULL, 'BBP', -1, 365, 'Bloodborne Pathogens - Red if expired, Green if within 365 days'),
(NULL, NULL, 'SafetyCares', -1, 365, 'Safety Cares - Red if expired, Green if within 365 days'),
(NULL, NULL, 'WorkplaceViolence', -1, 365, 'Workplace Violence Training - Red if expired, Green if within 365 days'),
(NULL, NULL, 'TB', -1, 365, 'TB Test - Red if expired, Green if within 365 days'),
(NULL, NULL, 'BackInjuryPrevention', -1, 365, 'Back Injury Prevention - Red if expired, Green if within 365 days'),
(NULL, NULL, 'ProfLic', -1, 30, 'Professional License - Red if expired, Green if within 30 days'),
(NULL, NULL, 'EvalDueBy', -1, 30, 'Evaluation Due - Red if expired, Green if within 30 days'),
(NULL, NULL, 'LastSupervision', 90, NULL, 'Last Supervision - Red if > 90 days since last supervision'),
(NULL, NULL, '3MoEval', -1, NULL, '3 Month Evaluation - Red if expired');

GO

-- Display the inserted data
SELECT 
    TriggerID,
    ISNULL(Section, '--') AS Section,
    ISNULL(Program, '--') AS Program,
    FieldName,
    Red,
    ISNULL(CAST(Green AS VARCHAR), '--') AS Green,
    [Description]
FROM catExpirationTriggers
ORDER BY 
    CASE Section WHEN 'Individuals' THEN 1 ELSE 2 END,
    CASE Program WHEN 'Day' THEN 1 WHEN 'Res' THEN 2 ELSE 3 END,
    FieldName;

PRINT '';
PRINT 'catExpirationTriggers table populated with default values.';
PRINT 'NOTE: These are sample/typical values. Review and adjust based on your business rules.';
PRINT '';
PRINT 'To modify a trigger value, use:';
PRINT '  UPDATE catExpirationTriggers SET Red = <value>, Green = <value> WHERE FieldName = ''<field>'' AND Program = ''<program>'';';
