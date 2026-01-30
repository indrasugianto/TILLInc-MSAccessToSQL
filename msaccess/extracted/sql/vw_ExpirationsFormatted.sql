-- =============================================
-- View: vw_ExpirationsFormatted
-- Description: Pre-calculates all formatting logic for the Expirations report
--              Eliminates VBA calculations and network round trips
--
-- This view replaces VBA logic in:
--   - Report_rptEXPIRATIONDATESclients
--   - Report_rptEXPIRATIONDATESday
--   - Report_rptEXPIRATIONDATEShouse
--
-- Version: 1.1
-- Created: 2026-01-30
-- Modified: 2026-01-30 - Added Department, Cluster lookups from tblLocations
-- =============================================

-- First, ensure we have the trigger configuration table
-- (Run the accompanying script: create_catExpirationTriggers.sql)

GO

CREATE OR ALTER VIEW vw_ExpirationsFormatted
AS
WITH TriggerValues AS (
    -- Cache all trigger values to avoid repeated lookups
    SELECT 
        MAX(CASE WHEN Section='Individuals' AND FieldName='DateISP' THEN Red END) AS Trig_Indiv_ISP_Red,
        MAX(CASE WHEN Section='Individuals' AND FieldName='DateISP' THEN Green END) AS Trig_Indiv_ISP_Green,
        MAX(CASE WHEN Section='Individuals' AND FieldName='PSDue' THEN Green END) AS Trig_Indiv_PSDue_Green,
        MAX(CASE WHEN Section='Individuals' AND FieldName='DateConsentFormsSigned' THEN Red END) AS Trig_Indiv_CFS_Red,
        MAX(CASE WHEN Section='Individuals' AND FieldName='DateConsentFormsSigned' THEN Green END) AS Trig_Indiv_CFS_Green,
        MAX(CASE WHEN Section='Individuals' AND FieldName='DateBMMExpires' THEN Red END) AS Trig_Indiv_BMMX_Red,
        MAX(CASE WHEN Section='Individuals' AND FieldName='DateBMMExpires' THEN Green END) AS Trig_Indiv_BMMX_Green,
        MAX(CASE WHEN Section='Individuals' AND FieldName='DateSPDAuthExpires' THEN Red END) AS Trig_Indiv_SPDX_Red,
        MAX(CASE WHEN Section='Individuals' AND FieldName='DateSPDAuthExpires' THEN Green END) AS Trig_Indiv_SPDX_Green,
        MAX(CASE WHEN Section='Individuals' AND FieldName='DateSignaturesDueBy' THEN Red END) AS Trig_Indiv_SPDA_Red,
        MAX(CASE WHEN Section='Individuals' AND FieldName='DateSignaturesDueBy' THEN Green END) AS Trig_Indiv_SPDA_Green,
        
        -- Day program triggers
        MAX(CASE WHEN Program='Day' AND FieldName='LastVehicleChecklistCompleted' THEN Red END) AS Trig_Day_LVC_Red,
        MAX(CASE WHEN Program='Day' AND FieldName='DAYStaffTrainedInPrivacyBefore' THEN Red END) AS Trig_Day_STP_Red,
        MAX(CASE WHEN Program='Day' AND FieldName='DAYStaffTrainedInPrivacyBefore' THEN Green END) AS Trig_Day_STP_Green,
        MAX(CASE WHEN Program='Day' AND FieldName='DAYAllPlansReviewedByStaffBefore' THEN Red END) AS Trig_Day_APRS_Red,
        MAX(CASE WHEN Program='Day' AND FieldName='DAYAllPlansReviewedByStaffBefore' THEN Green END) AS Trig_Day_APRS_Green,
        MAX(CASE WHEN Program='Day' AND FieldName='DAYQtrlySafetyChecklistDueBy' THEN Red END) AS Trig_Day_QSR_Red,
        MAX(CASE WHEN Program='Day' AND FieldName='DAYQtrlySafetyChecklistDueBy' THEN Green END) AS Trig_Day_QSR_Green,
        MAX(CASE WHEN Program='Day' AND FieldName='HROTrainsStaffBefore' THEN Red END) AS Trig_Day_HROTS_Red,
        MAX(CASE WHEN Program='Day' AND FieldName='HROTrainsStaffBefore' THEN Green END) AS Trig_Day_HROTS_Green,
        MAX(CASE WHEN Program='Day' AND FieldName='HROTrainsIndividualsBefore' THEN Red END) AS Trig_Day_HROTI_Red,
        MAX(CASE WHEN Program='Day' AND FieldName='HROTrainsIndividualsBefore' THEN Green END) AS Trig_Day_HROTI_Green,
        MAX(CASE WHEN Program='Day' AND FieldName='FSOTrainsStaffBefore' THEN Red END) AS Trig_Day_FSOTS_Red,
        MAX(CASE WHEN Program='Day' AND FieldName='FSOTrainsStaffBefore' THEN Green END) AS Trig_Day_FSOTS_Green,
        MAX(CASE WHEN Program='Day' AND FieldName='FSOTrainsIndividualsBefore' THEN Red END) AS Trig_Day_FSOTI_Red,
        MAX(CASE WHEN Program='Day' AND FieldName='FSOTrainsIndividualsBefore' THEN Green END) AS Trig_Day_FSOTI_Green,
        
        -- Residential program triggers
        MAX(CASE WHEN Program='Res' AND FieldName='LastVehicleChecklistCompleted' THEN Red END) AS Trig_Res_LVC_Red,
        MAX(CASE WHEN Program='Res' AND FieldName='MostRecentAsleepFireDrill' THEN Red END) AS Trig_Res_MRFD_Red,
        MAX(CASE WHEN Program='Res' AND FieldName='MostRecentAsleepFireDrill' THEN Green END) AS Trig_Res_MRFD_Green,
        MAX(CASE WHEN Program='Res' AND FieldName='HousePlansReviewedByStaffBefore' THEN Red END) AS Trig_Res_HPR_Red,
        MAX(CASE WHEN Program='Res' AND FieldName='HousePlansReviewedByStaffBefore' THEN Green END) AS Trig_Res_HPR_Green,
        MAX(CASE WHEN Program='Res' AND FieldName='HouseSafetyPlanExpires' THEN Red END) AS Trig_Res_HSPE_Red,
        MAX(CASE WHEN Program='Res' AND FieldName='HouseSafetyPlanExpires' THEN Green END) AS Trig_Res_HSPE_Green,
        MAX(CASE WHEN Program='Res' AND FieldName='MAPChecklistCompleted' THEN Red END) AS Trig_Res_MAP_Red,
        MAX(CASE WHEN Program='Res' AND FieldName='HROTrainsStaffBefore' THEN Red END) AS Trig_Res_HROTS_Red,
        MAX(CASE WHEN Program='Res' AND FieldName='HROTrainsStaffBefore' THEN Green END) AS Trig_Res_HROTS_Green,
        MAX(CASE WHEN Program='Res' AND FieldName='HROTrainsIndividualsBefore' THEN Red END) AS Trig_Res_HROTI_Red,
        MAX(CASE WHEN Program='Res' AND FieldName='HROTrainsIndividualsBefore' THEN Green END) AS Trig_Res_HROTI_Green,
        MAX(CASE WHEN Program='Res' AND FieldName='FSOTrainsStaffBefore' THEN Red END) AS Trig_Res_FSOTS_Red,
        MAX(CASE WHEN Program='Res' AND FieldName='FSOTrainsStaffBefore' THEN Green END) AS Trig_Res_FSOTS_Green,
        MAX(CASE WHEN Program='Res' AND FieldName='FSOTrainsIndividualsBefore' THEN Red END) AS Trig_Res_FSOTI_Red,
        MAX(CASE WHEN Program='Res' AND FieldName='FSOTrainsIndividualsBefore' THEN Green END) AS Trig_Res_FSOTI_Green
    FROM catExpirationTriggers
)
SELECT 
    e.*,
    
    -- Additional lookups from tblLocations (for main report VBA)
    loc.Department,
    loc.Cluster,
    loc.ClusterDescription,
    
    -- ========================================
    -- CLIENT FIELDS (RecordType = 'Client')
    -- ========================================
    
    -- DateISP calculations
    CASE 
        WHEN e.DateISP = '1900-01-01' THEN 'Missing'
        WHEN e.DateISP = '1900-01-02' THEN 'Optional'
        WHEN e.DateISP = '1900-01-03' THEN 'N/A'
        WHEN e.DateISP IS NULL THEN 'Missing'
        ELSE CONVERT(VARCHAR(10), e.DateISP, 101) -- MM/dd/yyyy format
    END AS DateISP_Display,
    
    CASE 
        WHEN e.DateISP IN ('1900-01-01', NULL) THEN 'RED'
        WHEN e.DateISP IN ('1900-01-02', '1900-01-03') THEN 'NORMAL'
        WHEN DATEDIFF(day, CAST(GETDATE() AS DATE), e.DateISP) < t.Trig_Indiv_ISP_Red THEN 'RED'
        WHEN DATEDIFF(day, CAST(GETDATE() AS DATE), e.DateISP) <= t.Trig_Indiv_ISP_Green THEN 'GREEN'
        ELSE 'NORMAL'
    END AS DateISP_Color,
    
    CASE 
        WHEN e.DateISP IN ('1900-01-01', '1900-01-02', '1900-01-03', NULL) THEN 0
        ELSE 1
    END AS DateISP_ShowDate,
    
    -- PSDue calculations (182 days before DateISP)
    CASE 
        WHEN e.DateISP IN ('1900-01-01', '1900-01-02', '1900-01-03', NULL) THEN NULL
        ELSE DATEADD(day, -182, e.DateISP)
    END AS PSDue_Calculated,
    
    CASE 
        WHEN e.DateISP IN ('1900-01-01', '1900-01-02', '1900-01-03', NULL) THEN 'N/A'
        ELSE CONVERT(VARCHAR(10), DATEADD(day, -182, e.DateISP), 101)
    END AS PSDue_Display,
    
    CASE 
        WHEN e.DateISP IN ('1900-01-01', '1900-01-02', '1900-01-03', NULL) THEN 'NORMAL'
        WHEN DATEDIFF(day, CAST(GETDATE() AS DATE), DATEADD(day, -182, e.DateISP)) <= t.Trig_Indiv_PSDue_Green 
             AND CAST(GETDATE() AS DATE) <= DATEADD(day, -182, e.DateISP) THEN 'GREEN'
        WHEN CAST(GETDATE() AS DATE) > DATEADD(day, -182, e.DateISP) THEN 'STRIKETHROUGH'
        ELSE 'NORMAL'
    END AS PSDue_Color,
    
    CASE 
        WHEN e.DateISP IN ('1900-01-01', '1900-01-02', '1900-01-03', NULL) THEN 1
        ELSE 0
    END AS PSDue_ShowText,
    
    -- DateConsentFormsSigned calculations (months based)
    CASE 
        WHEN e.DateConsentFormsSigned = '1900-01-01' THEN 'Missing'
        WHEN e.DateConsentFormsSigned = '1900-01-02' THEN 'Optional'
        WHEN e.DateConsentFormsSigned = '1900-01-03' THEN 'N/A'
        WHEN e.DateConsentFormsSigned IS NULL THEN 'Missing'
        ELSE CONVERT(VARCHAR(10), e.DateConsentFormsSigned, 101)
    END AS DateConsentFormsSigned_Display,
    
    CASE 
        WHEN e.DateConsentFormsSigned IN ('1900-01-01', NULL) THEN 'RED'
        WHEN e.DateConsentFormsSigned IN ('1900-01-02', '1900-01-03') THEN 'NORMAL'
        WHEN DATEDIFF(day, DATEADD(month, t.Trig_Indiv_CFS_Red, e.DateConsentFormsSigned), CAST(GETDATE() AS DATE)) > 0 THEN 'RED'
        WHEN DATEDIFF(day, DATEADD(month, t.Trig_Indiv_CFS_Green, e.DateConsentFormsSigned), CAST(GETDATE() AS DATE)) > 0 THEN 'GREEN'
        ELSE 'NORMAL'
    END AS DateConsentFormsSigned_Color,
    
    CASE 
        WHEN e.DateConsentFormsSigned IN ('1900-01-01', '1900-01-02', '1900-01-03', NULL) THEN 0
        ELSE 1
    END AS DateConsentFormsSigned_ShowDate,
    
    -- DateBMMExpires calculations
    CASE 
        WHEN e.DateBMMExpires = '1900-01-01' THEN 'Missing'
        WHEN e.DateBMMExpires = '1900-01-02' THEN 'Optional'
        WHEN e.DateBMMExpires = '1900-01-03' THEN 'N/A'
        WHEN e.DateBMMExpires IS NULL THEN 'Missing'
        ELSE CONVERT(VARCHAR(10), e.DateBMMExpires, 101)
    END AS DateBMMExpires_Display,
    
    CASE 
        WHEN e.DateBMMExpires IN ('1900-01-01', NULL) THEN 'RED'
        WHEN e.DateBMMExpires IN ('1900-01-02', '1900-01-03') THEN 'NORMAL'
        WHEN DATEDIFF(day, CAST(GETDATE() AS DATE), e.DateBMMExpires) < t.Trig_Indiv_BMMX_Red THEN 'RED'
        WHEN DATEDIFF(day, CAST(GETDATE() AS DATE), e.DateBMMExpires) <= t.Trig_Indiv_BMMX_Green THEN 'GREEN'
        ELSE 'NORMAL'
    END AS DateBMMExpires_Color,
    
    CASE 
        WHEN e.DateBMMExpires IN ('1900-01-01', '1900-01-02', '1900-01-03', NULL) THEN 0
        ELSE 1
    END AS DateBMMExpires_ShowDate,
    
    -- DateSPDAuthExpires calculations
    CASE 
        WHEN e.DateSPDAuthExpires = '1900-01-01' THEN 'Missing'
        WHEN e.DateSPDAuthExpires = '1900-01-02' THEN 'Optional'
        WHEN e.DateSPDAuthExpires = '1900-01-03' THEN 'N/A'
        WHEN e.DateSPDAuthExpires IS NULL THEN 'Missing'
        ELSE CONVERT(VARCHAR(10), e.DateSPDAuthExpires, 101)
    END AS DateSPDAuthExpires_Display,
    
    CASE 
        WHEN e.DateSPDAuthExpires IN ('1900-01-01', NULL) THEN 'RED'
        WHEN e.DateSPDAuthExpires IN ('1900-01-02', '1900-01-03') THEN 'NORMAL'
        WHEN DATEDIFF(day, CAST(GETDATE() AS DATE), e.DateSPDAuthExpires) < t.Trig_Indiv_SPDX_Red THEN 'RED'
        WHEN DATEDIFF(day, CAST(GETDATE() AS DATE), e.DateSPDAuthExpires) <= t.Trig_Indiv_SPDX_Green THEN 'GREEN'
        ELSE 'NORMAL'
    END AS DateSPDAuthExpires_Color,
    
    CASE 
        WHEN e.DateSPDAuthExpires IN ('1900-01-01', '1900-01-02', '1900-01-03', NULL) THEN 0
        ELSE 1
    END AS DateSPDAuthExpires_ShowDate,
    
    -- DateSignaturesDueBy calculations
    CASE 
        WHEN e.DateSignaturesDueBy = '1900-01-01' THEN 'Missing'
        WHEN e.DateSignaturesDueBy = '1900-01-02' THEN 'Optional'
        WHEN e.DateSignaturesDueBy = '1900-01-03' THEN 'N/A'
        WHEN e.DateSignaturesDueBy IS NULL THEN 'Missing'
        ELSE CONVERT(VARCHAR(10), e.DateSignaturesDueBy, 101)
    END AS DateSignaturesDueBy_Display,
    
    CASE 
        WHEN e.DateSignaturesDueBy IN ('1900-01-01', NULL) THEN 'RED'
        WHEN e.DateSignaturesDueBy IN ('1900-01-02', '1900-01-03') THEN 'NORMAL'
        WHEN DATEDIFF(day, CAST(GETDATE() AS DATE), e.DateSignaturesDueBy) < t.Trig_Indiv_SPDA_Red THEN 'RED'
        WHEN DATEDIFF(day, CAST(GETDATE() AS DATE), e.DateSignaturesDueBy) <= t.Trig_Indiv_SPDA_Green THEN 'GREEN'
        ELSE 'NORMAL'
    END AS DateSignaturesDueBy_Color,
    
    CASE 
        WHEN e.DateSignaturesDueBy IN ('1900-01-01', '1900-01-02', '1900-01-03', NULL) THEN 0
        ELSE 1
    END AS DateSignaturesDueBy_ShowDate,
    
    -- ========================================
    -- DAY & HOUSE FIELDS (RecordType = 'House')
    -- ========================================
    
    -- LastVehicleChecklistCompleted - Day program
    CASE 
        WHEN e.LastVehicleChecklistCompleted = '1900-01-01' THEN 'Missing'
        WHEN e.LastVehicleChecklistCompleted = '1900-01-02' THEN 'Optional'
        WHEN e.LastVehicleChecklistCompleted = '1900-01-03' THEN 'N/A'
        WHEN e.LastVehicleChecklistCompleted = '1900-01-05' THEN 'Pending'
        WHEN e.LastVehicleChecklistCompleted IS NULL THEN 'Missing'
        ELSE CONVERT(VARCHAR(10), e.LastVehicleChecklistCompleted, 101)
    END AS LastVehicleChecklistCompleted_Display_Day,
    
    CASE 
        WHEN e.LastVehicleChecklistCompleted IN ('1900-01-01', '1900-01-02', '1900-01-03', '1900-01-05', NULL) THEN 'NORMAL'
        WHEN DATEDIFF(day, e.LastVehicleChecklistCompleted, CAST(GETDATE() AS DATE)) >= t.Trig_Day_LVC_Red THEN 'RED'
        ELSE 'NORMAL'
    END AS LastVehicleChecklistCompleted_Color_Day,
    
    CASE 
        WHEN e.LastVehicleChecklistCompleted IN ('1900-01-01', '1900-01-02', '1900-01-03', '1900-01-05', NULL) THEN 0
        ELSE 1
    END AS LastVehicleChecklistCompleted_ShowDate_Day,
    
    -- LastVehicleChecklistCompleted - Residential program
    CASE 
        WHEN e.LastVehicleChecklistCompleted = '1900-01-01' THEN 'Missing'
        WHEN e.LastVehicleChecklistCompleted = '1900-01-02' THEN 'Optional'
        WHEN e.LastVehicleChecklistCompleted = '1900-01-03' THEN 'N/A'
        WHEN e.LastVehicleChecklistCompleted = '1900-01-05' THEN 'Pending'
        WHEN e.LastVehicleChecklistCompleted IS NULL THEN 'Missing'
        ELSE CONVERT(VARCHAR(10), e.LastVehicleChecklistCompleted, 101)
    END AS LastVehicleChecklistCompleted_Display_Res,
    
    CASE 
        WHEN e.LastVehicleChecklistCompleted IN ('1900-01-01', '1900-01-02', '1900-01-03', '1900-01-05', NULL) THEN 'NORMAL'
        WHEN DATEDIFF(day, e.LastVehicleChecklistCompleted, CAST(GETDATE() AS DATE)) >= t.Trig_Res_LVC_Red THEN 'RED'
        ELSE 'NORMAL'
    END AS LastVehicleChecklistCompleted_Color_Res,
    
    CASE 
        WHEN e.LastVehicleChecklistCompleted IN ('1900-01-01', '1900-01-02', '1900-01-03', '1900-01-05', NULL) THEN 0
        ELSE 1
    END AS LastVehicleChecklistCompleted_ShowDate_Res,
    
    -- DAYStaffTrainedInPrivacyBefore
    CASE 
        WHEN e.DAYStaffTrainedInPrivacyBefore = '1900-01-01' THEN 'Missing'
        WHEN e.DAYStaffTrainedInPrivacyBefore = '1900-01-02' THEN 'Optional'
        WHEN e.DAYStaffTrainedInPrivacyBefore = '1900-01-03' THEN 'N/A'
        WHEN e.DAYStaffTrainedInPrivacyBefore = '1900-01-05' THEN 'Pending'
        WHEN e.DAYStaffTrainedInPrivacyBefore IS NULL THEN 'Missing'
        ELSE CONVERT(VARCHAR(10), e.DAYStaffTrainedInPrivacyBefore, 101)
    END AS DAYStaffTrainedInPrivacyBefore_Display,
    
    CASE 
        WHEN e.DAYStaffTrainedInPrivacyBefore IN ('1900-01-01', '1900-01-02', '1900-01-03', '1900-01-05', NULL) THEN 'NORMAL'
        WHEN DATEDIFF(day, CAST(GETDATE() AS DATE), e.DAYStaffTrainedInPrivacyBefore) < t.Trig_Day_STP_Red THEN 'RED'
        WHEN DATEDIFF(day, CAST(GETDATE() AS DATE), e.DAYStaffTrainedInPrivacyBefore) <= t.Trig_Day_STP_Green THEN 'GREEN'
        ELSE 'NORMAL'
    END AS DAYStaffTrainedInPrivacyBefore_Color,
    
    CASE 
        WHEN e.DAYStaffTrainedInPrivacyBefore IN ('1900-01-01', '1900-01-02', '1900-01-03', '1900-01-05', NULL) THEN 0
        ELSE 1
    END AS DAYStaffTrainedInPrivacyBefore_ShowDate,
    
    -- DAYAllPlansReviewedByStaffBefore
    CASE 
        WHEN e.DAYAllPlansReviewedByStaffBefore = '1900-01-01' THEN 'Missing'
        WHEN e.DAYAllPlansReviewedByStaffBefore = '1900-01-02' THEN 'Optional'
        WHEN e.DAYAllPlansReviewedByStaffBefore = '1900-01-03' THEN 'N/A'
        WHEN e.DAYAllPlansReviewedByStaffBefore = '1900-01-05' THEN 'Pending'
        WHEN e.DAYAllPlansReviewedByStaffBefore IS NULL THEN 'Missing'
        ELSE CONVERT(VARCHAR(10), e.DAYAllPlansReviewedByStaffBefore, 101)
    END AS DAYAllPlansReviewedByStaffBefore_Display,
    
    CASE 
        WHEN e.DAYAllPlansReviewedByStaffBefore IN ('1900-01-01', '1900-01-02', '1900-01-03', '1900-01-05', NULL) THEN 'NORMAL'
        WHEN DATEDIFF(day, CAST(GETDATE() AS DATE), e.DAYAllPlansReviewedByStaffBefore) < t.Trig_Day_APRS_Red THEN 'RED'
        WHEN DATEDIFF(day, CAST(GETDATE() AS DATE), e.DAYAllPlansReviewedByStaffBefore) <= t.Trig_Day_APRS_Green THEN 'GREEN'
        ELSE 'NORMAL'
    END AS DAYAllPlansReviewedByStaffBefore_Color,
    
    CASE 
        WHEN e.DAYAllPlansReviewedByStaffBefore IN ('1900-01-01', '1900-01-02', '1900-01-03', '1900-01-05', NULL) THEN 0
        ELSE 1
    END AS DAYAllPlansReviewedByStaffBefore_ShowDate,
    
    -- DAYQtrlySafetyChecklistDueBy
    CASE 
        WHEN e.DAYQtrlySafetyChecklistDueBy = '1900-01-01' THEN 'Missing'
        WHEN e.DAYQtrlySafetyChecklistDueBy = '1900-01-02' THEN 'Optional'
        WHEN e.DAYQtrlySafetyChecklistDueBy = '1900-01-03' THEN 'N/A'
        WHEN e.DAYQtrlySafetyChecklistDueBy = '1900-01-05' THEN 'Pending'
        WHEN e.DAYQtrlySafetyChecklistDueBy IS NULL THEN 'Missing'
        ELSE CONVERT(VARCHAR(10), e.DAYQtrlySafetyChecklistDueBy, 101)
    END AS DAYQtrlySafetyChecklistDueBy_Display,
    
    CASE 
        WHEN e.DAYQtrlySafetyChecklistDueBy IN ('1900-01-01', '1900-01-02', '1900-01-03', '1900-01-05', NULL) THEN 'NORMAL'
        WHEN DATEDIFF(day, CAST(GETDATE() AS DATE), e.DAYQtrlySafetyChecklistDueBy) < t.Trig_Day_QSR_Red THEN 'RED'
        WHEN DATEDIFF(day, CAST(GETDATE() AS DATE), e.DAYQtrlySafetyChecklistDueBy) <= t.Trig_Day_QSR_Green THEN 'GREEN'
        ELSE 'NORMAL'
    END AS DAYQtrlySafetyChecklistDueBy_Color,
    
    CASE 
        WHEN e.DAYQtrlySafetyChecklistDueBy IN ('1900-01-01', '1900-01-02', '1900-01-03', '1900-01-05', NULL) THEN 0
        ELSE 1
    END AS DAYQtrlySafetyChecklistDueBy_ShowDate,
    
    -- MostRecentAsleepFireDrill (Residential only, calculates based on +14 months)
    CASE 
        WHEN e.MostRecentAsleepFireDrill = '1900-01-01' THEN 'Missing'
        WHEN e.MostRecentAsleepFireDrill = '1900-01-02' THEN 'Optional'
        WHEN e.MostRecentAsleepFireDrill = '1900-01-03' THEN 'N/A'
        WHEN e.MostRecentAsleepFireDrill = '1900-01-05' THEN 'Pending'
        WHEN e.MostRecentAsleepFireDrill IS NULL THEN 'Missing'
        ELSE CONVERT(VARCHAR(10), e.MostRecentAsleepFireDrill, 101)
    END AS MostRecentAsleepFireDrill_Display,
    
    CASE 
        WHEN e.MostRecentAsleepFireDrill IN ('1900-01-01', '1900-01-02', '1900-01-03', '1900-01-05', NULL) THEN 'NORMAL'
        WHEN DATEDIFF(day, CAST(GETDATE() AS DATE), DATEADD(month, 14, e.MostRecentAsleepFireDrill)) < t.Trig_Res_MRFD_Red THEN 'RED'
        WHEN DATEDIFF(day, CAST(GETDATE() AS DATE), DATEADD(month, 14, e.MostRecentAsleepFireDrill)) <= t.Trig_Res_MRFD_Green THEN 'GREEN'
        ELSE 'NORMAL'
    END AS MostRecentAsleepFireDrill_Color,
    
    CASE 
        WHEN e.MostRecentAsleepFireDrill IN ('1900-01-01', '1900-01-02', '1900-01-03', '1900-01-05', NULL) THEN 0
        ELSE 1
    END AS MostRecentAsleepFireDrill_ShowDate,
    
    -- HousePlansReviewedByStaffBefore (Residential)
    CASE 
        WHEN e.HousePlansReviewedByStaffBefore = '1900-01-01' THEN 'Missing'
        WHEN e.HousePlansReviewedByStaffBefore = '1900-01-02' THEN 'Optional'
        WHEN e.HousePlansReviewedByStaffBefore = '1900-01-03' THEN 'N/A'
        WHEN e.HousePlansReviewedByStaffBefore = '1900-01-05' THEN 'Pending'
        WHEN e.HousePlansReviewedByStaffBefore IS NULL THEN 'Missing'
        ELSE CONVERT(VARCHAR(10), e.HousePlansReviewedByStaffBefore, 101)
    END AS HousePlansReviewedByStaffBefore_Display,
    
    CASE 
        WHEN e.HousePlansReviewedByStaffBefore IN ('1900-01-01', '1900-01-02', '1900-01-03', '1900-01-05', NULL) THEN 'NORMAL'
        WHEN DATEDIFF(day, CAST(GETDATE() AS DATE), e.HousePlansReviewedByStaffBefore) < t.Trig_Res_HPR_Red THEN 'RED'
        WHEN DATEDIFF(day, CAST(GETDATE() AS DATE), e.HousePlansReviewedByStaffBefore) <= t.Trig_Res_HPR_Green THEN 'GREEN'
        ELSE 'NORMAL'
    END AS HousePlansReviewedByStaffBefore_Color,
    
    CASE 
        WHEN e.HousePlansReviewedByStaffBefore IN ('1900-01-01', '1900-01-02', '1900-01-03', '1900-01-05', NULL) THEN 0
        ELSE 1
    END AS HousePlansReviewedByStaffBefore_ShowDate,
    
    -- HouseSafetyPlanExpires (Residential)
    CASE 
        WHEN e.HouseSafetyPlanExpires = '1900-01-01' THEN 'Missing'
        WHEN e.HouseSafetyPlanExpires = '1900-01-02' THEN 'Optional'
        WHEN e.HouseSafetyPlanExpires = '1900-01-03' THEN 'N/A'
        WHEN e.HouseSafetyPlanExpires = '1900-01-05' THEN 'Pending'
        WHEN e.HouseSafetyPlanExpires IS NULL THEN 'Missing'
        ELSE CONVERT(VARCHAR(10), e.HouseSafetyPlanExpires, 101)
    END AS HouseSafetyPlanExpires_Display,
    
    CASE 
        WHEN e.HouseSafetyPlanExpires IN ('1900-01-01', '1900-01-02', '1900-01-03', '1900-01-05', NULL) THEN 'NORMAL'
        WHEN DATEDIFF(day, CAST(GETDATE() AS DATE), e.HouseSafetyPlanExpires) < t.Trig_Res_HSPE_Red THEN 'RED'
        WHEN DATEDIFF(day, CAST(GETDATE() AS DATE), e.HouseSafetyPlanExpires) <= t.Trig_Res_HSPE_Green THEN 'GREEN'
        ELSE 'NORMAL'
    END AS HouseSafetyPlanExpires_Color,
    
    CASE 
        WHEN e.HouseSafetyPlanExpires IN ('1900-01-01', '1900-01-02', '1900-01-03', '1900-01-05', NULL) THEN 0
        ELSE 1
    END AS HouseSafetyPlanExpires_ShowDate,
    
    -- MAPChecklistCompleted (Residential)
    CASE 
        WHEN e.MAPChecklistCompleted = '1900-01-01' THEN 'Missing'
        WHEN e.MAPChecklistCompleted = '1900-01-02' THEN 'Optional'
        WHEN e.MAPChecklistCompleted = '1900-01-03' THEN 'N/A'
        WHEN e.MAPChecklistCompleted = '1900-01-05' THEN 'Pending'
        WHEN e.MAPChecklistCompleted IS NULL THEN 'Missing'
        ELSE CONVERT(VARCHAR(10), e.MAPChecklistCompleted, 101)
    END AS MAPChecklistCompleted_Display,
    
    CASE 
        WHEN e.MAPChecklistCompleted IN ('1900-01-01', '1900-01-02', '1900-01-03', '1900-01-05', NULL) THEN 'NORMAL'
        WHEN DATEDIFF(day, e.MAPChecklistCompleted, CAST(GETDATE() AS DATE)) >= t.Trig_Res_MAP_Red THEN 'RED'
        ELSE 'NORMAL'
    END AS MAPChecklistCompleted_Color,
    
    CASE 
        WHEN e.MAPChecklistCompleted IN ('1900-01-01', '1900-01-02', '1900-01-03', '1900-01-05', NULL) THEN 0
        ELSE 1
    END AS MAPChecklistCompleted_ShowDate,
    
    -- HROTrainsStaffBefore - Day program
    CASE 
        WHEN e.HROTrainsStaffBefore = '1900-01-01' THEN 'Missing'
        WHEN e.HROTrainsStaffBefore = '1900-01-02' THEN 'Optional'
        WHEN e.HROTrainsStaffBefore = '1900-01-03' THEN 'N/A'
        WHEN e.HROTrainsStaffBefore = '1900-01-05' THEN 'Pending'
        WHEN e.HROTrainsStaffBefore IS NULL THEN 'Missing'
        ELSE CONVERT(VARCHAR(10), e.HROTrainsStaffBefore, 101)
    END AS HROTrainsStaffBefore_Display_Day,
    
    CASE 
        WHEN e.HROTrainsStaffBefore IN ('1900-01-01', '1900-01-02', '1900-01-03', '1900-01-05', NULL) THEN 'NORMAL'
        WHEN DATEDIFF(day, CAST(GETDATE() AS DATE), e.HROTrainsStaffBefore) < t.Trig_Day_HROTS_Red THEN 'RED'
        WHEN DATEDIFF(day, CAST(GETDATE() AS DATE), e.HROTrainsStaffBefore) <= t.Trig_Day_HROTS_Green THEN 'GREEN'
        ELSE 'NORMAL'
    END AS HROTrainsStaffBefore_Color_Day,
    
    CASE 
        WHEN e.HROTrainsStaffBefore IN ('1900-01-01', '1900-01-02', '1900-01-03', '1900-01-05', NULL) THEN 0
        ELSE 1
    END AS HROTrainsStaffBefore_ShowDate_Day,
    
    -- HROTrainsStaffBefore - Residential program
    CASE 
        WHEN e.HROTrainsStaffBefore = '1900-01-01' THEN 'Missing'
        WHEN e.HROTrainsStaffBefore = '1900-01-02' THEN 'Optional'
        WHEN e.HROTrainsStaffBefore = '1900-01-03' THEN 'N/A'
        WHEN e.HROTrainsStaffBefore = '1900-01-05' THEN 'Pending'
        WHEN e.HROTrainsStaffBefore IS NULL THEN 'Missing'
        ELSE CONVERT(VARCHAR(10), e.HROTrainsStaffBefore, 101)
    END AS HROTrainsStaffBefore_Display_Res,
    
    CASE 
        WHEN e.HROTrainsStaffBefore IN ('1900-01-01', '1900-01-02', '1900-01-03', '1900-01-05', NULL) THEN 'NORMAL'
        WHEN DATEDIFF(day, CAST(GETDATE() AS DATE), e.HROTrainsStaffBefore) < t.Trig_Res_HROTS_Red THEN 'RED'
        WHEN DATEDIFF(day, CAST(GETDATE() AS DATE), e.HROTrainsStaffBefore) <= t.Trig_Res_HROTS_Green THEN 'GREEN'
        ELSE 'NORMAL'
    END AS HROTrainsStaffBefore_Color_Res,
    
    CASE 
        WHEN e.HROTrainsStaffBefore IN ('1900-01-01', '1900-01-02', '1900-01-03', '1900-01-05', NULL) THEN 0
        ELSE 1
    END AS HROTrainsStaffBefore_ShowDate_Res,
    
    -- HROTrainsIndividualsBefore - Day program
    CASE 
        WHEN e.HROTrainsIndividualsBefore = '1900-01-01' THEN 'Missing'
        WHEN e.HROTrainsIndividualsBefore = '1900-01-02' THEN 'Optional'
        WHEN e.HROTrainsIndividualsBefore = '1900-01-03' THEN 'N/A'
        WHEN e.HROTrainsIndividualsBefore = '1900-01-05' THEN 'Pending'
        WHEN e.HROTrainsIndividualsBefore IS NULL THEN 'Missing'
        ELSE CONVERT(VARCHAR(10), e.HROTrainsIndividualsBefore, 101)
    END AS HROTrainsIndividualsBefore_Display_Day,
    
    CASE 
        WHEN e.HROTrainsIndividualsBefore IN ('1900-01-01', '1900-01-02', '1900-01-03', '1900-01-05', NULL) THEN 'NORMAL'
        WHEN DATEDIFF(day, CAST(GETDATE() AS DATE), e.HROTrainsIndividualsBefore) < t.Trig_Day_HROTI_Red THEN 'RED'
        WHEN DATEDIFF(day, CAST(GETDATE() AS DATE), e.HROTrainsIndividualsBefore) <= t.Trig_Day_HROTI_Green THEN 'GREEN'
        ELSE 'NORMAL'
    END AS HROTrainsIndividualsBefore_Color_Day,
    
    CASE 
        WHEN e.HROTrainsIndividualsBefore IN ('1900-01-01', '1900-01-02', '1900-01-03', '1900-01-05', NULL) THEN 0
        ELSE 1
    END AS HROTrainsIndividualsBefore_ShowDate_Day,
    
    -- HROTrainsIndividualsBefore - Residential program
    CASE 
        WHEN e.HROTrainsIndividualsBefore = '1900-01-01' THEN 'Missing'
        WHEN e.HROTrainsIndividualsBefore = '1900-01-02' THEN 'Optional'
        WHEN e.HROTrainsIndividualsBefore = '1900-01-03' THEN 'N/A'
        WHEN e.HROTrainsIndividualsBefore = '1900-01-05' THEN 'Pending'
        WHEN e.HROTrainsIndividualsBefore IS NULL THEN 'Missing'
        ELSE CONVERT(VARCHAR(10), e.HROTrainsIndividualsBefore, 101)
    END AS HROTrainsIndividualsBefore_Display_Res,
    
    CASE 
        WHEN e.HROTrainsIndividualsBefore IN ('1900-01-01', '1900-01-02', '1900-01-03', '1900-01-05', NULL) THEN 'NORMAL'
        WHEN DATEDIFF(day, CAST(GETDATE() AS DATE), e.HROTrainsIndividualsBefore) < t.Trig_Res_HROTI_Red THEN 'RED'
        WHEN DATEDIFF(day, CAST(GETDATE() AS DATE), e.HROTrainsIndividualsBefore) <= t.Trig_Res_HROTI_Green THEN 'GREEN'
        ELSE 'NORMAL'
    END AS HROTrainsIndividualsBefore_Color_Res,
    
    CASE 
        WHEN e.HROTrainsIndividualsBefore IN ('1900-01-01', '1900-01-02', '1900-01-03', '1900-01-05', NULL) THEN 0
        ELSE 1
    END AS HROTrainsIndividualsBefore_ShowDate_Res,
    
    -- FSOTrainsStaffBefore - Day program
    CASE 
        WHEN e.FSOTrainsStaffBefore = '1900-01-01' THEN 'Missing'
        WHEN e.FSOTrainsStaffBefore = '1900-01-02' THEN 'Optional'
        WHEN e.FSOTrainsStaffBefore = '1900-01-03' THEN 'N/A'
        WHEN e.FSOTrainsStaffBefore = '1900-01-05' THEN 'Pending'
        WHEN e.FSOTrainsStaffBefore IS NULL THEN 'Missing'
        ELSE CONVERT(VARCHAR(10), e.FSOTrainsStaffBefore, 101)
    END AS FSOTrainsStaffBefore_Display_Day,
    
    CASE 
        WHEN e.FSOTrainsStaffBefore IN ('1900-01-01', '1900-01-02', '1900-01-03', '1900-01-05', NULL) THEN 'NORMAL'
        WHEN DATEDIFF(day, CAST(GETDATE() AS DATE), e.FSOTrainsStaffBefore) < t.Trig_Day_FSOTS_Red THEN 'RED'
        WHEN DATEDIFF(day, CAST(GETDATE() AS DATE), e.FSOTrainsStaffBefore) <= t.Trig_Day_FSOTS_Green THEN 'GREEN'
        ELSE 'NORMAL'
    END AS FSOTrainsStaffBefore_Color_Day,
    
    CASE 
        WHEN e.FSOTrainsStaffBefore IN ('1900-01-01', '1900-01-02', '1900-01-03', '1900-01-05', NULL) THEN 0
        ELSE 1
    END AS FSOTrainsStaffBefore_ShowDate_Day,
    
    -- FSOTrainsStaffBefore - Residential program
    CASE 
        WHEN e.FSOTrainsStaffBefore = '1900-01-01' THEN 'Missing'
        WHEN e.FSOTrainsStaffBefore = '1900-01-02' THEN 'Optional'
        WHEN e.FSOTrainsStaffBefore = '1900-01-03' THEN 'N/A'
        WHEN e.FSOTrainsStaffBefore = '1900-01-05' THEN 'Pending'
        WHEN e.FSOTrainsStaffBefore IS NULL THEN 'Missing'
        ELSE CONVERT(VARCHAR(10), e.FSOTrainsStaffBefore, 101)
    END AS FSOTrainsStaffBefore_Display_Res,
    
    CASE 
        WHEN e.FSOTrainsStaffBefore IN ('1900-01-01', '1900-01-02', '1900-01-03', '1900-01-05', NULL) THEN 'NORMAL'
        WHEN DATEDIFF(day, CAST(GETDATE() AS DATE), e.FSOTrainsStaffBefore) < t.Trig_Res_FSOTS_Red THEN 'RED'
        WHEN DATEDIFF(day, CAST(GETDATE() AS DATE), e.FSOTrainsStaffBefore) <= t.Trig_Res_FSOTS_Green THEN 'GREEN'
        ELSE 'NORMAL'
    END AS FSOTrainsStaffBefore_Color_Res,
    
    CASE 
        WHEN e.FSOTrainsStaffBefore IN ('1900-01-01', '1900-01-02', '1900-01-03', '1900-01-05', NULL) THEN 0
        ELSE 1
    END AS FSOTrainsStaffBefore_ShowDate_Res,
    
    -- FSOTrainsIndividualsBefore - Day program
    CASE 
        WHEN e.FSOTrainsIndividualsBefore = '1900-01-01' THEN 'Missing'
        WHEN e.FSOTrainsIndividualsBefore = '1900-01-02' THEN 'Optional'
        WHEN e.FSOTrainsIndividualsBefore = '1900-01-03' THEN 'N/A'
        WHEN e.FSOTrainsIndividualsBefore = '1900-01-05' THEN 'Pending'
        WHEN e.FSOTrainsIndividualsBefore IS NULL THEN 'Missing'
        ELSE CONVERT(VARCHAR(10), e.FSOTrainsIndividualsBefore, 101)
    END AS FSOTrainsIndividualsBefore_Display_Day,
    
    CASE 
        WHEN e.FSOTrainsIndividualsBefore IN ('1900-01-01', '1900-01-02', '1900-01-03', '1900-01-05', NULL) THEN 'NORMAL'
        WHEN DATEDIFF(day, CAST(GETDATE() AS DATE), e.FSOTrainsIndividualsBefore) < t.Trig_Day_FSOTI_Red THEN 'RED'
        WHEN DATEDIFF(day, CAST(GETDATE() AS DATE), e.FSOTrainsIndividualsBefore) <= t.Trig_Day_FSOTI_Green THEN 'GREEN'
        ELSE 'NORMAL'
    END AS FSOTrainsIndividualsBefore_Color_Day,
    
    CASE 
        WHEN e.FSOTrainsIndividualsBefore IN ('1900-01-01', '1900-01-02', '1900-01-03', '1900-01-05', NULL) THEN 0
        ELSE 1
    END AS FSOTrainsIndividualsBefore_ShowDate_Day,
    
    -- FSOTrainsIndividualsBefore - Residential program
    CASE 
        WHEN e.FSOTrainsIndividualsBefore = '1900-01-01' THEN 'Missing'
        WHEN e.FSOTrainsIndividualsBefore = '1900-01-02' THEN 'Optional'
        WHEN e.FSOTrainsIndividualsBefore = '1900-01-03' THEN 'N/A'
        WHEN e.FSOTrainsIndividualsBefore = '1900-01-05' THEN 'Pending'
        WHEN e.FSOTrainsIndividualsBefore IS NULL THEN 'Missing'
        ELSE CONVERT(VARCHAR(10), e.FSOTrainsIndividualsBefore, 101)
    END AS FSOTrainsIndividualsBefore_Display_Res,
    
    CASE 
        WHEN e.FSOTrainsIndividualsBefore IN ('1900-01-01', '1900-01-02', '1900-01-03', '1900-01-05', NULL) THEN 'NORMAL'
        WHEN DATEDIFF(day, CAST(GETDATE() AS DATE), e.FSOTrainsIndividualsBefore) < t.Trig_Res_FSOTI_Red THEN 'RED'
        WHEN DATEDIFF(day, CAST(GETDATE() AS DATE), e.FSOTrainsIndividualsBefore) <= t.Trig_Res_FSOTI_Green THEN 'GREEN'
        ELSE 'NORMAL'
    END AS FSOTrainsIndividualsBefore_Color_Res,
    
    CASE 
        WHEN e.FSOTrainsIndividualsBefore IN ('1900-01-01', '1900-01-02', '1900-01-03', '1900-01-05', NULL) THEN 0
        ELSE 1
    END AS FSOTrainsIndividualsBefore_ShowDate_Res,
    
    -- ========================================
    -- NAME FORMATTING (HumanRightsOfficer, FireSafetyOfficer)
    -- ========================================
    
    -- HumanRightsOfficer formatted (LastName, FirstName -> FirstName LastName)
    CASE 
        WHEN e.HumanRightsOfficer IS NULL OR LTRIM(RTRIM(e.HumanRightsOfficer)) = '' THEN NULL
        WHEN CHARINDEX(',', e.HumanRightsOfficer) = 0 THEN NULL
        ELSE 
            LTRIM(RTRIM(SUBSTRING(e.HumanRightsOfficer, CHARINDEX(',', e.HumanRightsOfficer) + 1, 255))) + ' ' +
            LTRIM(RTRIM(SUBSTRING(e.HumanRightsOfficer, 1, CHARINDEX(',', e.HumanRightsOfficer) - 1)))
    END AS HumanRightsOfficer_Formatted,
    
    CASE 
        WHEN e.HumanRightsOfficer IS NULL OR LTRIM(RTRIM(e.HumanRightsOfficer)) = '' THEN 1
        WHEN CHARINDEX(',', e.HumanRightsOfficer) = 0 THEN 1
        ELSE 0
    END AS HumanRightsOfficer_IsBlank,
    
    -- FireSafetyOfficer formatted (LastName, FirstName -> FirstName LastName)
    CASE 
        WHEN e.FireSafetyOfficer IS NULL OR LTRIM(RTRIM(e.FireSafetyOfficer)) = '' THEN NULL
        WHEN CHARINDEX(',', e.FireSafetyOfficer) = 0 THEN NULL
        ELSE 
            LTRIM(RTRIM(SUBSTRING(e.FireSafetyOfficer, CHARINDEX(',', e.FireSafetyOfficer) + 1, 255))) + ' ' +
            LTRIM(RTRIM(SUBSTRING(e.FireSafetyOfficer, 1, CHARINDEX(',', e.FireSafetyOfficer) - 1)))
    END AS FireSafetyOfficer_Formatted,
    
    CASE 
        WHEN e.FireSafetyOfficer IS NULL OR LTRIM(RTRIM(e.FireSafetyOfficer)) = '' THEN 1
        WHEN CHARINDEX(',', e.FireSafetyOfficer) = 0 THEN 1
        ELSE 0
    END AS FireSafetyOfficer_IsBlank

FROM tblExpirations e
CROSS JOIN TriggerValues t
LEFT JOIN tblLocations loc ON e.Location = loc.GPName
WHERE e.RecordType IN ('Client', 'Staff', 'House');

GO
