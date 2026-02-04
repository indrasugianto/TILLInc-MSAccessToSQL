/****** Object:  View [dbo].[vw_ExpirationsFormatted]    Script Date: 2/4/2026 1:15:26 PM ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO



create or alter view [dbo].[vw_ExpirationsFormatted]
as
with TriggerValues
as (
   -- Cache all trigger values to avoid repeated lookups
   select max(   case
                     when Section = 'Individuals'
                          and FieldName = 'DateISP' then
                         Red
                 end
             ) as Trig_Indiv_ISP_Red,
          max(   case
                     when Section = 'Individuals'
                          and FieldName = 'DateISP' then
                         Green
                 end
             ) as Trig_Indiv_ISP_Green,
          max(   case
                     when Section = 'Individuals'
                          and FieldName = 'PSDue' then
                         Green
                 end
             ) as Trig_Indiv_PSDue_Green,
          max(   case
                     when Section = 'Individuals'
                          and FieldName = 'DateConsentFormsSigned' then
                         Red
                 end
             ) as Trig_Indiv_CFS_Red,
          max(   case
                     when Section = 'Individuals'
                          and FieldName = 'DateConsentFormsSigned' then
                         Green
                 end
             ) as Trig_Indiv_CFS_Green,
          max(   case
                     when Section = 'Individuals'
                          and FieldName = 'DateBMMExpires' then
                         Red
                 end
             ) as Trig_Indiv_BMMX_Red,
          max(   case
                     when Section = 'Individuals'
                          and FieldName = 'DateBMMExpires' then
                         Green
                 end
             ) as Trig_Indiv_BMMX_Green,
          max(   case
                     when Section = 'Individuals'
                          and FieldName = 'DateSPDAuthExpires' then
                         Red
                 end
             ) as Trig_Indiv_SPDX_Red,
          max(   case
                     when Section = 'Individuals'
                          and FieldName = 'DateSPDAuthExpires' then
                         Green
                 end
             ) as Trig_Indiv_SPDX_Green,
          max(   case
                     when Section = 'Individuals'
                          and FieldName = 'DateSignaturesDueBy' then
                         Red
                 end
             ) as Trig_Indiv_SPDA_Red,
          max(   case
                     when Section = 'Individuals'
                          and FieldName = 'DateSignaturesDueBy' then
                         Green
                 end
             ) as Trig_Indiv_SPDA_Green,

          -- Day program triggers
          max(   case
                     when Program = 'Day'
                          and FieldName = 'LastVehicleChecklistCompleted' then
                         Red
                 end
             ) as Trig_Day_LVC_Red,
          max(   case
                     when Program = 'Day'
                          and FieldName = 'DAYStaffTrainedInPrivacyBefore' then
                         Red
                 end
             ) as Trig_Day_STP_Red,
          max(   case
                     when Program = 'Day'
                          and FieldName = 'DAYStaffTrainedInPrivacyBefore' then
                         Green
                 end
             ) as Trig_Day_STP_Green,
          max(   case
                     when Program = 'Day'
                          and FieldName = 'DAYAllPlansReviewedByStaffBefore' then
                         Red
                 end
             ) as Trig_Day_APRS_Red,
          max(   case
                     when Program = 'Day'
                          and FieldName = 'DAYAllPlansReviewedByStaffBefore' then
                         Green
                 end
             ) as Trig_Day_APRS_Green,
          max(   case
                     when Program = 'Day'
                          and FieldName = 'DAYQtrlySafetyChecklistDueBy' then
                         Red
                 end
             ) as Trig_Day_QSR_Red,
          max(   case
                     when Program = 'Day'
                          and FieldName = 'DAYQtrlySafetyChecklistDueBy' then
                         Green
                 end
             ) as Trig_Day_QSR_Green,
          max(   case
                     when Program = 'Day'
                          and FieldName = 'HROTrainsStaffBefore' then
                         Red
                 end
             ) as Trig_Day_HROTS_Red,
          max(   case
                     when Program = 'Day'
                          and FieldName = 'HROTrainsStaffBefore' then
                         Green
                 end
             ) as Trig_Day_HROTS_Green,
          max(   case
                     when Program = 'Day'
                          and FieldName = 'HROTrainsIndividualsBefore' then
                         Red
                 end
             ) as Trig_Day_HROTI_Red,
          max(   case
                     when Program = 'Day'
                          and FieldName = 'HROTrainsIndividualsBefore' then
                         Green
                 end
             ) as Trig_Day_HROTI_Green,
          max(   case
                     when Program = 'Day'
                          and FieldName = 'FSOTrainsStaffBefore' then
                         Red
                 end
             ) as Trig_Day_FSOTS_Red,
          max(   case
                     when Program = 'Day'
                          and FieldName = 'FSOTrainsStaffBefore' then
                         Green
                 end
             ) as Trig_Day_FSOTS_Green,
          max(   case
                     when Program = 'Day'
                          and FieldName = 'FSOTrainsIndividualsBefore' then
                         Red
                 end
             ) as Trig_Day_FSOTI_Red,
          max(   case
                     when Program = 'Day'
                          and FieldName = 'FSOTrainsIndividualsBefore' then
                         Green
                 end
             ) as Trig_Day_FSOTI_Green,

          -- Residential program triggers
          max(   case
                     when Program = 'Res'
                          and FieldName = 'LastVehicleChecklistCompleted' then
                         Red
                 end
             ) as Trig_Res_LVC_Red,
          max(   case
                     when Program = 'Res'
                          and FieldName = 'MostRecentAsleepFireDrill' then
                         Red
                 end
             ) as Trig_Res_MRFD_Red,
          max(   case
                     when Program = 'Res'
                          and FieldName = 'MostRecentAsleepFireDrill' then
                         Green
                 end
             ) as Trig_Res_MRFD_Green,
          max(   case
                     when Program = 'Res'
                          and FieldName = 'HousePlansReviewedByStaffBefore' then
                         Red
                 end
             ) as Trig_Res_HPR_Red,
          max(   case
                     when Program = 'Res'
                          and FieldName = 'HousePlansReviewedByStaffBefore' then
                         Green
                 end
             ) as Trig_Res_HPR_Green,
          max(   case
                     when Program = 'Res'
                          and FieldName = 'HouseSafetyPlanExpires' then
                         Red
                 end
             ) as Trig_Res_HSPE_Red,
          max(   case
                     when Program = 'Res'
                          and FieldName = 'HouseSafetyPlanExpires' then
                         Green
                 end
             ) as Trig_Res_HSPE_Green,
          max(   case
                     when Program = 'Res'
                          and FieldName = 'MAPChecklistCompleted' then
                         Red
                 end
             ) as Trig_Res_MAP_Red,
          max(   case
                     when Program = 'Res'
                          and FieldName = 'HROTrainsStaffBefore' then
                         Red
                 end
             ) as Trig_Res_HROTS_Red,
          max(   case
                     when Program = 'Res'
                          and FieldName = 'HROTrainsStaffBefore' then
                         Green
                 end
             ) as Trig_Res_HROTS_Green,
          max(   case
                     when Program = 'Res'
                          and FieldName = 'HROTrainsIndividualsBefore' then
                         Red
                 end
             ) as Trig_Res_HROTI_Red,
          max(   case
                     when Program = 'Res'
                          and FieldName = 'HROTrainsIndividualsBefore' then
                         Green
                 end
             ) as Trig_Res_HROTI_Green,
          max(   case
                     when Program = 'Res'
                          and FieldName = 'FSOTrainsStaffBefore' then
                         Red
                 end
             ) as Trig_Res_FSOTS_Red,
          max(   case
                     when Program = 'Res'
                          and FieldName = 'FSOTrainsStaffBefore' then
                         Green
                 end
             ) as Trig_Res_FSOTS_Green,
          max(   case
                     when Program = 'Res'
                          and FieldName = 'FSOTrainsIndividualsBefore' then
                         Red
                 end
             ) as Trig_Res_FSOTI_Red,
          max(   case
                     when Program = 'Res'
                          and FieldName = 'FSOTrainsIndividualsBefore' then
                         Green
                 end
             ) as Trig_Res_FSOTI_Green,
          -- Staff triggers (Section = 'Staff', Program IS NULL in catExpirationTriggers)
          max(   case
                     when Section = 'Staff'
                          and Program is null
                          and FieldName = 'BBP' then
                         Red
                 end
             ) as Trig_Staff_BBP_Red,
          max(   case
                     when Section = 'Staff'
                          and Program is null
                          and FieldName = 'BBP' then
                         Green
                 end
             ) as Trig_Staff_BBP_Green,
          max(   case
                     when Section = 'Staff'
                          and Program is null
                          and FieldName = 'BackInjuryPrevention' then
                         Red
                 end
             ) as Trig_Staff_BIP_Red,
          max(   case
                     when Section = 'Staff'
                          and Program is null
                          and FieldName = 'BackInjuryPrevention' then
                         Green
                 end
             ) as Trig_Staff_BIP_Green,
          max(   case
                     when Section = 'Staff'
                          and Program is null
                          and FieldName = 'CPR' then
                         Red
                 end
             ) as Trig_Staff_CPR_Red,
          max(   case
                     when Section = 'Staff'
                          and Program is null
                          and FieldName = 'CPR' then
                         Green
                 end
             ) as Trig_Staff_CPR_Green,
          max(   case
                     when Section = 'Staff'
                          and Program is null
                          and FieldName = 'DriversLicense' then
                         Red
                 end
             ) as Trig_Staff_DL_Red,
          max(   case
                     when Section = 'Staff'
                          and Program is null
                          and FieldName = 'DriversLicense' then
                         Green
                 end
             ) as Trig_Staff_DL_Green,
          max(   case
                     when Section = 'Staff'
                          and Program is null
                          and FieldName = 'FirstAid' then
                         Red
                 end
             ) as Trig_Staff_FA_Red,
          max(   case
                     when Section = 'Staff'
                          and Program is null
                          and FieldName = 'FirstAid' then
                         Green
                 end
             ) as Trig_Staff_FA_Green,
          max(   case
                     when Section = 'Staff'
                          and Program is null
                          and FieldName = 'SafetyCares' then
                         Red
                 end
             ) as Trig_Staff_SC_Red,
          max(   case
                     when Section = 'Staff'
                          and Program is null
                          and FieldName = 'SafetyCares' then
                         Green
                 end
             ) as Trig_Staff_SC_Green,
          max(   case
                     when Section = 'Staff'
                          and Program is null
                          and FieldName = 'TB' then
                         Red
                 end
             ) as Trig_Staff_TB_Red,
          max(   case
                     when Section = 'Staff'
                          and Program is null
                          and FieldName = 'TB' then
                         Green
                 end
             ) as Trig_Staff_TB_Green,
          max(   case
                     when Section = 'Staff'
                          and Program is null
                          and FieldName = 'WorkplaceViolence' then
                         Red
                 end
             ) as Trig_Staff_WV_Red,
          max(   case
                     when Section = 'Staff'
                          and Program is null
                          and FieldName = 'WorkplaceViolence' then
                         Green
                 end
             ) as Trig_Staff_WV_Green,
          max(   case
                     when Section = 'Staff'
                          and Program is null
                          and FieldName = 'ProfLic' then
                         Red
                 end
             ) as Trig_Staff_PL_Red,
          max(   case
                     when Section = 'Staff'
                          and Program is null
                          and FieldName = 'ProfLic' then
                         Green
                 end
             ) as Trig_Staff_PL_Green,
          max(   case
                     when Section = 'Staff'
                          and Program is null
                          and FieldName = 'MAPCert' then
                         Red
                 end
             ) as Trig_Staff_MAP_Red,
          max(   case
                     when Section = 'Staff'
                          and Program is null
                          and FieldName = 'MAPCert' then
                         Green
                 end
             ) as Trig_Staff_MAP_Green,
          max(   case
                     when Section = 'Staff'
                          and Program is null
                          and FieldName = 'EvalDueBy' then
                         Red
                 end
             ) as Trig_Staff_EVL_Red,
          max(   case
                     when Section = 'Staff'
                          and Program is null
                          and FieldName = 'EvalDueBy' then
                         Green
                 end
             ) as Trig_Staff_EVL_Green,
          max(   case
                     when Section = 'Staff'
                          and Program is null
                          and FieldName = 'LastSupervision' then
                         Red
                 end
             ) as Trig_Staff_SUP_Red,
          max(   case
                     when Section = 'Staff'
                          and Program is null
                          and FieldName = '3MoEval' then
                         Red
                 end
             ) as Trig_Staff_3MO_Red
   from catExpirationTriggers)
select e.*,

       -- Additional lookups from tblLocations (for main report VBA)
       loc.Department,
       loc.Cluster,

       -- ========================================
       -- CLIENT FIELDS (RecordType = 'Client')
       -- ========================================

       -- DateISP calculations
       case
           when e.DateISP = '1900-01-01' then
               'Missing'
           when e.DateISP = '1900-01-02' then
               'Optional'
           when e.DateISP = '1900-01-03' then
               'N/A'
           when e.DateISP is null then
               ''
           else
               convert(varchar(10), e.DateISP, 101) -- MM/dd/yyyy format
       end as DateISP_Display,
       case
           when
           (
               e.DateISP = '1900-01-01'
               or e.DateISP is null
           ) then
               'RED'
           when e.DateISP in ( '1900-01-02', '1900-01-03' ) then
               'NORMAL'
           when datediff(day, cast(getdate() as date), e.DateISP) < t.Trig_Indiv_ISP_Red then
               'RED'
           when datediff(day, cast(getdate() as date), e.DateISP) <= t.Trig_Indiv_ISP_Green then
               'GREEN'
           else
               'NORMAL'
       end as DateISP_Color,
       case
           when
           (
               e.DateISP in ( '1900-01-01', '1900-01-02', '1900-01-03' )
               or e.DateISP is null
           ) then
               0
           else
               1
       end as DateISP_ShowDate,

       -- PSDue calculations (182 days before DateISP)
       case
           when
           (
               e.DateISP in ( '1900-01-01', '1900-01-02', '1900-01-03' )
               or e.DateISP is null
           ) then
               null
           else
               dateadd(day, -182, e.DateISP)
       end as PSDue_Calculated,
       case
           when
           (
               e.DateISP in ( '1900-01-01', '1900-01-02', '1900-01-03' )
               or e.DateISP is null
           ) then
               ''
           else
               convert(varchar(10), dateadd(day, -182, e.DateISP), 101)
       end as PSDue_Display,
       case
           when
           (
               e.DateISP in ( '1900-01-01', '1900-01-02', '1900-01-03' )
               or e.DateISP is null
           ) then
               'NORMAL'
           when datediff(day, cast(getdate() as date), dateadd(day, -182, e.DateISP)) <= t.Trig_Indiv_PSDue_Green
                and cast(getdate() as date) <= dateadd(day, -182, e.DateISP) then
               'GREEN'
           when cast(getdate() as date) > dateadd(day, -182, e.DateISP) then
               'STRIKETHROUGH'
           else
               'NORMAL'
       end as PSDue_Color,
       case
           when
           (
               e.DateISP in ( '1900-01-01', '1900-01-02', '1900-01-03' )
               or e.DateISP is null
           ) then
               1
           else
               0
       end as PSDue_ShowText,

       -- DateConsentFormsSigned calculations (months based)
       case
           when e.DateConsentFormsSigned = '1900-01-01' then
               'Missing'
           when e.DateConsentFormsSigned = '1900-01-02' then
               'Optional'
           when e.DateConsentFormsSigned = '1900-01-03' then
               'N/A'
           when e.DateConsentFormsSigned is null then
               ''
           else
               convert(varchar(10), e.DateConsentFormsSigned, 101)
       end as DateConsentFormsSigned_Display,
       case
           when
           (
               e.DateConsentFormsSigned = '1900-01-01'
               or e.DateConsentFormsSigned is null
           ) then
               'RED'
           when e.DateConsentFormsSigned in ( '1900-01-02', '1900-01-03' ) then
               'NORMAL'
           when datediff(day, dateadd(month, t.Trig_Indiv_CFS_Red, e.DateConsentFormsSigned), cast(getdate() as date)) > 0 then
               'RED'
           when datediff(day, dateadd(month, t.Trig_Indiv_CFS_Green, e.DateConsentFormsSigned), cast(getdate() as date)) > 0 then
               'GREEN'
           else
               'NORMAL'
       end as DateConsentFormsSigned_Color,
       case
           when
           (
               e.DateConsentFormsSigned in ( '1900-01-01', '1900-01-02', '1900-01-03' )
               or e.DateConsentFormsSigned is null
           ) then
               0
           else
               1
       end as DateConsentFormsSigned_ShowDate,

       -- DateBMMExpires calculations
       case
           when e.DateBMMExpires = '1900-01-01' then
               'Missing'
           when e.DateBMMExpires = '1900-01-02' then
               'Optional'
           when e.DateBMMExpires = '1900-01-03' then
               'N/A'
           when e.DateBMMExpires is null then
               ''
           else
               convert(varchar(10), e.DateBMMExpires, 101)
       end as DateBMMExpires_Display,
       case
           when
           (
               e.DateBMMExpires = '1900-01-01'
               or e.DateBMMExpires is null
           ) then
               'RED'
           when e.DateBMMExpires in ( '1900-01-02', '1900-01-03' ) then
               'NORMAL'
           when datediff(day, cast(getdate() as date), e.DateBMMExpires) < t.Trig_Indiv_BMMX_Red then
               'RED'
           when datediff(day, cast(getdate() as date), e.DateBMMExpires) <= t.Trig_Indiv_BMMX_Green then
               'GREEN'
           else
               'NORMAL'
       end as DateBMMExpires_Color,
       case
           when
           (
               e.DateBMMExpires in ( '1900-01-01', '1900-01-02', '1900-01-03' )
               or e.DateBMMExpires is null
           ) then
               0
           else
               1
       end as DateBMMExpires_ShowDate,

       -- DateSPDAuthExpires calculations
       case
           when e.DateSPDAuthExpires = '1900-01-01' then
               'Missing'
           when e.DateSPDAuthExpires = '1900-01-02' then
               'Optional'
           when e.DateSPDAuthExpires = '1900-01-03' then
               'N/A'
           when e.DateSPDAuthExpires is null then
               ''
           else
               convert(varchar(10), e.DateSPDAuthExpires, 101)
       end as DateSPDAuthExpires_Display,
       case
           when
           (
               e.DateSPDAuthExpires = '1900-01-01'
               or e.DateSPDAuthExpires is null
           ) then
               'RED'
           when e.DateSPDAuthExpires in ( '1900-01-02', '1900-01-03' ) then
               'NORMAL'
           when datediff(day, cast(getdate() as date), e.DateSPDAuthExpires) < t.Trig_Indiv_SPDX_Red then
               'RED'
           when datediff(day, cast(getdate() as date), e.DateSPDAuthExpires) <= t.Trig_Indiv_SPDX_Green then
               'GREEN'
           else
               'NORMAL'
       end as DateSPDAuthExpires_Color,
       case
           when
           (
               e.DateSPDAuthExpires in ( '1900-01-01', '1900-01-02', '1900-01-03' )
               or e.DateSPDAuthExpires is null
           ) then
               0
           else
               1
       end as DateSPDAuthExpires_ShowDate,

       -- DateSignaturesDueBy calculations
       case
           when e.DateSignaturesDueBy = '1900-01-01' then
               'Missing'
           when e.DateSignaturesDueBy = '1900-01-02' then
               'Optional'
           when e.DateSignaturesDueBy = '1900-01-03' then
               'N/A'
           when e.DateSignaturesDueBy is null then
               ''
           else
               convert(varchar(10), e.DateSignaturesDueBy, 101)
       end as DateSignaturesDueBy_Display,
       case
           when
           (
               e.DateSignaturesDueBy = '1900-01-01'
               or e.DateSignaturesDueBy is null
           ) then
               'RED'
           when e.DateSignaturesDueBy in ( '1900-01-02', '1900-01-03' ) then
               'NORMAL'
           when datediff(day, cast(getdate() as date), e.DateSignaturesDueBy) < t.Trig_Indiv_SPDA_Red then
               'RED'
           when datediff(day, cast(getdate() as date), e.DateSignaturesDueBy) <= t.Trig_Indiv_SPDA_Green then
               'GREEN'
           else
               'NORMAL'
       end as DateSignaturesDueBy_Color,
       case
           when
           (
               e.DateSignaturesDueBy in ( '1900-01-01', '1900-01-02', '1900-01-03' )
               or e.DateSignaturesDueBy is null
           ) then
               0
           else
               1
       end as DateSignaturesDueBy_ShowDate,

       -- ========================================
       -- DAY & HOUSE FIELDS (RecordType = 'House')
       -- ========================================

       -- LastVehicleChecklistCompleted - Day program
       case
           when e.LastVehicleChecklistCompleted = '1900-01-01' then
               'Missing'
           when e.LastVehicleChecklistCompleted = '1900-01-02' then
               'Optional'
           when e.LastVehicleChecklistCompleted = '1900-01-03' then
               'N/A'
           when e.LastVehicleChecklistCompleted = '1900-01-05' then
               'Pending'
           when e.LastVehicleChecklistCompleted is null then
               ''
           else
               convert(varchar(10), e.LastVehicleChecklistCompleted, 101)
       end as LastVehicleChecklistCompleted_Display_Day,
       case
           when
           (
               e.LastVehicleChecklistCompleted in ( '1900-01-01', '1900-01-02', '1900-01-03', '1900-01-05' )
               or e.LastVehicleChecklistCompleted is null
           ) then
               'NORMAL'
           when datediff(day, e.LastVehicleChecklistCompleted, cast(getdate() as date)) >= t.Trig_Day_LVC_Red then
               'RED'
           else
               'NORMAL'
       end as LastVehicleChecklistCompleted_Color_Day,
       case
           when
           (
               e.LastVehicleChecklistCompleted in ( '1900-01-01', '1900-01-02', '1900-01-03', '1900-01-05' )
               or e.LastVehicleChecklistCompleted is null
           ) then
               0
           else
               1
       end as LastVehicleChecklistCompleted_ShowDate_Day,

       -- LastVehicleChecklistCompleted - Residential program
       case
           when e.LastVehicleChecklistCompleted = '1900-01-01' then
               'Missing'
           when e.LastVehicleChecklistCompleted = '1900-01-02' then
               'Optional'
           when e.LastVehicleChecklistCompleted = '1900-01-03' then
               'N/A'
           when e.LastVehicleChecklistCompleted = '1900-01-05' then
               'Pending'
           when e.LastVehicleChecklistCompleted is null then
               ''
           else
               convert(varchar(10), e.LastVehicleChecklistCompleted, 101)
       end as LastVehicleChecklistCompleted_Display_Res,
       case
           when
           (
               e.LastVehicleChecklistCompleted in ( '1900-01-01', '1900-01-02', '1900-01-03', '1900-01-05' )
               or e.LastVehicleChecklistCompleted is null
           ) then
               'NORMAL'
           when datediff(day, e.LastVehicleChecklistCompleted, cast(getdate() as date)) >= t.Trig_Res_LVC_Red then
               'RED'
           else
               'NORMAL'
       end as LastVehicleChecklistCompleted_Color_Res,
       case
           when
           (
               e.LastVehicleChecklistCompleted in ( '1900-01-01', '1900-01-02', '1900-01-03', '1900-01-05' )
               or e.LastVehicleChecklistCompleted is null
           ) then
               0
           else
               1
       end as LastVehicleChecklistCompleted_ShowDate_Res,

       -- DAYStaffTrainedInPrivacyBefore
       case
           when e.DAYStaffTrainedInPrivacyBefore = '1900-01-01' then
               'Missing'
           when e.DAYStaffTrainedInPrivacyBefore = '1900-01-02' then
               'Optional'
           when e.DAYStaffTrainedInPrivacyBefore = '1900-01-03' then
               'N/A'
           when e.DAYStaffTrainedInPrivacyBefore = '1900-01-05' then
               'Pending'
           when e.DAYStaffTrainedInPrivacyBefore is null then
               ''
           else
               convert(varchar(10), e.DAYStaffTrainedInPrivacyBefore, 101)
       end as DAYStaffTrainedInPrivacyBefore_Display,
       case
           when
           (
               e.DAYStaffTrainedInPrivacyBefore in ( '1900-01-01', '1900-01-02', '1900-01-03', '1900-01-05' )
               or e.DAYStaffTrainedInPrivacyBefore is null
           ) then
               'NORMAL'
           when datediff(day, cast(getdate() as date), e.DAYStaffTrainedInPrivacyBefore) < t.Trig_Day_STP_Red then
               'RED'
           when datediff(day, cast(getdate() as date), e.DAYStaffTrainedInPrivacyBefore) <= t.Trig_Day_STP_Green then
               'GREEN'
           else
               'NORMAL'
       end as DAYStaffTrainedInPrivacyBefore_Color,
       case
           when
           (
               e.DAYStaffTrainedInPrivacyBefore in ( '1900-01-01', '1900-01-02', '1900-01-03', '1900-01-05' )
               or e.DAYStaffTrainedInPrivacyBefore is null
           ) then
               0
           else
               1
       end as DAYStaffTrainedInPrivacyBefore_ShowDate,

       -- DAYAllPlansReviewedByStaffBefore
       case
           when e.DAYAllPlansReviewedByStaffBefore = '1900-01-01' then
               'Missing'
           when e.DAYAllPlansReviewedByStaffBefore = '1900-01-02' then
               'Optional'
           when e.DAYAllPlansReviewedByStaffBefore = '1900-01-03' then
               'N/A'
           when e.DAYAllPlansReviewedByStaffBefore = '1900-01-05' then
               'Pending'
           when e.DAYAllPlansReviewedByStaffBefore is null then
               ''
           else
               convert(varchar(10), e.DAYAllPlansReviewedByStaffBefore, 101)
       end as DAYAllPlansReviewedByStaffBefore_Display,
       case
           when
           (
               e.DAYAllPlansReviewedByStaffBefore in ( '1900-01-01', '1900-01-02', '1900-01-03', '1900-01-05' )
               or e.DAYAllPlansReviewedByStaffBefore is null
           ) then
               'NORMAL'
           when datediff(day, cast(getdate() as date), e.DAYAllPlansReviewedByStaffBefore) < t.Trig_Day_APRS_Red then
               'RED'
           when datediff(day, cast(getdate() as date), e.DAYAllPlansReviewedByStaffBefore) <= t.Trig_Day_APRS_Green then
               'GREEN'
           else
               'NORMAL'
       end as DAYAllPlansReviewedByStaffBefore_Color,
       case
           when
           (
               e.DAYAllPlansReviewedByStaffBefore in ( '1900-01-01', '1900-01-02', '1900-01-03', '1900-01-05' )
               or e.DAYAllPlansReviewedByStaffBefore is null
           ) then
               0
           else
               1
       end as DAYAllPlansReviewedByStaffBefore_ShowDate,

       -- DAYQtrlySafetyChecklistDueBy
       case
           when e.DAYQtrlySafetyChecklistDueBy = '1900-01-01' then
               'Missing'
           when e.DAYQtrlySafetyChecklistDueBy = '1900-01-02' then
               'Optional'
           when e.DAYQtrlySafetyChecklistDueBy = '1900-01-03' then
               'N/A'
           when e.DAYQtrlySafetyChecklistDueBy = '1900-01-05' then
               'Pending'
           when e.DAYQtrlySafetyChecklistDueBy is null then
               ''
           else
               convert(varchar(10), e.DAYQtrlySafetyChecklistDueBy, 101)
       end as DAYQtrlySafetyChecklistDueBy_Display,
       case
           when
           (
               e.DAYQtrlySafetyChecklistDueBy in ( '1900-01-01', '1900-01-02', '1900-01-03', '1900-01-05' )
               or e.DAYQtrlySafetyChecklistDueBy is null
           ) then
               'NORMAL'
           when datediff(day, cast(getdate() as date), e.DAYQtrlySafetyChecklistDueBy) < t.Trig_Day_QSR_Red then
               'RED'
           when datediff(day, cast(getdate() as date), e.DAYQtrlySafetyChecklistDueBy) <= t.Trig_Day_QSR_Green then
               'GREEN'
           else
               'NORMAL'
       end as DAYQtrlySafetyChecklistDueBy_Color,
       case
           when
           (
               e.DAYQtrlySafetyChecklistDueBy in ( '1900-01-01', '1900-01-02', '1900-01-03', '1900-01-05' )
               or e.DAYQtrlySafetyChecklistDueBy is null
           ) then
               0
           else
               1
       end as DAYQtrlySafetyChecklistDueBy_ShowDate,

       -- MostRecentAsleepFireDrill (Residential only, calculates based on +14 months)
       case
           when e.MostRecentAsleepFireDrill = '1900-01-01' then
               'Missing'
           when e.MostRecentAsleepFireDrill = '1900-01-02' then
               'Optional'
           when e.MostRecentAsleepFireDrill = '1900-01-03' then
               'N/A'
           when e.MostRecentAsleepFireDrill = '1900-01-05' then
               'Pending'
           when e.MostRecentAsleepFireDrill is null then
               ''
           else
               convert(varchar(10), e.MostRecentAsleepFireDrill, 101)
       end as MostRecentAsleepFireDrill_Display,
       case
           when
           (
               e.MostRecentAsleepFireDrill in ( '1900-01-01', '1900-01-02', '1900-01-03', '1900-01-05' )
               or e.MostRecentAsleepFireDrill is null
           ) then
               'NORMAL'
           when datediff(day, cast(getdate() as date), dateadd(month, 14, e.MostRecentAsleepFireDrill)) < t.Trig_Res_MRFD_Red then
               'RED'
           when datediff(day, cast(getdate() as date), dateadd(month, 14, e.MostRecentAsleepFireDrill)) <= t.Trig_Res_MRFD_Green then
               'GREEN'
           else
               'NORMAL'
       end as MostRecentAsleepFireDrill_Color,
       case
           when
           (
               e.MostRecentAsleepFireDrill in ( '1900-01-01', '1900-01-02', '1900-01-03', '1900-01-05' )
               or e.MostRecentAsleepFireDrill is null
           ) then
               0
           else
               1
       end as MostRecentAsleepFireDrill_ShowDate,

       -- HousePlansReviewedByStaffBefore (Residential)
       case
           when e.HousePlansReviewedByStaffBefore = '1900-01-01' then
               'Missing'
           when e.HousePlansReviewedByStaffBefore = '1900-01-02' then
               'Optional'
           when e.HousePlansReviewedByStaffBefore = '1900-01-03' then
               'N/A'
           when e.HousePlansReviewedByStaffBefore = '1900-01-05' then
               'Pending'
           when e.HousePlansReviewedByStaffBefore is null then
               ''
           else
               convert(varchar(10), e.HousePlansReviewedByStaffBefore, 101)
       end as HousePlansReviewedByStaffBefore_Display,
       case
           when
           (
               e.HousePlansReviewedByStaffBefore in ( '1900-01-01', '1900-01-02', '1900-01-03', '1900-01-05' )
               or e.HousePlansReviewedByStaffBefore is null
           ) then
               'NORMAL'
           when datediff(day, cast(getdate() as date), e.HousePlansReviewedByStaffBefore) < t.Trig_Res_HPR_Red then
               'RED'
           when datediff(day, cast(getdate() as date), e.HousePlansReviewedByStaffBefore) <= t.Trig_Res_HPR_Green then
               'GREEN'
           else
               'NORMAL'
       end as HousePlansReviewedByStaffBefore_Color,
       case
           when
           (
               e.HousePlansReviewedByStaffBefore in ( '1900-01-01', '1900-01-02', '1900-01-03', '1900-01-05' )
               or e.HousePlansReviewedByStaffBefore is null
           ) then
               0
           else
               1
       end as HousePlansReviewedByStaffBefore_ShowDate,

       -- HouseSafetyPlanExpires (Residential)
       case
           when e.HouseSafetyPlanExpires = '1900-01-01' then
               'Missing'
           when e.HouseSafetyPlanExpires = '1900-01-02' then
               'Optional'
           when e.HouseSafetyPlanExpires = '1900-01-03' then
               'N/A'
           when e.HouseSafetyPlanExpires = '1900-01-05' then
               'Pending'
           when e.HouseSafetyPlanExpires is null then
               ''
           else
               convert(varchar(10), e.HouseSafetyPlanExpires, 101)
       end as HouseSafetyPlanExpires_Display,
       case
           when
           (
               e.HouseSafetyPlanExpires in ( '1900-01-01', '1900-01-02', '1900-01-03', '1900-01-05' )
               or e.HouseSafetyPlanExpires is null
           ) then
               'NORMAL'
           when datediff(day, cast(getdate() as date), e.HouseSafetyPlanExpires) < t.Trig_Res_HSPE_Red then
               'RED'
           when datediff(day, cast(getdate() as date), e.HouseSafetyPlanExpires) <= t.Trig_Res_HSPE_Green then
               'GREEN'
           else
               'NORMAL'
       end as HouseSafetyPlanExpires_Color,
       case
           when
           (
               e.HouseSafetyPlanExpires in ( '1900-01-01', '1900-01-02', '1900-01-03', '1900-01-05' )
               or e.HouseSafetyPlanExpires is null
           ) then
               0
           else
               1
       end as HouseSafetyPlanExpires_ShowDate,

       -- MAPChecklistCompleted (Residential)
       case
           when e.MAPChecklistCompleted = '1900-01-01' then
               'Missing'
           when e.MAPChecklistCompleted = '1900-01-02' then
               'Optional'
           when e.MAPChecklistCompleted = '1900-01-03' then
               'N/A'
           when e.MAPChecklistCompleted = '1900-01-05' then
               'Pending'
           when e.MAPChecklistCompleted is null then
               ''
           else
               convert(varchar(10), e.MAPChecklistCompleted, 101)
       end as MAPChecklistCompleted_Display,
       case
           when
           (
               e.MAPChecklistCompleted in ( '1900-01-01', '1900-01-02', '1900-01-03', '1900-01-05' )
               or e.MAPChecklistCompleted is null
           ) then
               'NORMAL'
           when datediff(day, e.MAPChecklistCompleted, cast(getdate() as date)) >= t.Trig_Res_MAP_Red then
               'RED'
           else
               'NORMAL'
       end as MAPChecklistCompleted_Color,
       case
           when
           (
               e.MAPChecklistCompleted in ( '1900-01-01', '1900-01-02', '1900-01-03', '1900-01-05' )
               or e.MAPChecklistCompleted is null
           ) then
               0
           else
               1
       end as MAPChecklistCompleted_ShowDate,

       -- HROTrainsStaffBefore - Day program
       case
           when e.HROTrainsStaffBefore = '1900-01-01' then
               'Missing'
           when e.HROTrainsStaffBefore = '1900-01-02' then
               'Optional'
           when e.HROTrainsStaffBefore = '1900-01-03' then
               'N/A'
           when e.HROTrainsStaffBefore = '1900-01-05' then
               'Pending'
           when e.HROTrainsStaffBefore is null then
               ''
           else
               convert(varchar(10), e.HROTrainsStaffBefore, 101)
       end as HROTrainsStaffBefore_Display_Day,
       case
           when
           (
               e.HROTrainsStaffBefore in ( '1900-01-01', '1900-01-02', '1900-01-03', '1900-01-05' )
               or e.HROTrainsStaffBefore is null
           ) then
               'NORMAL'
           when datediff(day, cast(getdate() as date), e.HROTrainsStaffBefore) < t.Trig_Day_HROTS_Red then
               'RED'
           when datediff(day, cast(getdate() as date), e.HROTrainsStaffBefore) <= t.Trig_Day_HROTS_Green then
               'GREEN'
           else
               'NORMAL'
       end as HROTrainsStaffBefore_Color_Day,
       case
           when
           (
               e.HROTrainsStaffBefore in ( '1900-01-01', '1900-01-02', '1900-01-03', '1900-01-05' )
               or e.HROTrainsStaffBefore is null
           ) then
               0
           else
               1
       end as HROTrainsStaffBefore_ShowDate_Day,

       -- HROTrainsStaffBefore - Residential program
       case
           when e.HROTrainsStaffBefore = '1900-01-01' then
               'Missing'
           when e.HROTrainsStaffBefore = '1900-01-02' then
               'Optional'
           when e.HROTrainsStaffBefore = '1900-01-03' then
               'N/A'
           when e.HROTrainsStaffBefore = '1900-01-05' then
               'Pending'
           when e.HROTrainsStaffBefore is null then
               ''
           else
               convert(varchar(10), e.HROTrainsStaffBefore, 101)
       end as HROTrainsStaffBefore_Display_Res,
       case
           when
           (
               e.HROTrainsStaffBefore in ( '1900-01-01', '1900-01-02', '1900-01-03', '1900-01-05' )
               or e.HROTrainsStaffBefore is null
           ) then
               'NORMAL'
           when datediff(day, cast(getdate() as date), e.HROTrainsStaffBefore) < t.Trig_Res_HROTS_Red then
               'RED'
           when datediff(day, cast(getdate() as date), e.HROTrainsStaffBefore) <= t.Trig_Res_HROTS_Green then
               'GREEN'
           else
               'NORMAL'
       end as HROTrainsStaffBefore_Color_Res,
       case
           when
           (
               e.HROTrainsStaffBefore in ( '1900-01-01', '1900-01-02', '1900-01-03', '1900-01-05' )
               or e.HROTrainsStaffBefore is null
           ) then
               0
           else
               1
       end as HROTrainsStaffBefore_ShowDate_Res,

       -- HROTrainsIndividualsBefore - Day program
       case
           when e.HROTrainsIndividualsBefore = '1900-01-01' then
               'Missing'
           when e.HROTrainsIndividualsBefore = '1900-01-02' then
               'Optional'
           when e.HROTrainsIndividualsBefore = '1900-01-03' then
               'N/A'
           when e.HROTrainsIndividualsBefore = '1900-01-05' then
               'Pending'
           when e.HROTrainsIndividualsBefore is null then
               ''
           else
               convert(varchar(10), e.HROTrainsIndividualsBefore, 101)
       end as HROTrainsIndividualsBefore_Display_Day,
       case
           when
           (
               e.HROTrainsIndividualsBefore in ( '1900-01-01', '1900-01-02', '1900-01-03', '1900-01-05' )
               or e.HROTrainsIndividualsBefore is null
           ) then
               'NORMAL'
           when datediff(day, cast(getdate() as date), e.HROTrainsIndividualsBefore) < t.Trig_Day_HROTI_Red then
               'RED'
           when datediff(day, cast(getdate() as date), e.HROTrainsIndividualsBefore) <= t.Trig_Day_HROTI_Green then
               'GREEN'
           else
               'NORMAL'
       end as HROTrainsIndividualsBefore_Color_Day,
       case
           when
           (
               e.HROTrainsIndividualsBefore in ( '1900-01-01', '1900-01-02', '1900-01-03', '1900-01-05' )
               or e.HROTrainsIndividualsBefore is null
           ) then
               0
           else
               1
       end as HROTrainsIndividualsBefore_ShowDate_Day,

       -- HROTrainsIndividualsBefore - Residential program
       case
           when e.HROTrainsIndividualsBefore = '1900-01-01' then
               'Missing'
           when e.HROTrainsIndividualsBefore = '1900-01-02' then
               'Optional'
           when e.HROTrainsIndividualsBefore = '1900-01-03' then
               'N/A'
           when e.HROTrainsIndividualsBefore = '1900-01-05' then
               'Pending'
           when e.HROTrainsIndividualsBefore is null then
               ''
           else
               convert(varchar(10), e.HROTrainsIndividualsBefore, 101)
       end as HROTrainsIndividualsBefore_Display_Res,
       case
           when
           (
               e.HROTrainsIndividualsBefore in ( '1900-01-01', '1900-01-02', '1900-01-03', '1900-01-05' )
               or e.HROTrainsIndividualsBefore is null
           ) then
               'NORMAL'
           when datediff(day, cast(getdate() as date), e.HROTrainsIndividualsBefore) < t.Trig_Res_HROTI_Red then
               'RED'
           when datediff(day, cast(getdate() as date), e.HROTrainsIndividualsBefore) <= t.Trig_Res_HROTI_Green then
               'GREEN'
           else
               'NORMAL'
       end as HROTrainsIndividualsBefore_Color_Res,
       case
           when
           (
               e.HROTrainsIndividualsBefore in ( '1900-01-01', '1900-01-02', '1900-01-03', '1900-01-05' )
               or e.HROTrainsIndividualsBefore is null
           ) then
               0
           else
               1
       end as HROTrainsIndividualsBefore_ShowDate_Res,

       -- FSOTrainsStaffBefore - Day program
       case
           when e.FSOTrainsStaffBefore = '1900-01-01' then
               'Missing'
           when e.FSOTrainsStaffBefore = '1900-01-02' then
               'Optional'
           when e.FSOTrainsStaffBefore = '1900-01-03' then
               'N/A'
           when e.FSOTrainsStaffBefore = '1900-01-05' then
               'Pending'
           when e.FSOTrainsStaffBefore is null then
               ''
           else
               convert(varchar(10), e.FSOTrainsStaffBefore, 101)
       end as FSOTrainsStaffBefore_Display_Day,
       case
           when
           (
               e.FSOTrainsStaffBefore in ( '1900-01-01', '1900-01-02', '1900-01-03', '1900-01-05' )
               or e.FSOTrainsStaffBefore is null
           ) then
               'NORMAL'
           when datediff(day, cast(getdate() as date), e.FSOTrainsStaffBefore) < t.Trig_Day_FSOTS_Red then
               'RED'
           when datediff(day, cast(getdate() as date), e.FSOTrainsStaffBefore) <= t.Trig_Day_FSOTS_Green then
               'GREEN'
           else
               'NORMAL'
       end as FSOTrainsStaffBefore_Color_Day,
       case
           when
           (
               e.FSOTrainsStaffBefore in ( '1900-01-01', '1900-01-02', '1900-01-03', '1900-01-05' )
               or e.FSOTrainsStaffBefore is null
           ) then
               0
           else
               1
       end as FSOTrainsStaffBefore_ShowDate_Day,

       -- FSOTrainsStaffBefore - Residential program
       case
           when e.FSOTrainsStaffBefore = '1900-01-01' then
               'Missing'
           when e.FSOTrainsStaffBefore = '1900-01-02' then
               'Optional'
           when e.FSOTrainsStaffBefore = '1900-01-03' then
               'N/A'
           when e.FSOTrainsStaffBefore = '1900-01-05' then
               'Pending'
           when e.FSOTrainsStaffBefore is null then
               ''
           else
               convert(varchar(10), e.FSOTrainsStaffBefore, 101)
       end as FSOTrainsStaffBefore_Display_Res,
       case
           when
           (
               e.FSOTrainsStaffBefore in ( '1900-01-01', '1900-01-02', '1900-01-03', '1900-01-05' )
               or e.FSOTrainsStaffBefore is null
           ) then
               'NORMAL'
           when datediff(day, cast(getdate() as date), e.FSOTrainsStaffBefore) < t.Trig_Res_FSOTS_Red then
               'RED'
           when datediff(day, cast(getdate() as date), e.FSOTrainsStaffBefore) <= t.Trig_Res_FSOTS_Green then
               'GREEN'
           else
               'NORMAL'
       end as FSOTrainsStaffBefore_Color_Res,
       case
           when
           (
               e.FSOTrainsStaffBefore in ( '1900-01-01', '1900-01-02', '1900-01-03', '1900-01-05' )
               or e.FSOTrainsStaffBefore is null
           ) then
               0
           else
               1
       end as FSOTrainsStaffBefore_ShowDate_Res,

       -- FSOTrainsIndividualsBefore - Day program
       case
           when e.FSOTrainsIndividualsBefore = '1900-01-01' then
               'Missing'
           when e.FSOTrainsIndividualsBefore = '1900-01-02' then
               'Optional'
           when e.FSOTrainsIndividualsBefore = '1900-01-03' then
               'N/A'
           when e.FSOTrainsIndividualsBefore = '1900-01-05' then
               'Pending'
           when e.FSOTrainsIndividualsBefore is null then
               ''
           else
               convert(varchar(10), e.FSOTrainsIndividualsBefore, 101)
       end as FSOTrainsIndividualsBefore_Display_Day,
       case
           when
           (
               e.FSOTrainsIndividualsBefore in ( '1900-01-01', '1900-01-02', '1900-01-03', '1900-01-05' )
               or e.FSOTrainsIndividualsBefore is null
           ) then
               'NORMAL'
           when datediff(day, cast(getdate() as date), e.FSOTrainsIndividualsBefore) < t.Trig_Day_FSOTI_Red then
               'RED'
           when datediff(day, cast(getdate() as date), e.FSOTrainsIndividualsBefore) <= t.Trig_Day_FSOTI_Green then
               'GREEN'
           else
               'NORMAL'
       end as FSOTrainsIndividualsBefore_Color_Day,
       case
           when
           (
               e.FSOTrainsIndividualsBefore in ( '1900-01-01', '1900-01-02', '1900-01-03', '1900-01-05' )
               or e.FSOTrainsIndividualsBefore is null
           ) then
               0
           else
               1
       end as FSOTrainsIndividualsBefore_ShowDate_Day,

       -- FSOTrainsIndividualsBefore - Residential program
       case
           when e.FSOTrainsIndividualsBefore = '1900-01-01' then
               'Missing'
           when e.FSOTrainsIndividualsBefore = '1900-01-02' then
               'Optional'
           when e.FSOTrainsIndividualsBefore = '1900-01-03' then
               'N/A'
           when e.FSOTrainsIndividualsBefore = '1900-01-05' then
               'Pending'
           when e.FSOTrainsIndividualsBefore is null then
               ''
           else
               convert(varchar(10), e.FSOTrainsIndividualsBefore, 101)
       end as FSOTrainsIndividualsBefore_Display_Res,
       case
           when
           (
               e.FSOTrainsIndividualsBefore in ( '1900-01-01', '1900-01-02', '1900-01-03', '1900-01-05' )
               or e.FSOTrainsIndividualsBefore is null
           ) then
               'NORMAL'
           when datediff(day, cast(getdate() as date), e.FSOTrainsIndividualsBefore) < t.Trig_Res_FSOTI_Red then
               'RED'
           when datediff(day, cast(getdate() as date), e.FSOTrainsIndividualsBefore) <= t.Trig_Res_FSOTI_Green then
               'GREEN'
           else
               'NORMAL'
       end as FSOTrainsIndividualsBefore_Color_Res,
       case
           when
           (
               e.FSOTrainsIndividualsBefore in ( '1900-01-01', '1900-01-02', '1900-01-03', '1900-01-05' )
               or e.FSOTrainsIndividualsBefore is null
           ) then
               0
           else
               1
       end as FSOTrainsIndividualsBefore_ShowDate_Res,

       -- ========================================
       -- NAME FORMATTING (HumanRightsOfficer, FireSafetyOfficer)
       -- ========================================

       -- HumanRightsOfficer formatted (LastName, FirstName -> FirstName LastName)
       case
           when e.HumanRightsOfficer is null
                or ltrim(rtrim(e.HumanRightsOfficer)) = '' then
               null
           when charindex(',', e.HumanRightsOfficer) = 0 then
               null
           else
               ltrim(rtrim(substring(e.HumanRightsOfficer, charindex(',', e.HumanRightsOfficer) + 1, 255))) + ' '
               + ltrim(rtrim(substring(e.HumanRightsOfficer, 1, charindex(',', e.HumanRightsOfficer) - 1)))
       end as HumanRightsOfficer_Formatted,
       case
           when e.HumanRightsOfficer is null
                or ltrim(rtrim(e.HumanRightsOfficer)) = '' then
               1
           when charindex(',', e.HumanRightsOfficer) = 0 then
               1
           else
               0
       end as HumanRightsOfficer_IsBlank,

       -- FireSafetyOfficer formatted (LastName, FirstName -> FirstName LastName)
       case
           when e.FireSafetyOfficer is null
                or ltrim(rtrim(e.FireSafetyOfficer)) = '' then
               null
           when charindex(',', e.FireSafetyOfficer) = 0 then
               null
           else
               ltrim(rtrim(substring(e.FireSafetyOfficer, charindex(',', e.FireSafetyOfficer) + 1, 255))) + ' '
               + ltrim(rtrim(substring(e.FireSafetyOfficer, 1, charindex(',', e.FireSafetyOfficer) - 1)))
       end as FireSafetyOfficer_Formatted,
       case
           when e.FireSafetyOfficer is null
                or ltrim(rtrim(e.FireSafetyOfficer)) = '' then
               1
           when charindex(',', e.FireSafetyOfficer) = 0 then
               1
           else
               0
       end as FireSafetyOfficer_IsBlank,

       -- ========================================
       -- STAFF FIELDS (RecordType = 'Staff')
       -- ========================================
       -- BBP
       case
           when e.BBP = '1900-01-01' then
               'Missing'
           when e.BBP = '1900-01-02' then
               'Optional'
           when e.BBP = '1900-01-03' then
               'N/A'
           when e.BBP = '1900-01-05' then
               'Pending'
           when e.BBP is null then
               ''
           else
               convert(varchar(10), e.BBP, 101)
       end as BBP_Display,
       case
           when e.BBP in ( '1900-01-01' )
                or e.BBP is null then
               'RED'
           when e.BBP in ( '1900-01-02', '1900-01-03', '1900-01-05' ) then
               'NORMAL'
           when datediff(day, cast(getdate() as date), e.BBP) < t.Trig_Staff_BBP_Red then
               'RED'
           when t.Trig_Staff_BBP_Green is not null
                and datediff(day, cast(getdate() as date), e.BBP) <= t.Trig_Staff_BBP_Green then
               'GREEN'
           else
               'NORMAL'
       end as BBP_Color,
       case
           when e.BBP in ( '1900-01-01', '1900-01-02', '1900-01-03', '1900-01-05' )
                or e.BBP is null then
               0
           else
               1
       end as BBP_ShowDate,
       -- BackInjuryPrevention
       case
           when e.BackInjuryPrevention = '1900-01-01' then
               'Missing'
           when e.BackInjuryPrevention = '1900-01-02' then
               'Optional'
           when e.BackInjuryPrevention = '1900-01-03' then
               'N/A'
           when e.BackInjuryPrevention = '1900-01-05' then
               'Pending'
           when e.BackInjuryPrevention is null then
               ''
           else
               convert(varchar(10), e.BackInjuryPrevention, 101)
       end as BackInjuryPrevention_Display,
       case
           when e.BackInjuryPrevention in ( '1900-01-01' )
                or e.BackInjuryPrevention is null then
               'RED'
           when e.BackInjuryPrevention in ( '1900-01-02', '1900-01-03', '1900-01-05' ) then
               'NORMAL'
           when datediff(day, cast(getdate() as date), e.BackInjuryPrevention) < t.Trig_Staff_BIP_Red then
               'RED'
           when t.Trig_Staff_BIP_Green is not null
                and datediff(day, cast(getdate() as date), e.BackInjuryPrevention) <= t.Trig_Staff_BIP_Green then
               'GREEN'
           else
               'NORMAL'
       end as BackInjuryPrevention_Color,
       case
           when e.BackInjuryPrevention in ( '1900-01-01', '1900-01-02', '1900-01-03', '1900-01-05' )
                or e.BackInjuryPrevention is null then
               0
           else
               1
       end as BackInjuryPrevention_ShowDate,
       -- CPR
       case
           when e.CPR = '1900-01-01' then
               'Missing'
           when e.CPR = '1900-01-02' then
               'Optional'
           when e.CPR = '1900-01-03' then
               'N/A'
           when e.CPR = '1900-01-05' then
               'Pending'
           when e.CPR is null then
               ''
           else
               convert(varchar(10), e.CPR, 101)
       end as CPR_Display,
       case
           when e.CPR in ( '1900-01-01' )
                or e.CPR is null then
               'RED'
           when e.CPR in ( '1900-01-02', '1900-01-03', '1900-01-05' ) then
               'NORMAL'
           when datediff(day, cast(getdate() as date), e.CPR) < t.Trig_Staff_CPR_Red then
               'RED'
           when t.Trig_Staff_CPR_Green is not null
                and datediff(day, cast(getdate() as date), e.CPR) <= t.Trig_Staff_CPR_Green then
               'GREEN'
           else
               'NORMAL'
       end as CPR_Color,
       case
           when e.CPR in ( '1900-01-01', '1900-01-02', '1900-01-03', '1900-01-05' )
                or e.CPR is null then
               0
           else
               1
       end as CPR_ShowDate,
       -- DefensiveDriving (no Red/Green in triggers; show date or "Done"/text)
       case
           when e.DefensiveDriving = '1900-01-01' then
               'Missing'
           when e.DefensiveDriving = '1900-01-02' then
               'Optional'
           when e.DefensiveDriving = '1900-01-03' then
               'N/A'
           when e.DefensiveDriving = '1900-01-04' then
               'Done'
           when e.DefensiveDriving = '1900-01-05' then
               'Pending'
           when e.DefensiveDriving is null then
               ''
           else
               convert(varchar(10), e.DefensiveDriving, 101)
       end as DefensiveDriving_Display,
       'NORMAL' as DefensiveDriving_Color,
       case
           when e.DefensiveDriving in ( '1900-01-01', '1900-01-02', '1900-01-03', '1900-01-05' )
                or e.DefensiveDriving is null then
               0
           else
               1
       end as DefensiveDriving_ShowDate,
       -- DriversLicense
       case
           when e.DriversLicense = '1900-01-01' then
               'Missing'
           when e.DriversLicense = '1900-01-02' then
               'Optional'
           when e.DriversLicense = '1900-01-03' then
               'N/A'
           when e.DriversLicense = '1900-01-05' then
               'Pending'
           when e.DriversLicense is null then
               ''
           else
               convert(varchar(10), e.DriversLicense, 101)
       end as DriversLicense_Display,
       case
           when e.DriversLicense in ( '1900-01-01' )
                or e.DriversLicense is null then
               'RED'
           when e.DriversLicense in ( '1900-01-02', '1900-01-03', '1900-01-05' ) then
               'NORMAL'
           when datediff(day, cast(getdate() as date), e.DriversLicense) < t.Trig_Staff_DL_Red then
               'RED'
           when t.Trig_Staff_DL_Green is not null
                and datediff(day, cast(getdate() as date), e.DriversLicense) <= t.Trig_Staff_DL_Green then
               'GREEN'
           else
               'NORMAL'
       end as DriversLicense_Color,
       case
           when e.DriversLicense in ( '1900-01-01', '1900-01-02', '1900-01-03', '1900-01-05' )
                or e.DriversLicense is null then
               0
           else
               1
       end as DriversLicense_ShowDate,
       -- FirstAid
       case
           when e.FirstAid = '1900-01-01' then
               'Missing'
           when e.FirstAid = '1900-01-02' then
               'Optional'
           when e.FirstAid = '1900-01-03' then
               'N/A'
           when e.FirstAid = '1900-01-05' then
               'Pending'
           when e.FirstAid is null then
               ''
           else
               convert(varchar(10), e.FirstAid, 101)
       end as FirstAid_Display,
       case
           when e.FirstAid in ( '1900-01-01' )
                or e.FirstAid is null then
               'RED'
           when e.FirstAid in ( '1900-01-02', '1900-01-03', '1900-01-05' ) then
               'NORMAL'
           when datediff(day, cast(getdate() as date), e.FirstAid) < t.Trig_Staff_FA_Red then
               'RED'
           when t.Trig_Staff_FA_Green is not null
                and datediff(day, cast(getdate() as date), e.FirstAid) <= t.Trig_Staff_FA_Green then
               'GREEN'
           else
               'NORMAL'
       end as FirstAid_Color,
       case
           when e.FirstAid in ( '1900-01-01', '1900-01-02', '1900-01-03', '1900-01-05' )
                or e.FirstAid is null then
               0
           else
               1
       end as FirstAid_ShowDate,
       -- PBS
       case
           when e.PBS = '1900-01-01' then
               'Missing'
           when e.PBS = '1900-01-02' then
               'Optional'
           when e.PBS = '1900-01-03' then
               'N/A'
           when e.PBS = '1900-01-04' then
               'Done'
           when e.PBS = '1900-01-05' then
               'Pending'
           when e.PBS is null then
               ''
           else
               convert(varchar(10), e.PBS, 101)
       end as PBS_Display,
       'NORMAL' as PBS_Color,
       case
           when e.PBS in ( '1900-01-01', '1900-01-02', '1900-01-03', '1900-01-05' )
                or e.PBS is null then
               0
           else
               1
       end as PBS_ShowDate,
       -- SafetyCares
       case
           when e.SafetyCares = '1900-01-01' then
               'Missing'
           when e.SafetyCares = '1900-01-02' then
               'Optional'
           when e.SafetyCares = '1900-01-03' then
               'N/A'
           when e.SafetyCares = '1900-01-05' then
               'Pending'
           when e.SafetyCares is null then
               ''
           else
               convert(varchar(10), e.SafetyCares, 101)
       end as SafetyCares_Display,
       case
           when e.SafetyCares in ( '1900-01-01' )
                or e.SafetyCares is null then
               'RED'
           when e.SafetyCares in ( '1900-01-02', '1900-01-03', '1900-01-05' ) then
               'NORMAL'
           when datediff(day, cast(getdate() as date), e.SafetyCares) < t.Trig_Staff_SC_Red then
               'RED'
           when t.Trig_Staff_SC_Green is not null
                and datediff(day, cast(getdate() as date), e.SafetyCares) <= t.Trig_Staff_SC_Green then
               'GREEN'
           else
               'NORMAL'
       end as SafetyCares_Color,
       case
           when e.SafetyCares in ( '1900-01-01', '1900-01-02', '1900-01-03', '1900-01-05' )
                or e.SafetyCares is null then
               0
           else
               1
       end as SafetyCares_ShowDate,
       -- TB
       case
           when e.TB = '1900-01-01' then
               'Missing'
           when e.TB = '1900-01-02' then
               'Optional'
           when e.TB = '1900-01-03' then
               'N/A'
           when e.TB = '1900-01-04' then
               'Completed'
           when e.TB = '1900-01-05' then
               'Pending'
           when e.TB is null then
               ''
           else
               convert(varchar(10), e.TB, 101)
       end as TB_Display,
       case
           when e.TB in ( '1900-01-01' )
                or e.TB is null then
               'RED'
           when e.TB in ( '1900-01-02', '1900-01-03', '1900-01-04', '1900-01-05' ) then
               'NORMAL'
           when datediff(day, cast(getdate() as date), e.TB) < t.Trig_Staff_TB_Red then
               'RED'
           when t.Trig_Staff_TB_Green is not null
                and datediff(day, cast(getdate() as date), e.TB) <= t.Trig_Staff_TB_Green then
               'GREEN'
           else
               'NORMAL'
       end as TB_Color,
       case
           when e.TB in ( '1900-01-01', '1900-01-02', '1900-01-03', '1900-01-04', '1900-01-05' )
                or e.TB is null then
               0
           else
               1
       end as TB_ShowDate,
       -- WheelchairSafety
       case
           when e.WheelchairSafety = '1900-01-01' then
               'Missing'
           when e.WheelchairSafety = '1900-01-02' then
               'Optional'
           when e.WheelchairSafety = '1900-01-03' then
               'N/A'
           when e.WheelchairSafety = '1900-01-04' then
               'Done'
           when e.WheelchairSafety = '1900-01-05' then
               'Pending'
           when e.WheelchairSafety is null then
               ''
           else
               convert(varchar(10), e.WheelchairSafety, 101)
       end as WheelchairSafety_Display,
       'NORMAL' as WheelchairSafety_Color,
       case
           when e.WheelchairSafety in ( '1900-01-01', '1900-01-02', '1900-01-03', '1900-01-05' )
                or e.WheelchairSafety is null then
               0
           else
               1
       end as WheelchairSafety_ShowDate,
       -- WorkplaceViolence
       case
           when e.WorkplaceViolence = '1900-01-01' then
               'Missing'
           when e.WorkplaceViolence = '1900-01-02' then
               'Optional'
           when e.WorkplaceViolence = '1900-01-03' then
               'N/A'
           when e.WorkplaceViolence = '1900-01-05' then
               'Pending'
           when e.WorkplaceViolence is null then
               ''
           else
               convert(varchar(10), e.WorkplaceViolence, 101)
       end as WorkplaceViolence_Display,
       case
           when e.WorkplaceViolence in ( '1900-01-01' )
                or e.WorkplaceViolence is null then
               'RED'
           when e.WorkplaceViolence in ( '1900-01-02', '1900-01-03', '1900-01-05' ) then
               'NORMAL'
           when datediff(day, cast(getdate() as date), e.WorkplaceViolence) < t.Trig_Staff_WV_Red then
               'RED'
           when t.Trig_Staff_WV_Green is not null
                and datediff(day, cast(getdate() as date), e.WorkplaceViolence) <= t.Trig_Staff_WV_Green then
               'GREEN'
           else
               'NORMAL'
       end as WorkplaceViolence_Color,
       case
           when e.WorkplaceViolence in ( '1900-01-01', '1900-01-02', '1900-01-03', '1900-01-05' )
                or e.WorkplaceViolence is null then
               0
           else
               1
       end as WorkplaceViolence_ShowDate,
       -- ProfessionalLicenses (trigger table uses ProfLic)
       case
           when e.ProfessionalLicenses = '1900-01-01' then
               'Missing'
           when e.ProfessionalLicenses = '1900-01-02' then
               'Optional'
           when e.ProfessionalLicenses = '1900-01-03' then
               'N/A'
           when e.ProfessionalLicenses = '1900-01-05' then
               'Pending'
           when e.ProfessionalLicenses is null then
               ''
           else
               convert(varchar(10), e.ProfessionalLicenses, 101)
       end as ProfessionalLicenses_Display,
       case
           when e.ProfessionalLicenses in ( '1900-01-01' )
                or e.ProfessionalLicenses is null then
               'RED'
           when e.ProfessionalLicenses in ( '1900-01-02', '1900-01-03', '1900-01-05' ) then
               'NORMAL'
           when datediff(day, cast(getdate() as date), e.ProfessionalLicenses) < t.Trig_Staff_PL_Red then
               'RED'
           when t.Trig_Staff_PL_Green is not null
                and datediff(day, cast(getdate() as date), e.ProfessionalLicenses) <= t.Trig_Staff_PL_Green then
               'GREEN'
           else
               'NORMAL'
       end as ProfessionalLicenses_Color,
       case
           when e.ProfessionalLicenses in ( '1900-01-01', '1900-01-02', '1900-01-03', '1900-01-05' )
                or e.ProfessionalLicenses is null then
               0
           else
               1
       end as ProfessionalLicenses_ShowDate,
       -- MAPCert
       case
           when e.MAPCert = '1900-01-01' then
               'Missing'
           when e.MAPCert = '1900-01-02' then
               'Optional'
           when e.MAPCert = '1900-01-03' then
               'N/A'
           when e.MAPCert = '1900-01-05' then
               'Pending'
           when e.MAPCert is null then
               ''
           else
               convert(varchar(10), e.MAPCert, 101)
       end as MAPCert_Display,
       case
           when e.MAPCert in ( '1900-01-01' )
                or e.MAPCert is null then
               'RED'
           when e.MAPCert in ( '1900-01-02', '1900-01-03', '1900-01-05' ) then
               'NORMAL'
           when datediff(day, cast(getdate() as date), e.MAPCert) < t.Trig_Staff_MAP_Red then
               'RED'
           when t.Trig_Staff_MAP_Green is not null
                and datediff(day, cast(getdate() as date), e.MAPCert) <= t.Trig_Staff_MAP_Green then
               'GREEN'
           else
               'NORMAL'
       end as MAPCert_Color,
       case
           when e.MAPCert in ( '1900-01-01', '1900-01-02', '1900-01-03', '1900-01-05' )
                or e.MAPCert is null then
               0
           else
               1
       end as MAPCert_ShowDate,
       -- EvalDueBy
       case
           when e.EvalDueBy = '1900-01-01' then
               'Missing'
           when e.EvalDueBy = '1900-01-02' then
               'Optional'
           when e.EvalDueBy = '1900-01-03' then
               'N/A'
           when e.EvalDueBy = '1900-01-05' then
               'Pending'
           when e.EvalDueBy is null then
               ''
           else
               convert(varchar(10), e.EvalDueBy, 101)
       end as EvalDueBy_Display,
       case
           when e.EvalDueBy in ( '1900-01-01' )
                or e.EvalDueBy is null then
               'RED'
           when e.EvalDueBy in ( '1900-01-02', '1900-01-03', '1900-01-05' ) then
               'NORMAL'
           when datediff(day, cast(getdate() as date), e.EvalDueBy) < t.Trig_Staff_EVL_Red then
               'RED'
           when t.Trig_Staff_EVL_Green is not null
                and datediff(day, cast(getdate() as date), e.EvalDueBy) <= t.Trig_Staff_EVL_Green then
               'GREEN'
           else
               'NORMAL'
       end as EvalDueBy_Color,
       case
           when e.EvalDueBy in ( '1900-01-01', '1900-01-02', '1900-01-03', '1900-01-05' )
                or e.EvalDueBy is null then
               0
           else
               1
       end as EvalDueBy_ShowDate,
       -- LastSupervision
       case
           when e.LastSupervision = '1900-01-01' then
               'Missing'
           when e.LastSupervision = '1900-01-02' then
               'Optional'
           when e.LastSupervision = '1900-01-03' then
               'N/A'
           when e.LastSupervision is null then
               ''
           else
               convert(varchar(10), e.LastSupervision, 101)
       end as LastSupervision_Display,
       case
           when e.LastSupervision is null then
               'NORMAL'
           when datediff(day, e.LastSupervision, cast(getdate() as date)) > t.Trig_Staff_SUP_Red then
               'RED'
           else
               'NORMAL'
       end as LastSupervision_Color,
       case
           when e.LastSupervision in ( '1900-01-01', '1900-01-02', '1900-01-03' )
                or e.LastSupervision is null then
               0
           else
               1
       end as LastSupervision_ShowDate
from tblExpirations e (nolock)
    cross join TriggerValues t (nolock)
    left join tblLocations loc (nolock)
        on e.Location = loc.GPName
where e.RecordType in ( 'Client', 'Staff', 'House' );

go


