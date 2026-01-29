' Module Name: GlobalVariables
' Module Type: Standard Module
' Lines of Code: 81
' Extracted: 1/29/2026 4:12:28 PM

Option Compare Database
Option Explicit
Public Const MAX_PATH As Long = 260, NAME_NOT_IN_COLLECTION = 3265
Public Const ExpMissing As String = "1900-01-01", ExpOptional As String = "1900-01-02", ExpNA As String = "1900-01-03", ExpCompleted = "1900-01-04", ExpPending = "1900-01-05"

Public ExportFileName As String, Loc As String, RetValue As Variant, CommandLine As String, SysCmdResult As Variant, TILLDataBase As Database, TILLDBErrorMessage As Variant
Public RememberPreviousBackColor As Variant, AddFamilyMemberToClient As Boolean, PeopleFormClosed As Boolean
Public GlobalAddress As Variant, GlobalCity As Variant, GlobalState As Variant, GlobalZIP As Variant, GlobalValidated As Variant, GlobalCongressionalDistrict As Variant
Public ExpMissingCalculated As Date, ExpOptionalCalculated As Date, ExpNACalculated As Date, InvalidFamilyCountMessageDisplay As Boolean, ResetPeopleRecordButtons As Boolean
Public DontRunExpirations As Boolean, WorkingOn As String, EmailNotificationPassword As String, AddressValidationFailed As Boolean
Public BenefitsManagerName As String, BenefitsManagerEmail As String

Public Trig_Staff_BBP_Red As Integer, Trig_Staff_BBP_Green As Integer, Trig_Staff_CPR_Red As Integer, Trig_Staff_CPR_Green As Integer
Public Trig_Staff_DL_Red  As Integer, Trig_Staff_DL_Green  As Integer, Trig_Staff_FA_Red  As Integer, Trig_Staff_FA_Green  As Integer
Public Trig_Staff_SC_Red  As Integer, Trig_Staff_SC_Green  As Integer, Trig_Staff_TB_Red  As Integer, Trig_Staff_TB_Green  As Integer
Public Trig_Staff_WV_Red  As Integer, Trig_Staff_WV_Green  As Integer, Trig_Staff_MAP_Red As Integer, Trig_Staff_MAP_Green As Integer
Public Trig_Staff_EVL_Red As Integer, Trig_Staff_EVL_Green As Integer, Trig_Staff_3MO_Red As Integer, Trig_Staff_3MO_Green As Integer
Public Trig_Staff_SUP_Red As Integer, Trig_Staff_SUP_Green As Integer, Trig_Staff_PL_Red  As Integer, Trig_Staff_PL_Green  As Integer
Public Trig_Staff_BIP_Red As Integer, Trig_Staff_BIP_Green As Integer

Public Trig_Res_LVC_Red   As Integer, Trig_Res_LVC_Green   As Integer, Trig_Res_MRFD_Red  As Integer, Trig_Res_MRFD_Green  As Integer
Public Trig_Res_HPR_Red   As Integer, Trig_Res_HPR_Green   As Integer, Trig_Res_HSPE_Red  As Integer, Trig_Res_HSPE_Green  As Integer
Public Trig_Res_MAP_Red   As Integer, Trig_Res_MAP_Green   As Integer, Trig_Res_HROTS_Red As Integer, Trig_Res_HROTS_Green As Integer
Public Trig_Res_HROTI_Red As Integer, Trig_Res_HROTI_Green As Integer, Trig_Res_FSOTS_Red As Integer, Trig_Res_FSOTS_Green As Integer
Public Trig_Res_FSOTI_Red As Integer, Trig_Res_FSOTI_Green As Integer

Public Trig_Day_LVC_Red   As Integer, Trig_Day_LVC_Green   As Integer, Trig_Day_STP_Red   As Integer, Trig_Day_STP_Green   As Integer
Public Trig_Day_APRS_Red  As Integer, Trig_Day_APRS_Green  As Integer, Trig_Day_QSR_Red   As Integer, Trig_Day_QSR_Green   As Integer
Public Trig_Day_HROTS_Red As Integer, Trig_Day_HROTS_Green As Integer, Trig_Day_HROTI_Red As Integer, Trig_Day_HROTI_Green As Integer
Public Trig_Day_FSOTS_Red As Integer, Trig_Day_FSOTS_Green As Integer, Trig_Day_FSOTI_Red As Integer, Trig_Day_FSOTI_Green As Integer

Public Trig_Indiv_ISP_Red  As Integer, Trig_Indiv_ISP_Green  As Integer, Trig_Indiv_PSDue_Green  As Integer, Trig_Indiv_CFS_Red    As Integer, Trig_Indiv_CFS_Green  As Integer
Public Trig_Indiv_BMMX_Red As Integer, Trig_Indiv_BMMX_Green As Integer, Trig_Indiv_BMMH_Red     As Integer, Trig_Indiv_BMMH_Green As Integer
Public Trig_Indiv_BMMS_Red As Integer, Trig_Indiv_BMMS_Green As Integer, Trig_Indiv_SPDX_Red     As Integer, Trig_Indiv_SPDX_Green As Integer
Public Trig_Indiv_SPDH_Red As Integer, Trig_Indiv_SPDH_Green As Integer, Trig_Indiv_SPDS_Red     As Integer, Trig_Indiv_SPDS_Green As Integer
Public Trig_Indiv_SPDA_Red As Integer, Trig_Indiv_SPDA_Green As Integer

'Public Declare Function FindExecutable Lib "shell32" Alias "FindExecutableA" (ByVal lpFile As String, ByVal lpDirectory As String, ByVal sResult As String) As Long

Public Sub InitializeTriggers()
    Trig_Staff_BBP_Red = DLookup("Red", "catExpirationTriggers", "FieldName='BBP'"): Trig_Staff_BBP_Green = DLookup("Green", "catExpirationTriggers", "FieldName='BBP'")
    Trig_Staff_CPR_Red = DLookup("Red", "catExpirationTriggers", "FieldName='CPR'"): Trig_Staff_CPR_Green = DLookup("Green", "catExpirationTriggers", "FieldName='CPR'")
    Trig_Staff_DL_Red = DLookup("Red", "catExpirationTriggers", "FieldName='DriversLicense'"): Trig_Staff_DL_Green = DLookup("Green", "catExpirationTriggers", "FieldName='DriversLicense'")
    Trig_Staff_FA_Red = DLookup("Red", "catExpirationTriggers", "FieldName='FirstAid'"): Trig_Staff_FA_Green = DLookup("Green", "catExpirationTriggers", "FieldName='FirstAid'")
    Trig_Staff_SC_Red = DLookup("Red", "catExpirationTriggers", "FieldName='SafetyCares'"): Trig_Staff_SC_Green = DLookup("Green", "catExpirationTriggers", "FieldName='SafetyCares'")
    Trig_Staff_TB_Red = DLookup("Red", "catExpirationTriggers", "FieldName='TB'"): Trig_Staff_TB_Green = DLookup("Green", "catExpirationTriggers", "FieldName='TB'")
    Trig_Staff_WV_Red = DLookup("Red", "catExpirationTriggers", "FieldName='WorkplaceViolence'"): Trig_Staff_WV_Green = DLookup("Green", "catExpirationTriggers", "FieldName='WorkplaceViolence'")
    Trig_Staff_BIP_Red = DLookup("Red", "catExpirationTriggers", "FieldName='BackInjuryPrevention'"): Trig_Staff_BIP_Green = DLookup("Green", "catExpirationTriggers", "FieldName='BackInjuryPrevention'")
    Trig_Staff_MAP_Red = DLookup("Red", "catExpirationTriggers", "FieldName='MAPCert'"): Trig_Staff_MAP_Green = DLookup("Green", "catExpirationTriggers", "FieldName='MAPCert'")
    Trig_Staff_EVL_Red = DLookup("Red", "catExpirationTriggers", "FieldName='EvalDueBy'"): Trig_Staff_EVL_Green = DLookup("Green", "catExpirationTriggers", "FieldName='EvalDueBy'")
    Trig_Staff_PL_Red = DLookup("Red", "catExpirationTriggers", "FieldName='ProfLic'"): Trig_Staff_PL_Green = DLookup("Green", "catExpirationTriggers", "FieldName='ProfLic'")
    Trig_Staff_3MO_Red = DLookup("Red", "catExpirationTriggers", "FieldName='3MoEval'"): Trig_Staff_3MO_Green = 0
    Trig_Staff_SUP_Red = DLookup("Red", "catExpirationTriggers", "FieldName='LastSupervision'"): Trig_Staff_SUP_Green = 0
    Trig_Res_LVC_Red = DLookup("Red", "catExpirationTriggers", "Program='Res' AND FieldName='LastVehicleChecklistCompleted'"): Trig_Res_LVC_Green = 0
    Trig_Res_MRFD_Red = DLookup("Red", "catExpirationTriggers", "Program='Res' AND FieldName='MostRecentAsleepFireDrill'"): Trig_Res_MRFD_Green = DLookup("Green", "catExpirationTriggers", "Program='Res' AND FieldName='MostRecentAsleepFireDrill'")
    Trig_Res_HPR_Red = DLookup("Red", "catExpirationTriggers", "Program='Res' AND FieldName='HousePlansReviewedByStaffBefore'"): Trig_Res_HPR_Green = DLookup("Green", "catExpirationTriggers", "Program='Res' AND FieldName='HousePlansReviewedByStaffBefore'")
    Trig_Res_HSPE_Red = DLookup("Red", "catExpirationTriggers", "Program='Res' AND FieldName='HouseSafetyPlanExpires'"): Trig_Res_HSPE_Green = DLookup("Green", "catExpirationTriggers", "Program='Res' AND FieldName='HouseSafetyPlanExpires'")
    Trig_Res_MAP_Red = DLookup("Red", "catExpirationTriggers", "Program='Res' AND FieldName='MAPChecklistCompleted'"): Trig_Res_MAP_Green = 0
    Trig_Res_HROTS_Red = DLookup("Red", "catExpirationTriggers", "Program='Res' AND FieldName='HROTrainsStaffBefore'"): Trig_Res_HROTS_Green = DLookup("Green", "catExpirationTriggers", "Program='Res' AND FieldName='HROTrainsStaffBefore'")
    Trig_Res_HROTI_Red = DLookup("Red", "catExpirationTriggers", "Program='Res' AND FieldName='HROTrainsIndividualsBefore'"): Trig_Res_HROTI_Green = DLookup("Green", "catExpirationTriggers", "Program='Res' AND FieldName='HROTrainsIndividualsBefore'")
    Trig_Res_FSOTS_Red = DLookup("Red", "catExpirationTriggers", "Program='Res' AND FieldName='FSOTrainsStaffBefore'"): Trig_Res_FSOTS_Green = DLookup("Green", "catExpirationTriggers", "Program='Res' AND FieldName='FSOTrainsStaffBefore'")
    Trig_Res_FSOTI_Red = DLookup("Red", "catExpirationTriggers", "Program='Res' AND FieldName='FSOTrainsIndividualsBefore'"): Trig_Res_FSOTI_Green = DLookup("Green", "catExpirationTriggers", "Program='Res' AND FieldName='FSOTrainsIndividualsBefore'")
    Trig_Day_LVC_Red = DLookup("Red", "catExpirationTriggers", "Program='Day' AND FieldName='LastVehicleChecklistCompleted'"): Trig_Day_LVC_Green = 0
    Trig_Day_STP_Red = DLookup("Red", "catExpirationTriggers", "Program='Day' AND FieldName='DAYStaffTrainedInPrivacyBefore'"): Trig_Day_STP_Green = DLookup("Green", "catExpirationTriggers", "Program='Day' AND FieldName='DAYStaffTrainedInPrivacyBefore'")
    Trig_Day_APRS_Red = DLookup("Red", "catExpirationTriggers", "Program='Day' AND FieldName='DAYAllPlansReviewedByStaffBefore'"): Trig_Day_APRS_Green = DLookup("Green", "catExpirationTriggers", "Program='Day' AND FieldName='DAYAllPlansReviewedByStaffBefore'")
    Trig_Day_QSR_Red = DLookup("Red", "catExpirationTriggers", "Program='Day' AND FieldName='DAYQtrlySafetyChecklistDueBy'"): Trig_Day_QSR_Green = DLookup("Green", "catExpirationTriggers", "Program='Day' AND FieldName='DAYQtrlySafetyChecklistDueBy'")
    Trig_Day_HROTS_Red = DLookup("Red", "catExpirationTriggers", "Program='Day' AND FieldName='HROTrainsStaffBefore'"): Trig_Day_HROTS_Green = DLookup("Green", "catExpirationTriggers", "Program='Day' AND FieldName='HROTrainsStaffBefore'")
    Trig_Day_HROTI_Red = DLookup("Red", "catExpirationTriggers", "Program='Day' AND FieldName='HROTrainsIndividualsBefore'"): Trig_Day_HROTI_Green = DLookup("Green", "catExpirationTriggers", "Program='Day' AND FieldName='HROTrainsIndividualsBefore'")
    Trig_Day_FSOTS_Red = DLookup("Red", "catExpirationTriggers", "Program='Day' AND FieldName='FSOTrainsStaffBefore'"): Trig_Day_FSOTS_Green = DLookup("Green", "catExpirationTriggers", "Program='Day' AND FieldName='FSOTrainsStaffBefore'")
    Trig_Day_FSOTI_Red = DLookup("Red", "catExpirationTriggers", "Program='Day' AND FieldName='FSOTrainsIndividualsBefore'"): Trig_Day_FSOTI_Green = DLookup("Green", "catExpirationTriggers", "Program='Day' AND FieldName='FSOTrainsIndividualsBefore'")
    Trig_Indiv_ISP_Red = DLookup("Red", "catExpirationTriggers", "Section='Individuals' AND FieldName='DateISP'"): Trig_Indiv_ISP_Green = DLookup("Green", "catExpirationTriggers", "Section='Individuals' AND FieldName='DateISP'")
    Trig_Indiv_PSDue_Green = DLookup("Green", "catExpirationTriggers", "Section='Individuals' AND FieldName='PSDue'")
    Trig_Indiv_CFS_Red = DLookup("Red", "catExpirationTriggers", "Section='Individuals' AND FieldName='DateConsentFormsSigned'"): Trig_Indiv_CFS_Green = DLookup("Green", "catExpirationTriggers", "Section='Individuals' AND FieldName='DateConsentFormsSigned'")
    Trig_Indiv_BMMX_Red = DLookup("Red", "catExpirationTriggers", "Section='Individuals' AND FieldName='DateBMMExpires'"): Trig_Indiv_BMMX_Green = DLookup("Green", "catExpirationTriggers", "Section='Individuals' AND FieldName='DateBMMExpires'")
    Trig_Indiv_BMMH_Red = DLookup("Red", "catExpirationTriggers", "Section='Individuals' AND FieldName='DateBMMAccessSignedHRC'"): Trig_Indiv_BMMH_Green = 0
    Trig_Indiv_BMMS_Red = DLookup("Red", "catExpirationTriggers", "Section='Individuals' AND FieldName='DateBMMAccessSigned'"): Trig_Indiv_BMMS_Green = 0
    Trig_Indiv_SPDX_Red = DLookup("Red", "catExpirationTriggers", "Section='Individuals' AND FieldName='DateSPDAuthExpires'"): Trig_Indiv_SPDX_Green = DLookup("Green", "catExpirationTriggers", "Section='Individuals' AND FieldName='DateSPDAuthExpires'")
    Trig_Indiv_SPDA_Red = DLookup("Red", "catExpirationTriggers", "Section='Individuals' AND FieldName='DateSignaturesDueBy'"): Trig_Indiv_SPDA_Green = DLookup("Green", "catExpirationTriggers", "Section='Individuals' AND FieldName='DateSignaturesDueBy'")
    ExpMissingCalculated = DateValue("1900-01-01"): ExpOptionalCalculated = DateValue("1900-01-02"): ExpNACalculated = DateValue("1900-01-03")
End Sub
