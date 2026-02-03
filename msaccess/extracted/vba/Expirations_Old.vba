Option Compare Database
Option Explicit

Public Function RunExpirationsReport(ExpDatesReportInitiatedFromReportsMenu As Boolean) As Boolean
On Error GoTo 0
'On Error GoTo ShowMeError
    Dim Criteria As Variant, TempDB As Database, ExportFileName As String
    
    Set TempDB = CurrentDb
    TempDB.QueryTimeout = 0

    Form_frmRpt.ProgressMessages = "": Form_frmRpt.ProgressMessages.Requery
    
    If DontRunExpirations Then
        MsgBox "There are no records in the Staff or Staff Skills tables." & vbCrLf & "SQL Server Refresh Skills job likely failed." & vbCrLf & _
               "You cannot run Expiration Dates report or Staff Evals and Supervisions until fixed." & vbCrLf & "Contact Tech Services to resolve.", vbOKOnly, "Error!"
        RunExpirationsReport = False
        Exit Function
    End If
'   First, generate a listing of staff with no skills records.
    Call AppendProgressMessages("Generating list of staff without skills.")
    Call ExecReport("rptSTAFFWITHNOSKILLS")
    
'   Initialize for Expiration Dates.
    RunExpirationsReport = True
    Call DropTempTables
    DoCmd.SetWarnings False

'   Create a temporary table for staff.
    Call AppendProgressMessages("Creating temporary table for staff.")
    Call ExpirationsHousekeeping(TempDB, 1)
'   DoCmd.OpenQuery "qryEXPIRATIONS01"
    TempDB.Execute "SELECT tblStaff.* INTO tempstaff FROM tblStaff ORDER BY tblStaff.LASTNAME, tblStaff.FRSTNAME;", dbSeeChanges: Call BriefDelay: Call BriefDelay
'   DoCmd.OpenQuery "qryEXPIRATIONS02"
    TempDB.Execute "UPDATE qrytblStaffDedhamManagers INNER JOIN tempstaff ON qrytblStaffDedhamManagers.SUPERVISORCODE_I = tempstaff.SUPERVISORCODE_I " & _
        "SET tempstaff.DIVISIONCODE_I = 'DEDHAM', tempstaff.DEPRTMNT = qrytblStaffDedhamManagers.NewLocation;", dbSeeChanges: Call BriefDelay: Call BriefDelay
'   DoCmd.OpenQuery "qryEXPIRATIONS02A"
    TempDB.Execute "ALTER TABLE tempstaff ADD CONSTRAINT tempstaffconstraint PRIMARY KEY (EMPLOYID);", dbSeeChanges: Call BriefDelay
'   DoCmd.OpenQuery "qryEXPIRATIONS03"
    TempDB.Execute "DELETE tempstaff.* FROM tempstaff WHERE LastName = 'EXAMPLE';", dbSeeChanges: Call BriefDelay

'   Create the temporary GPSupercodes table.
    Call AppendProgressMessages("Emptying the temporary GP Supervisors table.")
'   DoCmd.OpenQuery "qryEXPIRATIONS03A"
    TempDB.Execute "DELETE [~TempSuperCodes].* FROM [~TempSuperCodes];", dbSeeChanges: Call BriefDelay
'   DoCmd.OpenQuery "qryEXPIRATIONS04"
    TempDB.Execute "INSERT INTO [~TempSuperCodes] (GPCode, GPSuperCode, JobTitle) " & _
        "SELECT tblStaff.DEPRTMNT, tblStaff.SUPERVISORCODE_I, tblStaff.JOBTITLE FROM tblStaff " & _
        "WHERE (((tblStaff.JobTitle) = 'RESUNT' OR (tblStaff.JobTitle) = 'RESUPR' OR (tblStaff.JobTitle) = 'ASDRRE' OR (tblStaff.JobTitle) = 'DASUPR' OR (tblStaff.JobTitle) = 'SENDPM')) OR " & _
        "(((tblStaff.DEPRTMNT) = 'CHELSE') AND ((tblStaff.JobTitle) = 'PRGMGR')) OR " & _
        "(((tblStaff.DEPRTMNT) = 'NEWTON') AND ((tblStaff.JobTitle) = 'PRGMGR')) OR " & _
        "(((tblStaff.JobTitle) = 'RESMGR')) OR (((tblStaff.JobTitle) = 'SITECO')) " & _
        "ORDER BY tblStaff.DEPRTMNT;", dbSeeChanges: Call BriefDelay

'   Create a temporary table for staff skills.
    Call AppendProgressMessages("Creating temporary table for staff skills.")
    Call ExpirationsHousekeeping(TempDB, 2)
'   DoCmd.OpenQuery "qryEXPIRATIONS05"
    TempDB.Execute "SELECT tblStaffSkills.* INTO tempstaffskills FROM tblStaff " & _
        "LEFT JOIN tblStaffSkills ON tblStaff.EMPLOYID = tblStaffSkills.EMPID_I " & _
        "WHERE tblStaffSkills.SKILLNUMBER_I = 1 OR tblStaffSkills.SKILLNUMBER_I = 2 OR tblStaffSkills.SKILLNUMBER_I = 3 OR tblStaffSkills.SKILLNUMBER_I = 15 OR " & _
        "tblStaffSkills.SKILLNUMBER_I = 22 OR tblStaffSkills.SKILLNUMBER_I = 30 OR tblStaffSkills.SKILLNUMBER_I = 31 OR tblStaffSkills.SKILLNUMBER_I = 32 OR " & _
        "tblStaffSkills.SKILLNUMBER_I = 33 OR tblStaffSkills.SKILLNUMBER_I = 34 OR tblStaffSkills.SKILLNUMBER_I = 35 OR tblStaffSkills.SKILLNUMBER_I = 36 OR " & _
        "tblStaffSkills.SKILLNUMBER_I = 39;", dbSeeChanges: Call BriefDelay
'   DoCmd.OpenQuery "qryEXPIRATIONS05A"
    TempDB.Execute "ALTER TABLE tempstaffskills ADD CONSTRAINT tempstaffskillsconstraint PRIMARY KEY (EMPID_I, SKILLNUMBER_I);", dbSeeChanges: Call BriefDelay

'   Empty the expirations table.
    Call AppendProgressMessages("Empty the Expiration Dates table.")
'   DoCmd.OpenQuery "qryEXPIRATIONS05B"
    TempDB.Execute "DELETE tblExpirations.* FROM tblExpirations;", dbSeeChanges: Call BriefDelay

'   Build the program lookup table.  This looks for locations where a GPName is specified.  It goes through each program separately.
    Call AppendProgressMessages("Build base program lookup table.")
'   All staff with a GP Supervisor code.  Fix blank supercodes.
'   DoCmd.OpenQuery "qryEXPIRATIONS06"
    TempDB.Execute "SELECT [CityTown] & ' - ' & [LocationName] AS Location, tblLocations.CityTown, tblLocations.LocationName, tblLocations.GPName, tblPeople.GPSuperCode " & _
        "INTO temptbl " & _
        "FROM tblLocations INNER JOIN tblPeople ON (tblLocations.LocationName = tblPeople.OfficeLocationName) AND (tblLocations.CityTown = tblPeople.OfficeCityTown) " & _
        "WHERE tblLocations.GPName IS NOT NULL AND tblPeople.IsStaff = True " & _
        "ORDER BY [CityTown] & ' - ' & [LocationName]; ", dbSeeChanges: Call BriefDelay
'   DoCmd.OpenQuery "qryEXPIRATIONS07"
    TempDB.Execute "UPDATE temptbl " & _
        "SET temptbl.GPSuperCode = DLookUp ('GPSuperCode', '~TempSuperCodes', ""GPCode='"" & [temptbl].[GPName] & ""'"") " & _
        "WHERE ((([temptbl].[GPSuperCode]) IS NULL));", dbSeeChanges: Call BriefDelay
'   DoCmd.OpenQuery "qryEXPIRATIONS08"
    TempDB.Execute "INSERT INTO temptbl (Location, CityTown, LocationName, GPName, GPSuperCode) " & _
        "SELECT [CityTown] & ' - ' & [LocationName] AS Location, tblLocations.CityTown, tblLocations.LocationName, tblLocations.GPName, " & _
        "DLookUp ('GPSuperCode', 'tblPeople', ""FirstName='"" & tblLocations.StaffPrimaryContactFirstName & ""' AND LastName='"" & tblLocations.StaffPrimaryContactLastName & "" '"" ) AS Expr1 " & _
        "FROM tblLocations WHERE (((tblLocations.CityTown) <> 'Dedham') AND ((tblLocations.GPName) IS NOT NULL) AND ((tblLocations.Department) = 'Individualized Support Options')) " & _
        "ORDER BY [CityTown] & ' - ' & [LocationName];", dbSeeChanges: Call BriefDelay

    Call AppendProgressMessages("Parse program lookup table.")
    Call AppendProgressMessages("    This takes a few minutes.")
'   DoCmd.OpenQuery "qryEXPIRATIONS09"
    TempDB.Execute "SELECT IIf(IsNull([tblPeopleClientsResidentialServices]![CityTown]),'',[tblPeopleClientsResidentialServices]![CityTown] & ' - ' & [tblPeopleClientsResidentialServices]![Location]) AS LocRes, IIf(IsNull([tblPeopleClientsCLOServices]![CityTown]),'',[tblPeopleClientsCLOServices]![CityTown] & ' - ' & [tblPeopleClientsCLOServices]![Location]) AS LocCLO, IIf(IsNull([tblPeopleClientsDayServices]![CityTown]),'',[tblPeopleClientsDayServices]![CityTown] & ' - ' & [tblPeopleClientsDayServices]![LocationName]) AS LocDay, IIf(IsNull([tblPeopleClientsVocationalServices]![CityTown]),'',[tblPeopleClientsVocationalServices]![CityTown] & ' - ' & [tblPeopleClientsVocationalServices]![Location]) AS LocVoc, " & _
        "Null AS Supervisor, tblPeople.IndexedName, tblPeople.LastName, tblPeople.FirstName, tblPeople.MiddleInitial, tblPeopleClientsDemographics.DateISP, tblPeopleClientsDemographics.DateConsentFormsSigned, tblPeopleClientsDemographics.DateBMMExpires, tblPeopleClientsDemographics.DateBMMAccessSignedHRC, tblPeopleClientsDemographics.DateBMMAccessSigned, " & _
        "tblPeopleClientsDemographics.DateSPDAuthExpires, tblPeopleClientsDemographics.DateSignaturesDueBy, tblPeopleClientsDemographics.AllSPDSignaturesReceived INTO temptbl0 " & _
        "FROM ((((tblPeople RIGHT JOIN tblPeopleClientsDemographics ON tblPeople.IndexedName = tblPeopleClientsDemographics.IndexedName) LEFT JOIN tblPeopleClientsCLOServices ON tblPeopleClientsDemographics.IndexedName = tblPeopleClientsCLOServices.IndexedName) LEFT JOIN tblPeopleClientsDayServices ON tblPeopleClientsDemographics.IndexedName = tblPeopleClientsDayServices.IndexedName) LEFT JOIN tblPeopleClientsResidentialServices ON tblPeopleClientsDemographics.IndexedName = tblPeopleClientsResidentialServices.IndexedName) LEFT JOIN tblPeopleClientsVocationalServices ON tblPeopleClientsDemographics.IndexedName = tblPeopleClientsVocationalServices.IndexedName " & _
        "WHERE ((((tblPeopleClientsDemographics.ActiveDayServices)=True) AND ((tblPeopleClientsDayServices.Inactive)=False)) OR " & _
        "(((tblPeopleClientsDemographics.ActiveResidentialServices)=True) AND ((tblPeopleClientsResidentialServices.Inactive)=False)) OR " & _
        "(((tblPeopleClientsDemographics.ActiveCLO)=True) AND ((tblPeopleClientsCLOServices.Inactive)=False)) OR " & _
        "(((tblPeopleClientsDemographics.ActiveVocationalServices)=True) AND ((tblPeopleClientsVocationalServices.Inactive)=False))) AND tblPeople.IsDeceased = False;", dbSeeChanges: Call BriefDelay
    
'   CLO.
    Call AppendProgressMessages("Populate program lookup table with CLO information.")
'   DoCmd.OpenQuery "qryEXPIRATIONS10"
    TempDB.Execute "UPDATE (temptbl0 LEFT JOIN qrytblPeopleClientsDemographics ON temptbl0.IndexedName = qrytblPeopleClientsDemographics.IndexedName) LEFT JOIN tblPeopleClientsCLOServices ON temptbl0.IndexedName = tblPeopleClientsCLOServices.IndexedName SET temptbl0.LocCLO = Null " & _
        "WHERE (Len(temptbl0!LocCLO)>0 And qrytblPeopleClientsDemographics.ActiveCLO=False) Or (Len(temptbl0!LocCLO)>0 And qrytblPeopleClientsDemographics.ActiveCLO=True And tblPeopleClientsCLOServices.Inactive=True);", dbSeeChanges: Call BriefDelay
    
'   Residential.
    Call AppendProgressMessages("Populate program lookup table with residential program information.")
'   DoCmd.OpenQuery "qryEXPIRATIONS11
    TempDB.Execute "UPDATE (temptbl0 LEFT JOIN qrytblPeopleClientsDemographics ON temptbl0.IndexedName = qrytblPeopleClientsDemographics.IndexedName) LEFT JOIN tblPeopleClientsResidentialServices ON temptbl0.IndexedName = tblPeopleClientsResidentialServices.IndexedName SET temptbl0.LocRes = Null " & _
        "WHERE (((Len(temptbl0!LocRes))>0) And ((qrytblPeopleClientsDemographics.ActiveResidentialServices)=False)) Or (((Len(temptbl0!LocRes))>0) And ((qrytblPeopleClientsDemographics.ActiveResidentialServices)=True) And ((tblPeopleClientsResidentialServices.Inactive)=True));", dbSeeChanges: Call BriefDelay
    
'   Day.
    Call AppendProgressMessages("Populate program lookup table with day program information.")
'   DoCmd.OpenQuery "qryEXPIRATIONS12"
    TempDB.Execute "UPDATE (temptbl0 LEFT JOIN qrytblPeopleClientsDemographics ON temptbl0.IndexedName = qrytblPeopleClientsDemographics.IndexedName) LEFT JOIN tblPeopleClientsDayServices ON temptbl0.IndexedName = tblPeopleClientsDayServices.IndexedName SET temptbl0.LocDay = Null " & _
        "WHERE (((Len(temptbl0!LocDay))>0) And ((qrytblPeopleClientsDemographics.ActiveDayServices)=False)) Or (((Len(temptbl0!LocDay))>0) And ((qrytblPeopleClientsDemographics.ActiveDayServices)=True) And ((tblPeopleClientsDayServices.Inactive)=True));", dbSeeChanges: Call BriefDelay
    
'   Vocational.
    Call AppendProgressMessages("Populate program lookup table with vocational program information.")
'   DoCmd.OpenQuery "qryEXPIRATIONS13"
    TempDB.Execute "UPDATE (temptbl0 LEFT JOIN qrytblPeopleClientsDemographics ON temptbl0.IndexedName = qrytblPeopleClientsDemographics.IndexedName) LEFT JOIN tblPeopleClientsVocationalServices ON temptbl0.IndexedName = tblPeopleClientsVocationalServices.IndexedName SET temptbl0.LocVoc = Null " & _
        "WHERE (((Len(temptbl0!LocVoc))>0) And ((qrytblPeopleClientsDemographics.ActiveVocationalServices)=False)) Or (((Len(temptbl0!LocVoc))>0) And ((qrytblPeopleClientsDemographics.ActiveVocationalServices)=True) And ((tblPeopleClientsVocationalServices.Inactive)=True));", dbSeeChanges: Call BriefDelay

'   Populate house information.
    Call AppendProgressMessages("Populate house information.")
'   DoCmd.OpenQuery "qryEXPIRATIONS14"
    TempDB.Execute "INSERT INTO tblExpirations (Location, RecordType, LastName, FirstName, Supervisor, LastVehicleChecklistCompleted, MostRecentAsleepFireDrill, NextRecentAsleepFireDrill, " & _
        "HouseSafetyPlanExpires, HousePlansReviewedByStaffBefore, DAYStaffTrainedInPrivacyBefore, DAYAllPlansReviewedByStaffBefore, DAYQtrlySafetyChecklistDueBy, MAPChecklistCompleted, " & _
        "HumanRightsOfficer, HROTrainsStaffBefore, HROTrainsIndividualsBefore, FireSafetyOfficer, FSOTrainsStaffBefore, FSOTrainsIndividualsBefore) " & _
        "SELECT tblLocations.GPName AS Location, 'House' AS RecordType, '*' AS LastName, '*' AS FirstName, " & _
        "DLookUp ('GPSuperCode', 'temptbl', ""GPName='"" & tblLocations.GPName & ""'"") AS Supervisor, " & _
        "tblLocations.LastVehicleChecklistCompleted, tblLocations.MostRecentAsleepFireDrill, tblLocations.NextRecentAsleepFireDrill, tblLocations.HouseSafetyPlanExpires," & _
        "tblLocations.HousePlansReviewedByStaffBefore, tblLocations.DAYStaffTrainedInPrivacyBefore, tblLocations.DAYAllPlansReviewedByStaffBefore, tblLocations.DAYQtrlySafetyChecklistDueBy, " & _
        "tblLocations.MAPChecklistCompleted, tblLocations.HumanRightsOfficer, tblLocations.HROTrainsStaffBefore, tblLocations.HROTrainsIndividualsBefore, " & _
        "tblLocations.FireSafetyOfficer, tblLocations.FSOTrainsStaffBefore, tblLocations.FSOTrainsIndividualsBefore FROM tblLocations " & _
        "WHERE tblLocations.GPName IS NOT NULL AND DLookUp('GPSuperCode', 'temptbl', 'GPName=""' & tblLocations.GPName & '""') IS NOT NULL AND " & _
        "tblLocations.Department <> 'Clinical and Support Services' ORDER BY tblLocations.GPName;", dbSeeChanges: Call BriefDelay
    Call AppendProgressMessages("Populate CLO client information.")
'   DoCmd.OpenQuery "qryEXPIRATIONS15"
    TempDB.Execute "INSERT INTO tblExpirations (Location, RecordType, LastName, FirstName, Supervisor, DateISP, DateConsentFormsSigned, DateBMMExpires, DateBMMAccessSignedHRC, DateBMMAccessSigned, DateSPDAuthExpires, DateSignaturesDueBy, AllSPDSignaturesReceived) " & _
        "SELECT DLookUp (""GPName"", ""temptbl"", ""Location='"" & [LocCLO] & ""'"") AS Location, 'Client' AS RecordType, temptbl0.LastName, temptbl0.FirstName, " & _
        "DLookUp (""GPSuperCode"", ""temptbl"", ""Location='"" & [LocCLO] & ""'"") AS Supervisor, temptbl0.DateISP, temptbl0.DateConsentFormsSigned, temptbl0.DateBMMExpires, " & _
        "temptbl0.DateBMMAccessSignedHRC, temptbl0.DateBMMAccessSigned, temptbl0.DateSPDAuthExpires, temptbl0.DateSignaturesDueBy, temptbl0.AllSPDSignaturesReceived " & _
        "FROM temptbl0 " & _
        "WHERE (((DLookUp(""GPName"", ""temptbl"", ""Location='"" & [LocCLO] & ""'"")) IS NOT NULL) AND ((temptbl0.LastName) IS NOT NULL) AND ((temptbl0.FirstName) IS NOT NULL));", dbSeeChanges: Call BriefDelay
    Call AppendProgressMessages("Populate residential client information.")
'   DoCmd.OpenQuery "qryEXPIRATIONS16"
    TempDB.Execute "INSERT INTO tblExpirations (Location, RecordType, LastName, FirstName, Supervisor, DateISP, DateConsentFormsSigned, DateBMMExpires, DateBMMAccessSignedHRC, DateBMMAccessSigned, DateSPDAuthExpires, DateSignaturesDueBy, AllSPDSignaturesReceived) " & _
        "SELECT DLookUp(""GPName"", ""temptbl"", ""Location='"" & [LocRes] & ""'"") AS Location, 'Client' AS RecordType, temptbl0.LastName, temptbl0.FirstName, " & _
        "DLookUp (""GPSuperCode"", ""temptbl"", ""Location='"" & [LocRes] & ""'"") AS Supervisor, temptbl0.DateISP, temptbl0.DateConsentFormsSigned, temptbl0.DateBMMExpires, temptbl0.DateBMMAccessSignedHRC, temptbl0.DateBMMAccessSigned, temptbl0.DateSPDAuthExpires, temptbl0.DateSignaturesDueBy, temptbl0.AllSPDSignaturesReceived " & _
        "FROM temptbl0 " & _
        "WHERE (((DLookUp(""GPName"", ""temptbl"", ""Location='"" & [LocRes] & ""'"")) IS NOT NULL) AND ((temptbl0.LastName) IS NOT NULL) AND ((temptbl0.FirstName) IS NOT NULL));", dbSeeChanges: Call BriefDelay
    Call AppendProgressMessages("Populate vocational client information.")
'   DoCmd.OpenQuery "qryEXPIRATIONS17"
    TempDB.Execute "INSERT INTO tblExpirations (Location, RecordType, LastName, FirstName, Supervisor, DateISP, DateConsentFormsSigned, DateBMMExpires, DateBMMAccessSignedHRC, DateBMMAccessSigned, DateSPDAuthExpires, DateSignaturesDueBy, AllSPDSignaturesReceived) " & _
        "SELECT DLookUp(""GPName"", ""temptbl"", ""Location='"" & [LocVoc] & ""'"") AS Location, 'Client' AS RecordType, temptbl0.LastName, temptbl0.FirstName," & _
        "DLookUp(""GPSuperCode"", ""temptbl"", ""Location='"" & [LocVoc] & ""'"") AS Supervisor, temptbl0.DateISP, temptbl0.DateConsentFormsSigned, temptbl0.DateBMMExpires, temptbl0.DateBMMAccessSignedHRC, temptbl0.DateBMMAccessSigned, temptbl0.DateSPDAuthExpires, temptbl0.DateSignaturesDueBy, temptbl0.AllSPDSignaturesReceived " & _
        "FROM temptbl0 " & _
        "WHERE (((DLookUp(""GPName"", ""temptbl"", ""Location='"" & [LocVoc] & ""'"")) IS NOT NULL) AND ((temptbl0.LastName) IS NOT NULL) AND ((temptbl0.FirstName) IS NOT NULL));", dbSeeChanges: Call BriefDelay

'   Populate staff information.
    Call AppendProgressMessages("Populate staff information.")
'   DoCmd.OpenQuery "qryEXPIRATIONS18"
    TempDB.Execute "INSERT INTO tblExpirations ( Location, RecordType, LastName, FirstName, JobTitle, Supervisor, AdjustedStartDate ) " & _
        "SELECT [tempstaff]![DEPRTMNT] AS Location, 'Staff' AS RecordType, tempstaff.LASTNAME, tempstaff.FRSTNAME, tempstaff.JOBTITLE, tempstaff.SUPERVISORCODE_I, tempstaff.BENADJDATE AS AdjustedStartDate " & _
        "FROM tempstaff INNER JOIN tempstaffskills ON tempstaff.EMPLOYID = tempstaffskills.EMPID_I " & _
        "WHERE tempstaff.DEPRTMNT Is Not Null And tempstaff.LastName Is Not Null And tempstaff.FRSTNAME Is Not Null " & _
        "ORDER BY tempstaff.LASTNAME, tempstaff.FRSTNAME;", dbSeeChanges: Call BriefDelay

'   Update staff information.
    Call AppendProgressMessages("Populate staff skills information and descriptors.")
    Call AppendProgressMessages("    This takes a few minutes.")
'   DoCmd.OpenQuery "qryEXPIRATIONS19"
    TempDB.Execute "SELECT tempstaffskills.EMPID_I, tempstaffskills.SKILLNUMBER_I, tempstaffskills.EXPIREDSKILL_I INTO temptbl1 " & _
        "FROM tempstaffskills INNER JOIN tblStaff ON tempstaffskills.EMPID_I = tblStaff.EMPLOYID " & _
        "WHERE tempstaffskills.SKILLNUMBER_I = 1 OR tempstaffskills.SKILLNUMBER_I = 2 OR tempstaffskills.SKILLNUMBER_I = 3 OR tempstaffskills.SKILLNUMBER_I = 15 " & _
        "OR tempstaffskills.SKILLNUMBER_I = 22 OR tempstaffskills.SKILLNUMBER_I = 30 OR tempstaffskills.SKILLNUMBER_I = 31  OR tempstaffskills.SKILLNUMBER_I = 32 " & _
        "OR tempstaffskills.SKILLNUMBER_I = 33 OR tempstaffskills.SKILLNUMBER_I = 34 OR tempstaffskills.SKILLNUMBER_I = 35 OR tempstaffskills.SKILLNUMBER_I = 36 OR tempstaffskills.SKILLNUMBER_I = 39;", dbSeeChanges: Call BriefDelay
'   DoCmd.OpenQuery "qryEXPIRATIONS20"
    TempDB.Execute "SELECT temptbl1.*, DLookUp(""Skill"", ""catSkills"", ""SkillID="" & [SKILLNUMBER_I]) AS SkillDesc INTO temptbl2 FROM temptbl1;", dbSeeChanges: Call BriefDelay
    
    Call AppendProgressMessages("Crosstabulate staff skills.")
'   DoCmd.OpenQuery "qryEXPIRATIONS21"
    TempDB.Execute "SELECT * INTO temptbl3 FROM qryExpirationsStaffBySkills;", dbSeeChanges: Call BriefDelay
    
    Call AppendProgressMessages("Populate expirations table with staff skills.")
'   DoCmd.OpenQuery "qryEXPIRATIONS22"
    TempDB.Execute "UPDATE qrytblExpirations INNER JOIN temptbl3 ON (qrytblExpirations.FirstName = temptbl3.FRSTNAME) AND (qrytblExpirations.LastName = temptbl3.LASTNAME) AND (qrytblExpirations.Location = temptbl3.DEPRTMNT) " & _
        "SET qrytblExpirations.CPR = [temptbl3].[CPR], qrytblExpirations.FirstAid = [temptbl3].[FirstAid], qrytblExpirations.MAPCert = [temptbl3].[MAPCert], " & _
        "qrytblExpirations.DriversLicense = [temptbl3].[DriversLicense], qrytblExpirations.BBP = [temptbl3].[BBP], qrytblExpirations.BackInjuryPrevention = [temptbl3].[BackInjuryPrevention], " & _
        "qrytblExpirations.SafetyCares = [temptbl3].[SafetyCares], qrytblExpirations.TB = [temptbl3].[TB], qrytblExpirations.WorkplaceViolence = [temptbl3].[WorkplaceViolence], qrytblExpirations.DefensiveDriving = [temptbl3].[DefensiveDriving], " & _
        "qrytblExpirations.WheelchairSafety = [temptbl3].[WheelchairSafety], qrytblExpirations.PBS = [temptbl3].[PBS], qrytblExpirations.ProfessionalLicenses = [temptbl3].[ProfLic] " & _
        "WHERE (((qrytblExpirations.RecordType) = 'Staff'));", dbSeeChanges: Call BriefDelay
    
    Call AppendProgressMessages("Populate expiration dates table with staff evals and supervisions.")
'   DoCmd.OpenQuery "qryEXPIRATIONS23"
    TempDB.Execute "UPDATE qrytblExpirations INNER JOIN qrytblStaffEvalsAndSupervisions ON (qrytblExpirations.FirstName = qrytblStaffEvalsAndSupervisions.FirstName) AND (qrytblExpirations.LastName = qrytblStaffEvalsAndSupervisions.LastName) SET qrytblExpirations.ThreeMonthEvaluation = [qrytblStaffEvalsAndSupervisions]![ThreeMonthEval], qrytblExpirations.EvalDueBy = [qrytblStaffEvalsAndSupervisions]![EvalDueBy], qrytblExpirations.LastSupervision = [qrytblStaffEvalsAndSupervisions]![LastSupervision], qrytblExpirations.OnLeave = [qrytblStaffEvalsAndSupervisions]![OnLeave] " & _
        "WHERE qrytblExpirations.RecordType='Staff';", dbSeeChanges: Call BriefDelay

'   Finally, see if we can add supervisor expirations to their own report.
    Call AppendProgressMessages("Add supervisor's expirations to table.")
'   DoCmd.OpenQuery "qryEXPIRATIONS24"
    TempDB.Execute "INSERT INTO tblExpirations (Location, RecordType, LastName, FirstName, Supervisor, JobTitle, AdjustedStartDate, LastVehicleChecklistCompleted, MostRecentAsleepFireDrill, NextRecentAsleepFireDrill, HouseSafetyPlanExpires, HousePlansReviewedByStaffBefore, DAYStaffTrainedInPrivacyBefore, DAYAllPlansReviewedByStaffBefore, DAYQtrlySafetyChecklistDueBy, MAPChecklistCompleted, HumanRightsOfficer, HROTrainsStaffBefore, HROTrainsIndividualsBefore, FireSafetyOfficer, FSOTrainsStaffBefore, FSOTrainsIndividualsBefore, DateISP, DateConsentFormsSigned, DateBMMExpires, DateBMMAccessSignedHRC, DateBMMAccessSigned, DateSPDAuthExpires, DateSignaturesDueBy, AllSPDSignaturesReceived, BBP, BackInjuryPrevention, CPR, DefensiveDriving, DriversLicense, FirstAid, MAPCert, PBS, SafetyCares, TB, WheelchairSafety, WorkplaceViolence, ProfessionalLicenses, ThreeMonthEvaluation, EvalDueBy, LastSupervision, OnLeave ) " & _
        "SELECT tblLocations.GPName, 'Staff' AS RecordType, tblExpirations.LastName, tblExpirations.FirstName, tblPeople.GPSuperCode AS Supervisor, tblExpirations.JobTitle, tblExpirations.AdjustedStartDate, tblExpirations.LastVehicleChecklistCompleted, tblExpirations.MostRecentAsleepFireDrill, tblExpirations.NextRecentAsleepFireDrill, tblExpirations.HouseSafetyPlanExpires, tblExpirations.HousePlansReviewedByStaffBefore, tblExpirations.DAYStaffTrainedInPrivacyBefore, tblExpirations.DAYAllPlansReviewedByStaffBefore, tblExpirations.DAYQtrlySafetyChecklistDueBy, tblExpirations.MAPChecklistCompleted, tblExpirations.HumanRightsOfficer, tblExpirations.HROTrainsStaffBefore, tblExpirations.HROTrainsIndividualsBefore, tblExpirations.FireSafetyOfficer, tblExpirations.FSOTrainsStaffBefore, tblExpirations.FSOTrainsIndividualsBefore, tblExpirations.DateISP, tblExpirations.DateConsentFormsSigned, tblExpirations.DateBMMExpires, tblExpirations.DateBMMAccessSignedHRC, " & _
        "tblExpirations.DateBMMAccessSigned , tblExpirations.DateSPDAuthExpires , tblExpirations.DateSignaturesDueBy, tblExpirations.AllSPDSignaturesReceived, tblExpirations.BBP, tblExpirations.BackInjuryPrevention, tblExpirations.CPR, tblExpirations.DefensiveDriving, tblExpirations.DriversLicense, tblExpirations.FirstAid, tblExpirations.MAPCert, tblExpirations.PBS, tblExpirations.SafetyCares, tblExpirations.TB, tblExpirations.WheelchairSafety, tblExpirations.WorkplaceViolence, tblExpirations.ProfessionalLicenses, tblExpirations.ThreeMonthEvaluation, tblExpirations.EvalDueBy, tblExpirations.LastSupervision, tblExpirations.OnLeave " & _
        "FROM (tblLocations INNER JOIN tblPeople ON tblLocations.StaffPrimaryContactIndexedName = tblPeople.IndexedName) INNER JOIN tblExpirations ON (tblPeople.FirstName = tblExpirations.FirstName) AND (tblPeople.LastName = tblExpirations.LastName) " & _
        "WHERE (((tblLocations.GPName) Is Not Null) And ((tblPeople.GPSuperCode) Is Not Null) And ((tblLocations.Department) = 'Residential Services' OR (tblLocations.Department) = 'Day Services' Or (tblLocations.Department) = 'Vocational Services' Or (tblLocations.Department) = 'TILL NH' Or (tblLocations.Department) = 'Expirations Reporting') And ((Right(tblLocations!StaffPrimaryContactIndexedName, 5)) <> 'TBD//')) Or (((tblLocations.GPName) Is Not Null) And ((tblPeople.GPSuperCode) Is Not Null) And ((tblLocations.CityTown) <> 'Dedham') And ((tblLocations.Department) = 'Individualized Support Options') And ((Right(tblLocations!StaffPrimaryContactIndexedName, 5)) <> 'TBD//')) " & _
        "ORDER BY tblLocations.CityTown;", dbSeeChanges: Call BriefDelay

'   Here, produce the report.
    Call AppendProgressMessages("Generating the report.")
'   Call ExecReport("rptEXPIRATIONDATES"): Call BriefDelay
'   We export the report to a PDF without displaying it.
    ExportFileName = Application.CurrentProject.Path & "\" & "TILLDB-Report-ExpirationDates-" & Format(Date, "yyyymmdd") & ".pdf"
    If IsFileOpen(ExportFileName) Then
        If MsgBox(ExportFileName & " is already open.  Please close it and click OK to continue or Cancel to abort.", vbOKCancel, "ERROR!") = vbCancel Then
            MsgBox "Export aborted.", vbOKOnly, "Aborted"
        Else
            If Dir(ExportFileName) <> "" Then Kill ExportFileName
            DoCmd.OutputTo acOutputReport, "rptEXPIRATIONSDATES", acFormatPDF, ExportFileName, False
        MsgBox "The requested report has been exported to " & ExportFileName & "." & vbCrLf & vbCrLf & "This export may contain information that is protected under HIPAA and other privacy laws.  This export must be securely stored at all times and must be deleted when no longer being used.", _
            vbOKOnly, "Export Complete"
        End If
    Else
        If Dir(ExportFileName) <> "" Then Kill ExportFileName
        DoCmd.OutputTo acOutputReport, "rptEXPIRATIONDATES", acFormatPDF, ExportFileName, False
        MsgBox "The requested report has been exported to " & ExportFileName & "." & vbCrLf & vbCrLf & "This export may contain information that is protected under HIPAA and other privacy laws.  This export must be securely stored at all times and must be deleted when no longer being used.", _
            vbOKOnly, "Export Complete"
    End If

'   Empty the Expiration Dates table.
    Call AppendProgressMessages("Cleanup.")
    Call ExpirationsHousekeeping(TempDB, 1): Call ExpirationsHousekeeping(TempDB, 2)
'   DoCmd.OpenQuery "qryEXPIRATIONS25"
    TempDB.Execute "DELETE tblExpirations.* FROM tblExpirations;", dbSeeChanges: Call BriefDelay
'   DoCmd.OpenQuery "qryEXPIRATIONS26"
    TempDB.Execute "DELETE [~TempSuperCodes].* FROM [~TempSuperCodes];", dbSeeChanges: Call BriefDelay
    
    Form_frmRpt.ProgressMessages = "": Form_frmRpt.ProgressMessages.Requery
    Call DropTempTables
    
    DoCmd.SetWarnings True
    SysCmdResult = SysCmd(5)
    Exit Function
ShowMeError:
    Call DropTempTables
    Call ExpirationsHousekeeping(TempDB, 1): Call ExpirationsHousekeeping(TempDB, 2)
'   DoCmd.OpenQuery "qryEXPIRATIONS25"
    TempDB.Execute "DELETE tblExpirations.* FROM tblExpirations;", dbSeeChanges: Call BriefDelay
'   DoCmd.OpenQuery "qryEXPIRATIONS26"
    TempDB.Execute "DELETE [~TempSuperCodes].* FROM [~TempSuperCodes];", dbSeeChanges: Call BriefDelay
    Err.Source = "PublicSubroutines" & "(Line #" & Str(Err.Erl) & ")": TILLDBErrorMessage = "Error # " & Str(Err.Number) & " was generated by " & Err.Source & Chr(13) & Err.Description
    MsgBox TILLDBErrorMessage, vbOKOnly, "Error", Err.HelpFile, Err.HelpContext
    If ExpDatesReportInitiatedFromReportsMenu Then
        Form_frmRpt.ProgressMessages = "": Form_frmRpt.ProgressMessages.Requery
    Else
        SysCmdResult = SysCmd(5)
    End If
    DoCmd.SetWarnings True
End Function

Public Function RunRedReportNew(SupervisionsOnly As Boolean)
    Dim TempDB As Database
'   On Error GoTo ShowMeError

    RunRedReportNew = False
    Set TempDB = CurrentDb
    TempDB.QueryTimeout = 0
    DoCmd.SetWarnings False
    Form_frmRpt.ProgressMessages = ""
    Form_frmRpt.ProgressMessages.Requery
    
    If DontRunExpirations Then
        MsgBox "There are no records in the Staff or Staff Skills tables." & vbCrLf & "SQL Server Refresh Skills job likely failed." & vbCrLf & _
               "You cannot run Expiration Dates report or Staff Evals and Supervisions until fixed." & vbCrLf & "Contact Tech Services to resolve.", vbOKOnly, "Error!"
        Exit Function
    End If
    
    Call AppendProgressMessages("Housekeeping before starting the process.")
    Call BriefDelay(1):     If IsTableQuery("temptbl9") Then TempDB.Execute "DROP TABLE temptbl9", dbSeeChanges: Call BriefDelay
    Call BriefDelay(1):     If IsTableQuery("temptbl8") Then TempDB.Execute "DROP TABLE temptbl8", dbSeeChanges: Call BriefDelay
    Call BriefDelay(1):     If IsTableQuery("temptbl7") Then TempDB.Execute "DROP TABLE temptbl7", dbSeeChanges: Call BriefDelay
    Call BriefDelay(1):     If IsTableQuery("temptbl6") Then TempDB.Execute "DROP TABLE temptbl6", dbSeeChanges: Call BriefDelay
    Call BriefDelay(1):     If IsTableQuery("temptbl5") Then TempDB.Execute "DROP TABLE temptbl5", dbSeeChanges: Call BriefDelay
    Call BriefDelay(1):     If IsTableQuery("temptbl4") Then TempDB.Execute "DROP TABLE temptbl4", dbSeeChanges: Call BriefDelay
    Call BriefDelay(1):     If IsTableQuery("temptbl3") Then TempDB.Execute "DROP TABLE temptbl3", dbSeeChanges: Call BriefDelay
    Call BriefDelay(1):     If IsTableQuery("temptbl2") Then TempDB.Execute "DROP TABLE temptbl2", dbSeeChanges: Call BriefDelay
    Call BriefDelay(1):     If IsTableQuery("temptbl1") Then TempDB.Execute "DROP TABLE temptbl1", dbSeeChanges: Call BriefDelay
    Call BriefDelay(1):     If IsTableQuery("temptbl0") Then TempDB.Execute "DROP TABLE temptbl0", dbSeeChanges: Call BriefDelay
    Call BriefDelay(1):     If IsTableQuery("temptbl") Then TempDB.Execute "DROP TABLE temptbl", dbSeeChanges: Call BriefDelay

    Call AppendProgressMessages("Create temporary staff table.")
    Call BriefDelay
    TempDB.Execute "SELECT EMPLOYID, EMPLCLAS, INACTIVE, LASTNAME, FRSTNAME, MIDLNAME, " & _
        "DIVISIONCODE_I, DEPRTMNT, JOBTITLE, SUPERVISORCODE_I, STRTDATE, BENADJDATE INTO temptbl0 FROM tblStaff;", dbSeeChanges: Call BriefDelay

    Call AppendProgressMessages("Create temporary staff skills table.")
    Call BriefDelay
    TempDB.Execute "CREATE TABLE temptbl1 (EMPID_I VARCHAR(8) NOT NULL, SKILLNUMBER_I INT, Skill VARCHAR(25), EXPIREDSKILL_I DATE, LASTNAME VARCHAR(30), FRSTNAME VARCHAR(30), " & _
        "DIVISIONCODE_I VARCHAR(8), LOCATION VARCHAR(8), LocationName VARCHAR(100), SUPERVISORCODE_I VARCHAR(6), SupervisorIndexedName VARCHAR(160), RedFlag BIT, OnLeave BIT);", dbSeeChanges: Call BriefDelay
    
    Call AppendProgressMessages("Add active skills records for selected skills and valid dates.")
    Call AppendProgressMessages("   This may take a few minutes.")
    Call BriefDelay
    TempDB.Execute "INSERT INTO temptbl1 (EMPID_I, SKILLNUMBER_I, Skill, EXPIREDSKILL_I, LASTNAME, FRSTNAME, DIVISIONCODE_I, LOCATION, LocationName, SUPERVISORCODE_I, SupervisorIndexedName, RedFlag, OnLeave) " & _
        "SELECT EMPID_I, SKILLNUMBER_I, """" AS Skill, EXPIREDSKILL_I, StrConv (temptbl0.LASTNAME, 3) AS LASTNAME, StrConv (temptbl0.FRSTNAME, 3) AS FRSTNAME, Trim(temptbl0.DIVISIONCODE_I) AS DIVISIONCODE_I, Trim(temptbl0.DEPRTMNT) AS LOCATION, """" AS LocationName, " & _
        "temptbl0.SUPERVISORCODE_I, """" AS SupervisorIndexedName, False AS RedFlag, False As OnLeave " & _
        "FROM tblStaffSkills INNER JOIN temptbl0 ON tblStaffSkills.EMPID_I = temptbl0.EMPLOYID " & _
        "WHERE (SKILLNUMBER_I =  1 AND Year(EXPIREDSKILL_I) > 1980 AND temptbl0.INACTIVE = 0) OR " & _
              "(SKILLNUMBER_I =  2 AND Year(EXPIREDSKILL_I) > 1980 AND temptbl0.INACTIVE = 0) OR " & _
              "(SKILLNUMBER_I =  3 AND Year(EXPIREDSKILL_I) > 1980 AND temptbl0.INACTIVE = 0) OR " & _
              "(SKILLNUMBER_I = 12 AND Year(EXPIREDSKILL_I) > 1980 AND temptbl0.INACTIVE = 0) OR " & _
              "(SKILLNUMBER_I = 15 AND Year(EXPIREDSKILL_I) > 1980 AND temptbl0.INACTIVE = 0) OR " & _
              "(SKILLNUMBER_I = 22 AND Year(EXPIREDSKILL_I) > 1980 AND temptbl0.INACTIVE = 0) OR " & _
              "(SKILLNUMBER_I = 30 AND Year(EXPIREDSKILL_I) > 1980 AND temptbl0.INACTIVE = 0) OR " & _
              "(SKILLNUMBER_I = 31 AND Year(EXPIREDSKILL_I) > 1980 AND temptbl0.INACTIVE = 0) OR " & _
              "(SKILLNUMBER_I = 35 AND Year(EXPIREDSKILL_I) > 1980 AND temptbl0.INACTIVE = 0) OR " & _
              "(SKILLNUMBER_I = 36 AND Year(EXPIREDSKILL_I) > 1980 AND temptbl0.INACTIVE = 0) OR " & _
              "(SKILLNUMBER_I = 39 AND Year(EXPIREDSKILL_I) > 1980 AND temptbl0.INACTIVE = 0) OR " & _
              "(SKILLNUMBER_I = 12 AND Year(EXPIREDSKILL_I) > 1980 AND temptbl0.INACTIVE = 0) " & _
        "ORDER BY LASTNAME, FRSTNAME, SKILLNUMBER_I;", dbSeeChanges: Call BriefDelay

    Call BriefDelay
    TempDB.Execute "UPDATE temptbl1 INNER JOIN tblStaff ON temptbl1.EMPID_I = tblStaff.EMPLOYID SET LOCATION = Trim(tblStaff.DEPRTMNT);", dbSeeChanges: Call BriefDelay

    Call AppendProgressMessages("Create temporary locations table.")
    Call BriefDelay
    TempDB.Execute "SELECT tblLocations.CityTown & '-' & tblLocations.LocationName AS LocationName, GPName INTO temptbl2 FROM tblLocations WHERE CityTown <> 'HQ';", dbSeeChanges: Call BriefDelay

    Call AppendProgressMessages("Map Locations in temporary staff skills table.")
    Call BriefDelay
    TempDB.Execute "UPDATE temptbl1 INNER JOIN temptbl2 ON temptbl1.LOCATION = temptbl2.GPName SET temptbl1.LocationName = [temptbl2].[LocationName];", dbSeeChanges: Call BriefDelay
    Call BriefDelay
    TempDB.Execute "UPDATE temptbl1 SET LocationName = 'Dedham - HQ' WHERE FieldIsEmpty(Trim(temptbl1.LocationName)) AND ((Trim(temptbl1.DIVISIONCODE_I = 'ADMINI') OR Trim(temptbl1.DIVISIONCODE_I = 'DEDHAM')) AND (Trim(temptbl1.LOCATION) = 'ADMIN' OR Trim(temptbl1.LOCATION) = 'DEDHAM'));", dbSeeChanges: Call BriefDelay
    Call BriefDelay
    TempDB.Execute "UPDATE temptbl1 SET LocationName = 'Dedham - HQ' WHERE FieldIsEmpty(Trim(temptbl1.LocationName)) AND (Trim(temptbl1.DIVISIONCODE_I) = 'CHOICE');", dbSeeChanges: Call BriefDelay
    Call BriefDelay
    TempDB.Execute "UPDATE temptbl1 SET LocationName = 'Dedham - HQ' WHERE FieldIsEmpty(Trim(temptbl1.LocationName)) AND (Trim(temptbl1.DIVISIONCODE_I) = 'RESIDE' OR Trim(temptbl1.DIVISIONCODE_I) = 'DEDHAM') AND Trim(temptbl1.LOCATION) = 'RESIDE';", dbSeeChanges: Call BriefDelay
    Call BriefDelay
    TempDB.Execute "UPDATE temptbl1 SET LocationName = 'Dedham - HQ' WHERE FieldIsEmpty(Trim(temptbl1.LocationName)) AND (Trim(temptbl1.DIVISIONCODE_I) = 'SUPPOR');", dbSeeChanges: Call BriefDelay
    Call BriefDelay
    TempDB.Execute "UPDATE temptbl1 SET LocationName = 'Dedham - HQ' WHERE FieldIsEmpty(Trim(temptbl1.LocationName)) AND (Trim(temptbl1.DIVISIONCODE_I) = 'TRANSP');", dbSeeChanges: Call BriefDelay
    Call BriefDelay
    TempDB.Execute "UPDATE temptbl1 SET LocationName = '*** NO LOCATION ***' WHERE FieldIsEmpty(Trim(temptbl1.LOCATION));", dbSeeChanges: Call BriefDelay

    Call AppendProgressMessages("Create temporary staff supervisors table.")
    Call BriefDelay
    TempDB.Execute "SELECT Trim(tblPeopleStaffSupervisors.SUPERVISORCODE_I) AS SUPERVISORCODE_I, Trim(tblPeopleStaffSupervisors.SUPERVISORNAME) AS SUPERVISORNAME, " & _
        "tblPeopleStaffSupervisors.SUPEMPLID, Trim(tblPeopleStaffSupervisors.LASTNAME) AS LASTNAME, Trim(tblPeopleStaffSupervisors.FIRSTNAME) AS FRSTNAME, " & _
        "Trim(tblPeopleStaffSupervisors.INDEXEDNAME) AS INDEXEDNAME, Trim(tblPeopleStaffSupervisors.LOCATION) AS LOCATION " & _
        "INTO temptbl3 FROM tblPeopleStaffSupervisors;", dbSeeChanges: Call BriefDelay
    
    Call AppendProgressMessages("Update staff supervisors properties into temporary staff skills table.")
    Call BriefDelay
    TempDB.Execute "UPDATE temptbl1 INNER JOIN temptbl3 ON temptbl1.SUPERVISORCODE_I = temptbl3.SUPERVISORCODE_I " & _
        "SET temptbl1.SupervisorIndexedName = Trim(temptbl3.INDEXEDNAME);"

    Call AppendProgressMessages("Create temporary staff evaluations table.")
    Call AppendProgressMessages("   This may take a few minutes.")
    Call BriefDelay
    TempDB.Execute "SELECT Trim(tblStaffEvalsAndSupervisions.EmployeeID) AS EmployeeID, Trim(tblStaffEvalsAndSupervisions.Location) AS Location, " & _
        "Trim(tblStaffEvalsAndSupervisions.LastName) AS LastName, Trim(tblStaffEvalsAndSupervisions.FirstName) AS FirstName, " & _
        "Trim(tblStaffEvalsAndSupervisions.JobTitle) AS JobTitle, DateValue(tblStaffEvalsAndSupervisions.EvalDueBy) AS EvalDueBy, ThreeMonthEval, " & _
        "DateValue(tblStaffEvalsAndSupervisions.LastSupervision) AS LastSupervisionDate, OnLeave, Trim(tblStaffEvalsAndSupervisions.Department) AS Department, " & _
        "Trim(tblStaffEvalsAndSupervisions.Loc) AS LocationName, Trim(tblStaffEvalsAndSupervisions.SupervisorCode) AS SupervisorCode, Notes, DeleteFlag " & _
        "INTO temptbl4 FROM tblStaffEvalsAndSupervisions;", dbSeeChanges: Call BriefDelay
    
    Call AppendProgressMessages("Add Supervisons.")
    Call BriefDelay
    TempDB.Execute "SELECT temptbl4.EmployeeID AS EMPID_I, 99 AS SKILLNUMBER_I, temptbl4.LastSupervisionDate AS EXPIREDSKILL_I, " & _
        "StrConv (temptbl4.LASTNAME, 3) AS LASTNAME, StrConv (temptbl4.FirstName, 3) AS FRSTNAME, temptbl4.LocationName, " & _
        "Trim(temptbl4.SupervisorCode) AS SUPERVISORCODE_I, temptbl3.INDEXEDNAME AS SupervisorIndexedName, False AS RedFlag, temptbl4.OnLeave AS OnLeave " & _
        "INTO temptbl5 FROM temptbl4 INNER JOIN temptbl3 ON temptbl4.SupervisorCode = temptbl3.SUPERVISORCODE_I;", dbSeeChanges: Call BriefDelay
    
    Call AppendProgressMessages("Add Evals.")
    Call BriefDelay
    TempDB.Execute "INSERT INTO temptbl5 (EMPID_I, SKILLNUMBER_I, LASTNAME, FRSTNAME, LocationName, SUPERVISORCODE_I, SupervisorIndexedName, EXPIREDSKILL_I, RedFlag, OnLeave) " & _
        "SELECT temptbl4.EmployeeID, 98 AS SKILLNUMBER_I, StrConv (temptbl4.LastName, 3) AS LASTNAME, StrConv (temptbl4.FirstName, 3) AS FRSTNAME, " & _
        "Trim(temptbl4.LocationName) AS LocationName, Trim(temptbl4.SupervisorCode) AS SUPERVISORCODE_I, temptbl3.INDEXEDNAME, temptbl4.EvalDueBy AS EXPIREDSKILL_I, " & _
        "False AS RedFlag, temptbl4.OnLeave FROM temptbl4 INNER JOIN temptbl3 ON temptbl4.SupervisorCode = temptbl3.SUPERVISORCODE_I;", dbSeeChanges: Call BriefDelay
    
    Call AppendProgressMessages("Update Supervisor Names for Supervisions and Evals.")
    Call AppendProgressMessages("   This may take a few minutes.")
    Call BriefDelay
    TempDB.Execute "UPDATE (temptbl5 INNER JOIN tblStaff ON temptbl5.EMPID_I = tblStaff.EMPLOYID) " & _
        "INNER JOIN temptbl3 ON tblStaff.SUPERVISORCODE_I = temptbl3.SUPERVISORCODE_I " & _
        "SET temptbl5.LASTNAME = StrConv(temptbl5.LASTNAME,3), temptbl5.FRSTNAME = StrConv(temptbl5.FRSTNAME,3), temptbl5.SUPERVISORCODE_I = tblStaff.SUPERVISORCODE_I, " & _
        "temptbl5.SupervisorIndexedName = temptbl3.INDEXEDNAME;", dbSeeChanges: Call BriefDelay

    Call AppendProgressMessages("Insert staff evaluations into temporary staff skills table.")
    Call AppendProgressMessages("   This may take a few minutes.")
    Call BriefDelay
    TempDB.Execute "INSERT INTO temptbl1 (EMPID_I, SKILLNUMBER_I, SKILL, EXPIREDSKILL_I, LASTNAME, FRSTNAME, LocationName, SUPERVISORCODE_I, SupervisorIndexedName, RedFlag, OnLeave) " & _
        "SELECT temptbl5.EMPID_I, temptbl5.SKILLNUMBER_I, '' AS Skill, temptbl5.EXPIREDSKILL_I, temptbl5.LASTNAME, " & _
        "temptbl5.FRSTNAME, temptbl5.LocationName, temptbl5.SUPERVISORCODE_I, temptbl5.SupervisorIndexedName, temptbl5.RedFlag, " & _
        "temptbl5.OnLeave " & _
        "FROM temptbl5;", dbSeeChanges: Call BriefDelay
    
    Call AppendProgressMessages("Finalizing staff skills table for reporting.")
    Call BriefDelay
    TempDB.Execute "SELECT temptbl1.EMPID_I, temptbl1.SKILLNUMBER_I, temptbl1.Skill, temptbl1.EXPIREDSKILL_I, temptbl1.LASTNAME, " & _
        "temptbl1.FRSTNAME, temptbl1.LocationName AS LocationName, temptbl1.SUPERVISORCODE_I, temptbl1.SupervisorIndexedName, """" AS SupervisorName, temptbl1.RedFlag, temptbl1.OnLeave " & _
        "INTO temptbl6 FROM temptbl1 " & _
        "ORDER BY temptbl1.LASTNAME, temptbl1.FRSTNAME, temptbl1.SKILLNUMBER_I;", dbSeeChanges: Call BriefDelay
    
    Call AppendProgressMessages("Delete extraneous records.")
    Call BriefDelay
    TempDB.Execute "DELETE temptbl6.* FROM temptbl6 WHERE temptbl6.EMPID_I = 'CLUS9234';", dbSeeChanges: Call BriefDelay

    Call AppendProgressMessages("Update on-leave flag in the staff skills table where appropriate.")
    Call BriefDelay
    TempDB.Execute "SELECT EMPID_I, OnLeave, SKILLNUMBER_I " & _
        "INTO temptbl7 FROM temptbl6 WHERE (OnLeave = True AND (SKILLNUMBER_I = 99 OR SKILLNUMBER_I = 98));", dbSeeChanges: Call BriefDelay
    Call BriefDelay
    TempDB.Execute "UPDATE temptbl6 INNER JOIN temptbl7 ON temptbl6.EMPID_I = temptbl7.EMPID_I SET temptbl6.OnLeave = True WHERE temptbl6.EMPID_I = temptbl7.EMPID_I;"

    Call AppendProgressMessages("Calculate red flags and set the red flag field where appropriate.")
    Call BriefDelay
    TempDB.Execute "UPDATE temptbl6 SET REDFLAG = True " & _
        "WHERE (temptbl6.SKILLNUMBER_I=1   And DateDiff('d',Date(),EXPIREDSKILL_I)< " & Trig_Staff_CPR_Red & " ) OR " & _
              "(temptbl6.SKILLNUMBER_I=2   And DateDiff('d',Date(),EXPIREDSKILL_I)< " & Trig_Staff_FA_Red & " ) OR " & _
              "(temptbl6.SKILLNUMBER_I=3   And DateDiff('d',Date(),EXPIREDSKILL_I)< " & Trig_Staff_MAP_Red & " ) OR " & _
              "(temptbl6.SKILLNUMBER_I=12  And DateDiff('d',Date(),EXPIREDSKILL_I)< " & Trig_Staff_PL_Red & " ) OR " & _
              "(temptbl6.SKILLNUMBER_I=15  And DateDiff('d',Date(),EXPIREDSKILL_I)< " & Trig_Staff_DL_Red & " ) OR " & _
              "(temptbl6.SKILLNUMBER_I=22  And DateDiff('d',Date(),EXPIREDSKILL_I)< " & Trig_Staff_BBP_Red & " ) OR " & _
              "(temptbl6.SKILLNUMBER_I=30  And DateDiff('d',Date(),EXPIREDSKILL_I)< " & Trig_Staff_SC_Red & " ) OR " & _
              "(temptbl6.SKILLNUMBER_I=31  And DateDiff('d',Date(),EXPIREDSKILL_I)< " & Trig_Staff_WV_Red & " ) OR " & _
              "(temptbl6.SKILLNUMBER_I=35  And DateDiff('d',Date(),EXPIREDSKILL_I)< " & Trig_Staff_PL_Red & " ) OR " & _
              "(temptbl6.SKILLNUMBER_I=36  And DateDiff('d',Date(),EXPIREDSKILL_I)< " & Trig_Staff_TB_Red & " ) OR " & _
              "(temptbl6.SKILLNUMBER_I=39  And DateDiff('d',Date(),EXPIREDSKILL_I)< " & Trig_Staff_BIP_Red & ");", dbSeeChanges: Call BriefDelay
    Call BriefDelay
    TempDB.Execute "UPDATE temptbl6 SET RedFlag=True " & _
        "WHERE EXPIREDSKILL_I < Date() AND EXPIREDSKILL_I <> DateValue('01/01/1900') AND (SKILLNUMBER_I = 99 OR SKILLNUMBER_I = 98);", dbSeeChanges: Call BriefDelay
      
    Call BriefDelay
    TempDB.Execute "UPDATE temptbl6 SET RedFlag=True WHERE FieldIsEmpty(EXPIREDSKILL_I) AND SKILLNUMBER_I = 99;", dbSeeChanges: Call BriefDelay
        
    Call AppendProgressMessages("Translate the skills code into text.")
    Call BriefDelay
    TempDB.Execute "SELECT catSkills.SkillID, catSkills.SKILL INTO temptbl9 FROM catSkills;", dbSeeChanges: Call BriefDelay
    Call BriefDelay
    TempDB.Execute "UPDATE temptbl6 INNER JOIN temptbl9 ON temptbl6.SKILLNUMBER_I = temptbl9.SkillID SET temptbl6.Skill = [temptbl9].[SKILL];", dbSeeChanges: Call BriefDelay
    Call BriefDelay
    TempDB.Execute "UPDATE temptbl6 INNER JOIN temptbl3 ON temptbl6.SupervisorIndexedName = temptbl3.INDEXEDNAME " & _
        "SET temptbl6.SupervisorName = [temptbl3].[FRSTNAME] & ' ' & [temptbl3].[LASTNAME];", dbSeeChanges: Call BriefDelay
    
    Call AppendProgressMessages("Creating final consolidated table for report generation.")
    Call BriefDelay
    TempDB.Execute "INSERT INTO RedReport (EMPID_I, SKILLNUMBER_I, Skill, EXPIREDSKILL_I, LASTNAME, FRSTNAME, StaffName, LocationName, SUPERVISORCODE_I, SupervisorIndexedName, SupervisorName, RedFlag, OnLeave) " & _
        "SELECT temptbl6.EMPID_I, temptbl6.SKILLNUMBER_I, temptbl6.Skill, temptbl6.EXPIREDSKILL_I, temptbl6.LASTNAME, temptbl6.FRSTNAME, [temptbl6].[FRSTNAME] & ' ' & [temptbl6].[LASTNAME] AS StaffName, " & _
        "temptbl6.LocationName, temptbl6.SUPERVISORCODE_I, temptbl6.SupervisorIndexedName, temptbl6.SupervisorName, temptbl6.RedFlag, temptbl6.OnLeave " & _
        "FROM temptbl6 WHERE temptbl6.RedFlag = True;", dbSeeChanges: Call BriefDelay

    Call BriefDelay
    If SupervisionsOnly Then
        Call AppendProgressMessages("This report will show only Last Supervision expirations.")
        TempDB.Execute "SELECT RedReport.* INTO temptbl8 FROM RedReport WHERE RedReport.SKILLNUMBER_I = 99;", dbSeeChanges: Call BriefDelay
    Else
        Call AppendProgressMessages("This report will show all expirations except Last Supervisions.")
        TempDB.Execute "SELECT RedReport.* INTO temptbl8 FROM RedReport WHERE RedReport.SKILLNUMBER_I <> 99;", dbSeeChanges: Call BriefDelay
    End If

    Call AppendProgressMessages("Creating Red Report.")
    Call BriefDelay
    DoCmd.OpenReport "rptEXPIRATIONDATESredreport", acViewPreview
    Call LoopUntilClosed("rptEXPIRATIONDATESredreport", acReport)
    
    Call AppendProgressMessages("Clean-up before exiting.")
    Call AppendProgressMessages("  Do not navigate from this page until all status messages are removed.")
    Call BriefDelay
    TempDB.Execute "DELETE RedReport.* FROM RedReport;", dbSeeChanges: Call BriefDelay
    Call BriefDelay(1)
    If IsObjectOpen("temptbl9", acTable) Then DoCmd.Close acTable, "temptbl9"
    If IsTableQuery("temptbl9") Then TempDB.Execute "DROP TABLE temptbl9", dbSeeChanges: Call BriefDelay
    Call BriefDelay(1)
    If IsObjectOpen("temptbl8", acTable) Then DoCmd.Close acTable, "temptbl8"
    If IsTableQuery("temptbl8") Then TempDB.Execute "DROP TABLE temptbl8", dbSeeChanges: Call BriefDelay
    Call BriefDelay(1)
    If IsObjectOpen("temptbl7", acTable) Then DoCmd.Close acTable, "temptbl7"
    If IsTableQuery("temptbl7") Then TempDB.Execute "DROP TABLE temptbl7", dbSeeChanges: Call BriefDelay
    Call BriefDelay(1)
    If IsObjectOpen("temptbl6", acTable) Then DoCmd.Close acTable, "temptbl6"
    If IsTableQuery("temptbl6") Then TempDB.Execute "DROP TABLE temptbl6", dbSeeChanges: Call BriefDelay
    Call BriefDelay(1)
    If IsObjectOpen("temptbl5", acTable) Then DoCmd.Close acTable, "temptbl5"
    If IsTableQuery("temptbl5") Then TempDB.Execute "DROP TABLE temptbl5", dbSeeChanges: Call BriefDelay
    Call BriefDelay(1)
    If IsObjectOpen("temptbl4", acTable) Then DoCmd.Close acTable, "temptbl4"
    If IsTableQuery("temptbl4") Then TempDB.Execute "DROP TABLE temptbl4", dbSeeChanges: Call BriefDelay
    Call BriefDelay(1)
    If IsObjectOpen("temptbl3", acTable) Then DoCmd.Close acTable, "temptbl3"
    If IsTableQuery("temptbl3") Then TempDB.Execute "DROP TABLE temptbl3", dbSeeChanges: Call BriefDelay
    Call BriefDelay(1)
    If IsObjectOpen("temptbl2", acTable) Then DoCmd.Close acTable, "temptbl2"
    If IsTableQuery("temptbl2") Then TempDB.Execute "DROP TABLE temptbl2", dbSeeChanges: Call BriefDelay
    Call BriefDelay(1)
    If IsObjectOpen("temptbl1", acTable) Then DoCmd.Close acTable, "temptbl1"
    If IsTableQuery("temptbl1") Then TempDB.Execute "DROP TABLE temptbl1", dbSeeChanges: Call BriefDelay
    Call BriefDelay(1)
    If IsObjectOpen("temptbl0", acTable) Then DoCmd.Close acTable, "temptbl0"
    If IsTableQuery("temptbl0") Then TempDB.Execute "DROP TABLE temptbl0", dbSeeChanges: Call BriefDelay
    Call BriefDelay(1)
    If IsObjectOpen("temptbl", acTable) Then DoCmd.Close acTable, "temptbl"
    If IsTableQuery("temptbl") Then TempDB.Execute "DROP TABLE temptbl", dbSeeChanges: Call BriefDelay

    Form_frmRpt.ProgressMessages = ""
    Form_frmRpt.ProgressMessages.Requery
    DoCmd.SetWarnings True
    RunRedReportNew = True
    Set TempDB = Nothing
    Exit Function
ShowMeError:
    Set TempDB = Nothing
    Call DropTempTables
    Form_frmRpt.ProgressMessages = "": Form_frmRpt.ProgressMessages.Requery
    MsgBox "Error # " & Str(Err.Number) & " was generated by RunRedReport" & Chr(13) & Err.Description, vbOKOnly, "Error", Err.HelpFile, Err.HelpContext
    RunRedReportNew = False
End Function

Public Sub SetExpirationFieldProperties(Field As Label, Optional Criteria As String, Optional CalcField As Boolean = False, Optional CalcCriteria As Date)
    If CalcField Then
        Select Case CalcCriteria
            Case ExpMissingCalculated
                Field.Visible = True: Field.Caption = "Missing": Field.ForeColor = RGB(255, 0, 0): Field.FontWeight = 700: Field.BorderStyle = 1: Field.BorderColor = RGB(255, 0, 0)
            Case ExpOptionalCalculated
                Field.Visible = True: Field.Caption = "Optional": Field.ForeColor = RGB(0, 0, 0): Field.FontWeight = 400: Field.BorderStyle = 0: Field.BorderColor = RGB(0, 0, 0)
            Case ExpNACalculated
                Field.Visible = True: Field.Caption = "N/A": Field.ForeColor = RGB(0, 0, 0): Field.FontWeight = 400: Field.BorderStyle = 0: Field.BorderColor = RGB(0, 0, 0)
            Case Else
        End Select
    Else
        Select Case Format(Criteria, "YYYY-MM-DD")
            Case ExpMissing
                Field.Visible = True: Field.Caption = "Missing": Field.ForeColor = RGB(255, 0, 0): Field.FontWeight = 700: Field.BorderStyle = 1: Field.BorderColor = RGB(255, 0, 0)
            Case ExpOptional
                Field.Visible = True: Field.Caption = "Optional": Field.ForeColor = RGB(0, 0, 0): Field.FontWeight = 400: Field.BorderStyle = 0: Field.BorderColor = RGB(0, 0, 0)
            Case ExpNA
                Field.Visible = True: Field.Caption = "N/A": Field.ForeColor = RGB(0, 0, 0): Field.FontWeight = 400: Field.BorderStyle = 0: Field.BorderColor = RGB(0, 0, 0)
            Case ExpCompleted
                Field.Visible = True: Field.Caption = "Completed": Field.ForeColor = RGB(0, 0, 0): Field.FontWeight = 400: Field.BorderStyle = 0: Field.BorderColor = RGB(0, 0, 0)
            Case ExpPending
                Field.Visible = True: Field.Caption = "Pending": Field.ForeColor = RGB(0, 0, 0): Field.FontWeight = 400: Field.BorderStyle = 0: Field.BorderColor = RGB(0, 0, 0)
            Case Else
        End Select
    End If
End Sub

Public Sub ExpirationFieldMissing(Field As Label)
    Field.Visible = True: Field.Caption = "Missing": Field.ForeColor = RGB(255, 0, 0): Field.FontWeight = 700: Field.BorderStyle = 1: Field.BorderColor = RGB(255, 0, 0)
End Sub

Public Sub ExpirationFieldOptional(Field As Label)
    Field.Visible = True: Field.Caption = "Optional": Field.ForeColor = RGB(0, 0, 0): Field.FontWeight = 400: Field.BorderStyle = 0: Field.BorderColor = RGB(0, 0, 0)
End Sub

Public Sub ExpirationFieldNA(Field As Label)
    Field.Visible = True: Field.Caption = "N/A": Field.ForeColor = RGB(0, 0, 0): Field.FontWeight = 400: Field.BorderStyle = 0: Field.BorderColor = RGB(0, 0, 0)
End Sub

Public Sub AppendProgressMessages(Text As Variant)
    Form_frmRpt.ProgressMessages = Form_frmRpt.ProgressMessages & vbCrLf & Text
    Form_frmRpt.ProgressMessages.Requery
    Call BriefDelay(1)
End Sub

Public Sub ExpirationsHousekeeping(DB As Database, Step As Integer)
    Select Case Step
        Case 1
            If IsObjectOpen("tempstaff", acTable) Then DoCmd.Close acTable, "tempstaff"
            If IsTableQuery("tempstaff") Then DB.Execute "DROP TABLE tempstaff", dbSeeChanges: Call BriefDelay
        Case 2
            If IsObjectOpen("tempstaffskills", acTable) Then DoCmd.Close acTable, "tempstaffskills"
            If IsTableQuery("tempstaffskills") Then DB.Execute "DROP TABLE tempstaffskills", dbSeeChanges: Call BriefDelay
    End Select
End Sub
