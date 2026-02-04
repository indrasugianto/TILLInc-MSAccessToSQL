-- Query Name: qryEXPIRATIONS14
-- Extracted: 2026-02-04 13:04:22

INSERT INTO tblExpirations ( Location, RecordType, LastName, FirstName, Supervisor, LastVehicleChecklistCompleted, MostRecentAsleepFireDrill, NextRecentAsleepFireDrill, HouseSafetyPlanExpires, HousePlansReviewedByStaffBefore, DAYStaffTrainedInPrivacyBefore, DAYAllPlansReviewedByStaffBefore, DAYQtrlySafetyChecklistDueBy, MAPChecklistCompleted, HumanRightsOfficer, HROTrainsStaffBefore, HROTrainsIndividualsBefore, FireSafetyOfficer, FSOTrainsStaffBefore, FSOTrainsIndividualsBefore )
SELECT tblLocations.GPName AS Location, "House" AS RecordType, "*" AS LastName, "*" AS FirstName, DLookUp("GPSuperCode","temptbl","GPName='" & tblLocations.GPName & "'") AS Supervisor, tblLocations.LastVehicleChecklistCompleted, tblLocations.MostRecentAsleepFireDrill, tblLocations.NextRecentAsleepFireDrill, tblLocations.HouseSafetyPlanExpires, tblLocations.HousePlansReviewedByStaffBefore, tblLocations.DAYStaffTrainedInPrivacyBefore, tblLocations.DAYAllPlansReviewedByStaffBefore, tblLocations.DAYQtrlySafetyChecklistDueBy, tblLocations.MAPChecklistCompleted, tblLocations.HumanRightsOfficer, tblLocations.HROTrainsStaffBefore, tblLocations.HROTrainsIndividualsBefore, tblLocations.FireSafetyOfficer, tblLocations.FSOTrainsStaffBefore, tblLocations.FSOTrainsIndividualsBefore
FROM tblLocations
WHERE (((tblLocations.GPName) Is Not Null) And ((DLookUp("GPSuperCode","temptbl","GPName='" & tblLocations.GPName & "'")) Is Not Null) And ((tblLocations.Department)<>"Clinical and Support Services"))
ORDER BY tblLocations.GPName;

