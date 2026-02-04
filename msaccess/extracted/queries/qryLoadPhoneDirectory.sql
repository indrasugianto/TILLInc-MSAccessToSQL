-- Query Name: qryLoadPhoneDirectory
-- Extracted: 2026-02-04 13:04:22

INSERT INTO tblPhoneDirectory ( Department, Location, LocationDetail, LastName, FirstName, EmailAddress, JobTitle, InternalExtension, HasPhoneOnDesktop, ExternalPhoneNumber )
SELECT tblPeople.Department, tblPeople.OfficeCityTown AS Location, tblPeople.OfficeLocationName AS LocationDetail, tblPeople.LastName, tblPeople.FirstName, tblPeople.EmailAddress, tblPeople.StaffTitle AS JobTitle, tblPeople.DID AS InternalExtension, tblPeople.HasPhoneOnDesktop, tblPeople.StaffExtPhone AS ExternalPhoneNumber
FROM tblPeople
WHERE (((tblPeople.IsStaff)=True))
ORDER BY tblPeople.Department, tblPeople.OfficeCityTown, tblPeople.OfficeLocationName, tblPeople.LastName, tblPeople.FirstName, IIf([Department]='Residential Services',1,0), tblPeople.Department, tblPeople.LastName, tblPeople.FirstName;

