-- Query Name: qryEXPIRATIONS11
-- Extracted: 2026-02-04 13:04:22

UPDATE (temptbl0 LEFT JOIN qrytblPeopleClientsDemographics ON temptbl0.IndexedName = qrytblPeopleClientsDemographics.IndexedName) LEFT JOIN tblPeopleClientsResidentialServices ON temptbl0.IndexedName = tblPeopleClientsResidentialServices.IndexedName SET temptbl0.LocRes = Null
WHERE (((Len(temptbl0!LocRes))>0) And ((qrytblPeopleClientsDemographics.ActiveResidentialServices)=False)) Or (((Len(temptbl0!LocRes))>0) And ((qrytblPeopleClientsDemographics.ActiveResidentialServices)=True) And ((tblPeopleClientsResidentialServices.Inactive)=True));

