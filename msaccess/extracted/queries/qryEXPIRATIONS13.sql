-- Query Name: qryEXPIRATIONS13
-- Extracted: 2026-01-29 16:09:05

UPDATE (temptbl0 LEFT JOIN qrytblPeopleClientsDemographics ON temptbl0.IndexedName = qrytblPeopleClientsDemographics.IndexedName) LEFT JOIN tblPeopleClientsVocationalServices ON temptbl0.IndexedName = tblPeopleClientsVocationalServices.IndexedName SET temptbl0.LocVoc = Null
WHERE (((Len(temptbl0!LocVoc))>0) And ((qrytblPeopleClientsDemographics.ActiveVocationalServices)=False)) Or (((Len(temptbl0!LocVoc))>0) And ((qrytblPeopleClientsDemographics.ActiveVocationalServices)=True) And ((tblPeopleClientsVocationalServices.Inactive)=True));

