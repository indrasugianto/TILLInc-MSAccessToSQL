-- Query Name: qryEXPIRATIONS10
-- Extracted: 2026-02-04 13:04:22

UPDATE (temptbl0 LEFT JOIN qrytblPeopleClientsDemographics ON temptbl0.IndexedName = qrytblPeopleClientsDemographics.IndexedName) LEFT JOIN tblPeopleClientsCLOServices ON temptbl0.IndexedName = tblPeopleClientsCLOServices.IndexedName SET temptbl0.LocCLO = Null
WHERE (Len(temptbl0!LocCLO)>0 And qrytblPeopleClientsDemographics.ActiveCLO=False) Or (Len(temptbl0!LocCLO)>0 And qrytblPeopleClientsDemographics.ActiveCLO=True And tblPeopleClientsCLOServices.Inactive=True);

