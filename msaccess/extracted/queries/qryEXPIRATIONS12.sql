-- Query Name: qryEXPIRATIONS12
-- Extracted: 2026-02-04 13:04:22

UPDATE (temptbl0 LEFT JOIN qrytblPeopleClientsDemographics ON temptbl0.IndexedName = qrytblPeopleClientsDemographics.IndexedName) LEFT JOIN tblPeopleClientsDayServices ON temptbl0.IndexedName = tblPeopleClientsDayServices.IndexedName SET temptbl0.LocDay = Null
WHERE (((Len(temptbl0!LocDay))>0) And ((qrytblPeopleClientsDemographics.ActiveDayServices)=False)) Or (((Len(temptbl0!LocDay))>0) And ((qrytblPeopleClientsDemographics.ActiveDayServices)=True) And ((tblPeopleClientsDayServices.Inactive)=True));

