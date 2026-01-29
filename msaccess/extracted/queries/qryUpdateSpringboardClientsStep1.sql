-- Query Name: qryUpdateSpringboardClientsStep1
-- Extracted: 2026-01-29 16:09:05

UPDATE tblPeopleClientsSpringboardServices SET tblPeopleClientsSpringboardServices.GroupCode = Null, tblPeopleClientsSpringboardServices.LeaderIndexedName = Null, tblPeopleClientsSpringboardServices.Leader = Null
WHERE (((tblPeopleClientsSpringboardServices.GroupCode) Is Not Null) AND ((tblPeopleClientsSpringboardServices.Inactive)=True));

