-- Query Name: qryUpdateSpringboardLeaders01
-- Extracted: 2026-02-04 13:04:22

UPDATE tblPeopleClientsSpringboardServices SET tblPeopleClientsSpringboardServices.GroupCode = Null, tblPeopleClientsSpringboardServices.LeaderIndexedName = Null, tblPeopleClientsSpringboardServices.Leader = Null
WHERE (((tblPeopleClientsSpringboardServices.GroupCode) Is Not Null) AND ((tblPeopleClientsSpringboardServices.Inactive)=True));

