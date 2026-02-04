-- Query Name: qryUpdateSpringboardLeaders06
-- Extracted: 2026-02-04 13:04:22

UPDATE tblPeopleClientsSpringboardServices INNER JOIN temptbl ON tblPeopleClientsSpringboardServices.GroupCode = temptbl.SprGroup SET tblPeopleClientsSpringboardServices.LeaderIndexedName = [temptbl]![IndexedName], tblPeopleClientsSpringboardServices.Leader = [temptbl]![FirstName] & ' ' & [temptbl]![LastName];

