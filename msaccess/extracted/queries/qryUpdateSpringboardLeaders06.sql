-- Query Name: qryUpdateSpringboardLeaders06
-- Extracted: 2026-01-29 16:09:05

UPDATE tblPeopleClientsSpringboardServices INNER JOIN temptbl ON tblPeopleClientsSpringboardServices.GroupCode = temptbl.SprGroup SET tblPeopleClientsSpringboardServices.LeaderIndexedName = [temptbl]![IndexedName], tblPeopleClientsSpringboardServices.Leader = [temptbl]![FirstName] & ' ' & [temptbl]![LastName];

