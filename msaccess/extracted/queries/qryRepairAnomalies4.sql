-- Query Name: qryRepairAnomalies4
-- Extracted: 2026-01-29 16:09:05

UPDATE (tblPeople INNER JOIN tblPeopleClientsCLOServices ON tblPeople.IndexedName = tblPeopleClientsCLOServices.IndexedName) INNER JOIN tblPeopleClientsVendors ON tblPeople.IndexedName = tblPeopleClientsVendors.IndexedName SET tblPeopleClientsVendors.LivingIndependently = False
WHERE (((tblPeopleClientsVendors.LivingIndependently)=True) AND ((tblPeople.IsCilentCLO)=True) AND ((tblPeopleClientsCLOServices.Inactive)=False));

