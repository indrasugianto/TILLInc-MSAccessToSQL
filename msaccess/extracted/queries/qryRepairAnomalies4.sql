-- Query Name: qryRepairAnomalies4
-- Extracted: 2026-02-04 13:04:22

UPDATE (tblPeople INNER JOIN tblPeopleClientsCLOServices ON tblPeople.IndexedName = tblPeopleClientsCLOServices.IndexedName) INNER JOIN tblPeopleClientsVendors ON tblPeople.IndexedName = tblPeopleClientsVendors.IndexedName SET tblPeopleClientsVendors.LivingIndependently = False
WHERE (((tblPeopleClientsVendors.LivingIndependently)=True) AND ((tblPeople.IsCilentCLO)=True) AND ((tblPeopleClientsCLOServices.Inactive)=False));

