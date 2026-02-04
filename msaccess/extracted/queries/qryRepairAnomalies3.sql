-- Query Name: qryRepairAnomalies3
-- Extracted: 2026-02-04 13:04:22

UPDATE (tblPeople INNER JOIN tblPeopleClientsResidentialServices ON tblPeople.IndexedName = tblPeopleClientsResidentialServices.IndexedName) INNER JOIN tblPeopleClientsVendors ON tblPeople.IndexedName = tblPeopleClientsVendors.IndexedName SET tblPeopleClientsVendors.LivingIndependently = False
WHERE (((tblPeopleClientsVendors.LivingIndependently)=True) AND ((tblPeople.IsClientRes)=True) AND ((tblPeopleClientsResidentialServices.Inactive)=False));

