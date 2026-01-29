-- Query Name: qryEXPIRATIONS22
-- Extracted: 2026-01-29 16:09:05

UPDATE qrytblExpirations INNER JOIN temptbl3 ON (qrytblExpirations.FirstName = temptbl3.FRSTNAME) AND (qrytblExpirations.LastName = temptbl3.LASTNAME) AND (qrytblExpirations.Location = temptbl3.DEPRTMNT) SET qrytblExpirations.CPR = [temptbl3].[CPR], qrytblExpirations.FirstAid = [temptbl3].[FirstAid], qrytblExpirations.MAPCert = [temptbl3].[MAPCert], qrytblExpirations.DriversLicense = [temptbl3].[DriversLicense], qrytblExpirations.BBP = [temptbl3].[BBP], qrytblExpirations.BackInjuryPrevention = [temptbl3].[BackInjuryPrevention], qrytblExpirations.SafetyCares = [temptbl3].[SafetyCares], qrytblExpirations.TB = [temptbl3].[TB], qrytblExpirations.WorkplaceViolence = [temptbl3].[WorkplaceViolence], qrytblExpirations.DefensiveDriving = [temptbl3].[DefensiveDriving], qrytblExpirations.WheelchairSafety = [temptbl3].[WheelchairSafety], qrytblExpirations.PBS = [temptbl3].[PBS], qrytblExpirations.ProfessionalLicenses = [temptbl3].[ProfLic]
WHERE (((qrytblExpirations.RecordType)="Staff"));

