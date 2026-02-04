-- Query Name: qryDonorAppealCreateMostRecentDonations
-- Extracted: 2026-02-04 13:04:22

SELECT temptbl0.IndexedName, Max(Int(CDbl(DateValue([temptbl0].[DateOfDonation])))) AS LastDonationNumeric, CVDate(Max(Int(CDbl(DateValue([temptbl0].[DateOfDonation]))))) AS LastDonationDate, Format(CDate(Max(Int(CDbl(DateValue([temptbl0].[DateOfDonation]))))),"mm/dd/yyyy") AS FormattedDate, DLookUp("Amount","temptbl0","DateOfDonation=""" & Format(CDate(Max(Int(CDbl(DateValue([temptbl0].[DateOfDonation]))))),"mm/dd/yyyy") & """ AND IndexedName = """ & [IndexedName] & """") AS LastDonationAmount, IIf(Year(CVDate(Max(Int(CDbl(DateValue([temptbl0].[DateOfDonation]))))))>=Year(Now())-1,"Current","Lapsed") AS CurrentOrLapsed INTO tmpMostRecentDonations
FROM temptbl0
GROUP BY temptbl0.IndexedName
ORDER BY temptbl0.IndexedName, Max(Int(CDbl(DateValue([temptbl0].[DateOfDonation])))) DESC;

