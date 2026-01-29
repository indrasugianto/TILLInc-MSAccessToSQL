-- Query Name: qryEXPIRATIONS23
-- Extracted: 2026-01-29 16:09:05

UPDATE qrytblExpirations INNER JOIN qrytblStaffEvalsAndSupervisions ON (qrytblExpirations.LastName = qrytblStaffEvalsAndSupervisions.LastName) AND (qrytblExpirations.FirstName = qrytblStaffEvalsAndSupervisions.FirstName) SET qrytblExpirations.ThreeMonthEvaluation = [qrytblStaffEvalsAndSupervisions]![ThreeMonthEval], qrytblExpirations.EvalDueBy = [qrytblStaffEvalsAndSupervisions]![EvalDueBy], qrytblExpirations.LastSupervision = [qrytblStaffEvalsAndSupervisions]![LastSupervision], qrytblExpirations.OnLeave = [qrytblStaffEvalsAndSupervisions]![OnLeave]
WHERE qrytblExpirations.RecordType="Staff";

