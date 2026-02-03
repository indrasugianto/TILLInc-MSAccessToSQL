# Documentation Cleanup Summary

**Date:** January 30, 2026 · **Last updated:** February 2, 2026

Redundant optimization docs were removed; field-addition guides were replaced by the VBA module `ModReportFieldManager.vba`. Current structure:

- **Expiration optimization:** Start at [EXPIRATION_REPORT_OPTIMIZATION_README.md](EXPIRATION_REPORT_OPTIMIZATION_README.md)
- **Full index:** [.docs/README.md](README.md)

**Removed (Jan 2026):** ADD_FIELDS_*.md, QUICK_START_CHECKLIST.md, FIX_FIELD_NOT_FOUND_ERROR.md, PERFORMANCE_OPTIMIZATION_SUMMARY.md, list_view_fields.sql (7 items).

**Removed (Feb 2026):** README_OPTIMIZATION.md (merged into EXPIRATION_REPORT_OPTIMIZATION_README), FINAL_IMPLEMENTATION_STATUS.md (superseded by main optimization README).

**Added (Feb 2026):** Staff subreport (FixStaffReport, vw_ExpirationsFormatted Staff columns, rptEXPIRATIONDATESstaff_VBA_UPDATED.md); vw_ExpirationsFormatted_REASSESSMENT.md (view vs. report VBA, NULL fix, checklist); RunExpirationsReport_SP_Assessment.md in index. View fix: ClusterDescription removed (column does not exist in tblLocations); NULL handling fixed in view (IN (..., null) does not match NULL in T-SQL — replaced with explicit OR col IS NULL).

**Reconciled (Feb 2026):** All .docs references updated; expiration optimization count set to 7 files; RunExpirationsReport section to 5 files; list_view_fields.sql references removed (file was already removed).
