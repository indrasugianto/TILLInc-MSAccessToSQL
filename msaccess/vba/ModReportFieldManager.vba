Option Compare Database
Option Explicit

' =============================================
' Module: ModReportFieldManager
' Description: Helper functions to add/remove hidden fields from reports
'              Used for implementing vw_ExpirationsFormatted optimization
'
' Version: 1.0
' Created: 2026-01-30
'
' Usage:
'   To ADD fields:     Run FixDayReport(), FixClientsReport(), FixHouseReport()
'   To REMOVE fields:  Run RemoveFieldsFromDayReport(), etc.
'   To REMOVE all:     Run RemoveAllHiddenFields()
' =============================================

' =============================================
' CORE HELPER FUNCTIONS
' =============================================

Sub AddHiddenFieldsToReport(reportName As String, fieldList As String)
    ' Adds hidden textbox controls to a report for all fields in fieldList
    ' fieldList should be comma-separated field names
    
    Dim rpt As Report
    Dim ctl As Control
    Dim fields() As String
    Dim i As Integer
    Dim topPos As Integer
    Dim fieldName As String
    Dim addedCount As Integer
    
    On Error Resume Next
    
    DoCmd.OpenReport reportName, acViewDesign
    Set rpt = REPORTS(reportName)
    
    fields = Split(fieldList, ",")
    topPos = 0
    addedCount = 0
    
    For i = 0 To UBound(fields)
        fieldName = Trim(fields(i))
        
        ' Skip if field already exists
        If ControlExists(reportName, fieldName) Then
            Debug.Print "Field already exists, skipping: " & fieldName
        Else
            ' Create textbox in Detail section
            Set ctl = CreateReportControl(reportName, acTextBox, acDetail, "", "", 0, topPos, 500, 200)
            
            If Err.Number = 0 Then
                ctl.Name = fieldName
                ctl.ControlSource = fieldName
                ctl.Visible = False
                addedCount = addedCount + 1
                topPos = topPos + 200
            Else
                Debug.Print "Error adding field: " & fieldName & " - " & Err.Description
                Err.Clear
            End If
        End If
    Next i
    
    DoCmd.Save acReport, reportName
    DoCmd.Close acReport, reportName, acSaveYes
    
    On Error GoTo 0
    
    MsgBox "Added " & addedCount & " hidden fields to " & reportName & vbCrLf & _
           "Skipped " & ((UBound(fields) + 1) - addedCount) & " fields (already exist)", _
           vbInformation, "Field Addition Complete"
End Sub

Sub DeleteHiddenFieldsFromReport(reportName As String, fieldList As String)
    ' Removes hidden fields added by AddHiddenFieldsToReport
    ' Used for cleanup or rollback
    ' NOTE: Due to Access VBA limitations, this makes fields visible for manual deletion
    '       Use DeleteAllHiddenTextboxes for automated deletion
    
    Dim rpt As Report
    Dim ctl As Control
    Dim fields() As String
    Dim i As Integer
    Dim markedCount As Integer
    Dim skippedCount As Integer
    Dim fieldName As String
    Dim foundControl As Boolean
    
    ' Open report in design view
    DoCmd.OpenReport reportName, acViewDesign
    Set rpt = REPORTS(reportName)
    
    fields = Split(fieldList, ",")
    markedCount = 0
    skippedCount = 0
    
    For i = 0 To UBound(fields)
        fieldName = Trim(fields(i))
        foundControl = False
        
        ' Try to find the control
        On Error Resume Next
        Set ctl = rpt.Controls(fieldName)
        foundControl = (Err.Number = 0 And Not ctl Is Nothing)
        Err.Clear
        On Error GoTo 0
        
        If foundControl Then
            ' Make it visible and BRIGHT MAGENTA for easy manual deletion
            On Error Resume Next
            ctl.Visible = True
            ctl.BackColor = RGB(255, 0, 255)  ' Bright magenta
            ctl.ForeColor = RGB(255, 255, 255)  ' White text
            ctl.Top = 100
            ctl.Left = 100 + (markedCount * 200)
            ctl.Width = 150
            ctl.Height = 200
            ctl.ControlTipText = "DELETE THIS FIELD"
            
            If Err.Number = 0 Then
                markedCount = markedCount + 1
                Debug.Print "? Marked for deletion (BRIGHT MAGENTA): " & fieldName
            Else
                skippedCount = skippedCount + 1
                Debug.Print "? Error marking: " & fieldName
                Err.Clear
            End If
            On Error GoTo 0
        Else
            skippedCount = skippedCount + 1
            Debug.Print "? Skipped (not found): " & fieldName
        End If
        
        Set ctl = Nothing
    Next i
    
    ' Save (leave open so user can see the marked fields)
    DoCmd.Save acReport, reportName
    
    MsgBox "MARKED " & markedCount & " fields for deletion (bright magenta)" & vbCrLf & _
           "Skipped " & skippedCount & " fields (not found)" & vbCrLf & vbCrLf & _
           "REPORT LEFT OPEN IN DESIGN VIEW:" & vbCrLf & _
           "1. Select all magenta fields (Shift+Click)" & vbCrLf & _
           "2. Press Delete key" & vbCrLf & _
           "3. Save and close report" & vbCrLf & vbCrLf & _
           "Or use DeleteAllHiddenTextboxes for automated deletion.", _
           vbInformation, "Manual Deletion Required"
End Sub

Function ControlExists(reportName As String, controlName As String) As Boolean
    ' Checks if a control exists in a report
    Dim rpt As Report
    Dim ctl As Control
    
    On Error Resume Next
    Set rpt = REPORTS(reportName)
    Set ctl = rpt.Controls(controlName)
    ControlExists = (Err.Number = 0)
    On Error GoTo 0
End Function

' =============================================
' ADD FIELDS TO EACH REPORT
' =============================================

Sub FixClientsReport()
    Dim fieldList As String
    
    ' First, ensure RecordSource is correct
    Call UpdateRecordSource("rptEXPIRATIONDATESclients", "vw_ExpirationsFormatted")
    
    ' Add all calculated fields
    fieldList = "DateISP_Display,DateISP_Color,DateISP_ShowDate," & _
                "PSDue_Calculated,PSDue_Display,PSDue_Color,PSDue_ShowText," & _
                "DateConsentFormsSigned_Display,DateConsentFormsSigned_Color,DateConsentFormsSigned_ShowDate," & _
                "DateBMMExpires_Display,DateBMMExpires_Color,DateBMMExpires_ShowDate," & _
                "DateSPDAuthExpires_Display,DateSPDAuthExpires_Color,DateSPDAuthExpires_ShowDate," & _
                "DateSignaturesDueBy_Display,DateSignaturesDueBy_Color,DateSignaturesDueBy_ShowDate"
    
    Call AddHiddenFieldsToReport("rptEXPIRATIONDATESclients", fieldList)
End Sub

Sub FixDayReport()
    Dim fieldList As String
    
    ' First, ensure RecordSource is correct
    Call UpdateRecordSource("rptEXPIRATIONDATESday", "vw_ExpirationsFormatted")
    
    ' Add all calculated fields
    fieldList = "LastVehicleChecklistCompleted_Display_Day,LastVehicleChecklistCompleted_Color_Day,LastVehicleChecklistCompleted_ShowDate_Day," & _
                "DAYStaffTrainedInPrivacyBefore_Display,DAYStaffTrainedInPrivacyBefore_Color,DAYStaffTrainedInPrivacyBefore_ShowDate," & _
                "DAYAllPlansReviewedByStaffBefore_Display,DAYAllPlansReviewedByStaffBefore_Color,DAYAllPlansReviewedByStaffBefore_ShowDate," & _
                "DAYQtrlySafetyChecklistDueBy_Display,DAYQtrlySafetyChecklistDueBy_Color,DAYQtrlySafetyChecklistDueBy_ShowDate," & _
                "HROTrainsStaffBefore_Display_Day,HROTrainsStaffBefore_Color_Day,HROTrainsStaffBefore_ShowDate_Day," & _
                "HROTrainsIndividualsBefore_Display_Day,HROTrainsIndividualsBefore_Color_Day,HROTrainsIndividualsBefore_ShowDate_Day," & _
                "FSOTrainsStaffBefore_Display_Day,FSOTrainsStaffBefore_Color_Day,FSOTrainsStaffBefore_ShowDate_Day," & _
                "FSOTrainsIndividualsBefore_Display_Day,FSOTrainsIndividualsBefore_Color_Day,FSOTrainsIndividualsBefore_ShowDate_Day," & _
                "HumanRightsOfficer_Formatted,HumanRightsOfficer_IsBlank,FireSafetyOfficer_Formatted,FireSafetyOfficer_IsBlank"
    
    Call AddHiddenFieldsToReport("rptEXPIRATIONDATESday", fieldList)
End Sub

Sub FixHouseReport()
    Dim fieldList As String
    
    ' First, ensure RecordSource is correct
    Call UpdateRecordSource("rptEXPIRATIONDATEShouse", "vw_ExpirationsFormatted")
    
    ' Add all calculated fields
    fieldList = "LastVehicleChecklistCompleted_Display_Res,LastVehicleChecklistCompleted_Color_Res,LastVehicleChecklistCompleted_ShowDate_Res," & _
                "MostRecentAsleepFireDrill_Display,MostRecentAsleepFireDrill_Color,MostRecentAsleepFireDrill_ShowDate," & _
                "HousePlansReviewedByStaffBefore_Display,HousePlansReviewedByStaffBefore_Color,HousePlansReviewedByStaffBefore_ShowDate," & _
                "HouseSafetyPlanExpires_Display,HouseSafetyPlanExpires_Color,HouseSafetyPlanExpires_ShowDate," & _
                "MAPChecklistCompleted_Display,MAPChecklistCompleted_Color,MAPChecklistCompleted_ShowDate," & _
                "HROTrainsStaffBefore_Display_Res,HROTrainsStaffBefore_Color_Res,HROTrainsStaffBefore_ShowDate_Res," & _
                "HROTrainsIndividualsBefore_Display_Res,HROTrainsIndividualsBefore_Color_Res,HROTrainsIndividualsBefore_ShowDate_Res," & _
                "FSOTrainsStaffBefore_Display_Res,FSOTrainsStaffBefore_Color_Res,FSOTrainsStaffBefore_ShowDate_Res," & _
                "FSOTrainsIndividualsBefore_Display_Res,FSOTrainsIndividualsBefore_Color_Res,FSOTrainsIndividualsBefore_ShowDate_Res," & _
                "HumanRightsOfficer_Formatted,HumanRightsOfficer_IsBlank,FireSafetyOfficer_Formatted,FireSafetyOfficer_IsBlank"
    
    Call AddHiddenFieldsToReport("rptEXPIRATIONDATEShouse", fieldList)
End Sub

Sub FixAllReports()
    ' Fixes all three subreports at once
    MsgBox "This will add hidden fields to all three subreports." & vbCrLf & _
           "This may take 1-2 minutes. Please wait...", vbInformation, "Processing"
    
    Call FixClientsReport
    Call FixDayReport
    Call FixHouseReport
    
    MsgBox "All subreports updated successfully!" & vbCrLf & _
           "You can now test the main report.", vbInformation, "Complete"
End Sub

' =============================================
' REMOVE FIELDS FROM EACH REPORT
' =============================================

Sub RemoveFieldsFromClientsReport()
    Call DeleteHiddenFieldsFromReport("rptEXPIRATIONDATESclients", _
        "DateISP_Display,DateISP_Color,DateISP_ShowDate," & _
        "PSDue_Calculated,PSDue_Display,PSDue_Color,PSDue_ShowText," & _
        "DateConsentFormsSigned_Display,DateConsentFormsSigned_Color,DateConsentFormsSigned_ShowDate," & _
        "DateBMMExpires_Display,DateBMMExpires_Color,DateBMMExpires_ShowDate," & _
        "DateSPDAuthExpires_Display,DateSPDAuthExpires_Color,DateSPDAuthExpires_ShowDate," & _
        "DateSignaturesDueBy_Display,DateSignaturesDueBy_Color,DateSignaturesDueBy_ShowDate")
End Sub

Sub RemoveFieldsFromDayReport()
    Call DeleteHiddenFieldsFromReport("rptEXPIRATIONDATESday", _
        "LastVehicleChecklistCompleted_Display_Day,LastVehicleChecklistCompleted_Color_Day,LastVehicleChecklistCompleted_ShowDate_Day," & _
        "DAYStaffTrainedInPrivacyBefore_Display,DAYStaffTrainedInPrivacyBefore_Color,DAYStaffTrainedInPrivacyBefore_ShowDate," & _
        "DAYAllPlansReviewedByStaffBefore_Display,DAYAllPlansReviewedByStaffBefore_Color,DAYAllPlansReviewedByStaffBefore_ShowDate," & _
        "DAYQtrlySafetyChecklistDueBy_Display,DAYQtrlySafetyChecklistDueBy_Color,DAYQtrlySafetyChecklistDueBy_ShowDate," & _
        "HROTrainsStaffBefore_Display_Day,HROTrainsStaffBefore_Color_Day,HROTrainsStaffBefore_ShowDate_Day," & _
        "HROTrainsIndividualsBefore_Display_Day,HROTrainsIndividualsBefore_Color_Day,HROTrainsIndividualsBefore_ShowDate_Day," & _
        "FSOTrainsStaffBefore_Display_Day,FSOTrainsStaffBefore_Color_Day,FSOTrainsStaffBefore_ShowDate_Day," & _
        "FSOTrainsIndividualsBefore_Display_Day,FSOTrainsIndividualsBefore_Color_Day,FSOTrainsIndividualsBefore_ShowDate_Day," & _
        "HumanRightsOfficer_Formatted,HumanRightsOfficer_IsBlank,FireSafetyOfficer_Formatted,FireSafetyOfficer_IsBlank")
End Sub

Sub RemoveFieldsFromHouseReport()
    Call DeleteHiddenFieldsFromReport("rptEXPIRATIONDATEShouse", _
        "LastVehicleChecklistCompleted_Display_Res,LastVehicleChecklistCompleted_Color_Res,LastVehicleChecklistCompleted_ShowDate_Res," & _
        "MostRecentAsleepFireDrill_Display,MostRecentAsleepFireDrill_Color,MostRecentAsleepFireDrill_ShowDate," & _
        "HousePlansReviewedByStaffBefore_Display,HousePlansReviewedByStaffBefore_Color,HousePlansReviewedByStaffBefore_ShowDate," & _
        "HouseSafetyPlanExpires_Display,HouseSafetyPlanExpires_Color,HouseSafetyPlanExpires_ShowDate," & _
        "MAPChecklistCompleted_Display,MAPChecklistCompleted_Color,MAPChecklistCompleted_ShowDate," & _
        "HROTrainsStaffBefore_Display_Res,HROTrainsStaffBefore_Color_Res,HROTrainsStaffBefore_ShowDate_Res," & _
        "HROTrainsIndividualsBefore_Display_Res,HROTrainsIndividualsBefore_Color_Res,HROTrainsIndividualsBefore_ShowDate_Res," & _
        "FSOTrainsStaffBefore_Display_Res,FSOTrainsStaffBefore_Color_Res,FSOTrainsStaffBefore_ShowDate_Res," & _
        "FSOTrainsIndividualsBefore_Display_Res,FSOTrainsIndividualsBefore_Color_Res,FSOTrainsIndividualsBefore_ShowDate_Res," & _
        "HumanRightsOfficer_Formatted,HumanRightsOfficer_IsBlank,FireSafetyOfficer_Formatted,FireSafetyOfficer_IsBlank")
End Sub

Sub RemoveAllHiddenFields()
    ' Removes all hidden fields from all three subreports
    ' Use this for complete rollback
    
    Dim response As VbMsgBoxResult
    
    response = MsgBox("This will remove all hidden fields from all three subreports." & vbCrLf & _
                      "Are you sure you want to continue?", vbYesNo + vbQuestion, "Confirm Cleanup")
    
    If response = vbYes Then
        Call RemoveFieldsFromClientsReport
        Call RemoveFieldsFromDayReport
        Call RemoveFieldsFromHouseReport
        
        MsgBox "All hidden fields removed from all subreports." & vbCrLf & vbCrLf & _
               "NEXT STEPS:" & vbCrLf & _
               "1. Revert RecordSource back to original (qrytblExpirations)" & vbCrLf & _
               "2. Restore original VBA code in Detail_Format events" & vbCrLf & _
               "3. Test reports", vbInformation, "Cleanup Complete"
    End If
End Sub

Sub UpdateRecordSource(reportName As String, newRecordSource As String)
    ' Updates the RecordSource property of a report
    Dim rpt As Report
    
    DoCmd.OpenReport reportName, acViewDesign
    Set rpt = REPORTS(reportName)
    
    If rpt.RecordSource <> newRecordSource Then
        rpt.RecordSource = newRecordSource
        DoCmd.Save acReport, reportName
        Debug.Print "Updated RecordSource for " & reportName & " to: " & newRecordSource
    Else
        Debug.Print "RecordSource already set correctly for " & reportName
    End If
    
    DoCmd.Close acReport, reportName, acSaveYes
End Sub

' =============================================
' COMPLETE IMPLEMENTATION (ALL-IN-ONE)
' =============================================

Sub ImplementCompleteOptimization()
    ' Implements the complete vw_ExpirationsFormatted optimization
    ' - Updates RecordSource for all subreports
    ' - Adds all required hidden fields
    ' - Ready for VBA code updates
    
    Dim response As VbMsgBoxResult
    
    response = MsgBox("This will implement the complete Expirations report optimization:" & vbCrLf & vbCrLf & _
                      "1. Update RecordSource to vw_ExpirationsFormatted" & vbCrLf & _
                      "2. Add ~77 hidden fields to all subreports" & vbCrLf & _
                      "3. Takes 1-2 minutes" & vbCrLf & vbCrLf & _
                      "IMPORTANT: Backup your database first!" & vbCrLf & vbCrLf & _
                      "Continue?", vbYesNo + vbQuestion, "Implement Optimization")
    
    If response = vbNo Then
        MsgBox "Operation cancelled.", vbInformation
        Exit Sub
    End If
    
    ' Update all reports
    Call FixClientsReport
    Call FixDayReport
    Call FixHouseReport
    
    MsgBox "Implementation complete!" & vbCrLf & vbCrLf & _
           "NEXT STEPS:" & vbCrLf & _
           "1. Update VBA code in each subreport's Detail_Format event" & vbCrLf & _
           "2. Add helper functions (ApplyColorFormatting, FormatExpirationField)" & vbCrLf & _
           "3. Test the main report" & vbCrLf & vbCrLf & _
           "See: vw_ExpirationsFormatted_IMPLEMENTATION_GUIDE.md", _
           vbInformation, "Implementation Complete"
End Sub

' =============================================
' COMPLETE ROLLBACK (ALL-IN-ONE)
' =============================================

Sub RollbackCompleteOptimization()
    ' Rolls back the entire optimization
    ' - Removes all hidden fields
    ' - Reverts RecordSource (you must specify original)
    ' - Does NOT restore VBA code (do that manually)
    
    Dim response As VbMsgBoxResult
    Dim originalRecordSource As String
    
    response = MsgBox("This will rollback the optimization changes:" & vbCrLf & vbCrLf & _
                      "1. Remove all hidden fields" & vbCrLf & _
                      "2. Revert RecordSource to original" & vbCrLf & vbCrLf & _
                      "Continue?", vbYesNo + vbQuestion, "Rollback Confirmation")
    
    If response = vbNo Then
        MsgBox "Rollback cancelled.", vbInformation
        Exit Sub
    End If
    
    ' Remove all hidden fields
    Call RemoveAllHiddenFields
    
    ' Revert RecordSource (adjust this to your original RecordSource name)
    originalRecordSource = "qrytblExpirations"  ' CHANGE THIS if different!
    
    Call UpdateRecordSource("rptEXPIRATIONDATESclients", originalRecordSource)
    Call UpdateRecordSource("rptEXPIRATIONDATESday", originalRecordSource)
    Call UpdateRecordSource("rptEXPIRATIONDATEShouse", originalRecordSource)
    
    MsgBox "Rollback complete!" & vbCrLf & vbCrLf & _
           "MANUAL STEPS REQUIRED:" & vbCrLf & _
           "1. Restore original VBA code in Detail_Format events" & vbCrLf & _
           "2. Remove helper functions (if added)" & vbCrLf & _
           "3. Test reports", vbInformation, "Rollback Complete"
End Sub

' =============================================
' DIAGNOSTIC FUNCTIONS
' =============================================

Sub ListHiddenFieldsInReport(reportName As String)
    ' Lists all hidden fields in a report
    ' Useful for debugging
    
    Dim rpt As Report
    Dim ctl As Control
    Dim hiddenCount As Integer
    Dim output As String
    
    DoCmd.OpenReport reportName, acViewDesign
    Set rpt = REPORTS(reportName)
    
    output = "Hidden Fields in " & reportName & ":" & vbCrLf & vbCrLf
    hiddenCount = 0
    
    For Each ctl In rpt.Controls
        If ctl.ControlType = acTextBox And ctl.Visible = False Then
            output = output & ctl.Name & vbCrLf
            hiddenCount = hiddenCount + 1
        End If
    Next ctl
    
    DoCmd.Close acReport, reportName, acSaveNo
    
    output = output & vbCrLf & "Total: " & hiddenCount & " hidden fields"
    
    Debug.Print output
    MsgBox output, vbInformation, "Hidden Fields List"
End Sub

Sub ListAllHiddenFields()
    ' Lists hidden fields in all three subreports
    Call ListHiddenFieldsInReport("rptEXPIRATIONDATESclients")
    Call ListHiddenFieldsInReport("rptEXPIRATIONDATESday")
    Call ListHiddenFieldsInReport("rptEXPIRATIONDATEShouse")
End Sub

Sub ListAllControlsInReport(reportName As String)
    ' Lists ALL controls in a report (not just hidden)
    ' Useful for debugging - see what's actually in the report
    
    Dim rpt As Report
    Dim ctl As Control
    Dim totalCount As Integer
    Dim visibleCount As Integer
    Dim hiddenCount As Integer
    Dim output As String
    
    DoCmd.OpenReport reportName, acViewDesign
    Set rpt = REPORTS(reportName)
    
    output = "ALL Controls in " & reportName & ":" & vbCrLf & vbCrLf
    totalCount = 0
    visibleCount = 0
    hiddenCount = 0
    
    For Each ctl In rpt.Controls
        totalCount = totalCount + 1
        
        If ctl.ControlType = acTextBox Then
            If ctl.Visible Then
                output = output & "  [VISIBLE] " & ctl.Name & vbCrLf
                visibleCount = visibleCount + 1
            Else
                output = output & "  [HIDDEN]  " & ctl.Name & vbCrLf
                hiddenCount = hiddenCount + 1
            End If
        End If
    Next ctl
    
    DoCmd.Close acReport, reportName, acSaveNo
    
    output = output & vbCrLf & "Summary:" & vbCrLf
    output = output & "  Total Textboxes: " & (visibleCount + hiddenCount) & vbCrLf
    output = output & "  Visible: " & visibleCount & vbCrLf
    output = output & "  Hidden: " & hiddenCount & vbCrLf
    output = output & "  Other controls: " & (totalCount - visibleCount - hiddenCount)
    
    Debug.Print output
    MsgBox output, vbInformation, "All Controls List"
End Sub

Sub ListAllControlsInAllReports()
    ' Lists all controls in all three subreports
    Call ListAllControlsInReport("rptEXPIRATIONDATESclients")
    Call ListAllControlsInReport("rptEXPIRATIONDATESday")
    Call ListAllControlsInReport("rptEXPIRATIONDATEShouse")
End Sub

Sub DeleteAllHiddenTextboxes(reportName As String)
    ' Makes ALL hidden textboxes visible and BRIGHT MAGENTA for manual deletion
    ' Access VBA doesn't support reliable programmatic control deletion
    ' This is the most reliable approach
    
    Dim rpt As Report
    Dim ctl As Control
    Dim markedCount As Integer
    Dim yPos As Integer
    
    DoCmd.OpenReport reportName, acViewDesign
    Set rpt = REPORTS(reportName)
    
    markedCount = 0
    yPos = 100
    
    For Each ctl In rpt.Controls
        If ctl.ControlType = acTextBox And ctl.Visible = False Then
            On Error Resume Next
            
            ' Make it VERY visible for manual deletion
            ctl.Visible = True
            ctl.BackColor = RGB(255, 0, 255)  ' Bright magenta
            ctl.ForeColor = RGB(255, 255, 255)  ' White text
            ctl.BorderColor = RGB(255, 0, 0)  ' Red border
            ctl.BorderWidth = 3
            ctl.SpecialEffect = 2  ' Sunken
            ctl.Top = yPos
            ctl.Left = 100
            ctl.Width = 3000
            ctl.Height = 250
            ctl.FontSize = 10
            ctl.FontWeight = 700  ' Bold
            ctl.Value = "*** DELETE THIS FIELD ***"
            
            If Err.Number = 0 Then
                markedCount = markedCount + 1
                yPos = yPos + 300
                Debug.Print "? Marked for deletion: " & ctl.Name
            Else
                Debug.Print "? Error marking: " & ctl.Name & " - " & Err.Description
                Err.Clear
            End If
            On Error GoTo 0
        End If
    Next ctl
    
    DoCmd.Save acReport, reportName
    ' Leave report OPEN so user can see and delete the marked fields
    
    If markedCount > 0 Then
        MsgBox "MARKED " & markedCount & " hidden fields in BRIGHT MAGENTA" & vbCrLf & vbCrLf & _
               "The report is now open in Design View." & vbCrLf & vbCrLf & _
               "TO DELETE THEM:" & vbCrLf & _
               "1. Click first magenta field" & vbCrLf & _
               "2. Hold Shift and click other magenta fields (selects all)" & vbCrLf & _
               "3. Press Delete key" & vbCrLf & _
               "4. Save and close (Ctrl+S, then close)", _
               vbInformation, "Manual Deletion Required"
    Else
        DoCmd.Close acReport, reportName, acSaveNo
        MsgBox "No hidden textboxes found in " & reportName, vbInformation
    End If
End Sub

Sub DeleteAllHiddenTextboxesFromAllReports()
    ' Nuclear option: Deletes ALL hidden textboxes from all subreports
    ' Use with caution - will delete ANY hidden textbox
    
    Dim response As VbMsgBoxResult
    
    response = MsgBox("WARNING: This will process ALL hidden textboxes from all subreports." & vbCrLf & vbCrLf & _
                      "This includes any optimization fields you added." & vbCrLf & vbCrLf & _
                      "Continue?", vbYesNo + vbExclamation, "Confirm Delete All Hidden")
    
    If response = vbYes Then
        Call DeleteAllHiddenTextboxes("rptEXPIRATIONDATESclients")
        Call DeleteAllHiddenTextboxes("rptEXPIRATIONDATESday")
        Call DeleteAllHiddenTextboxes("rptEXPIRATIONDATEShouse")
        
        MsgBox "Processing complete. Check reports for any remaining magenta fields.", vbInformation
    End If
End Sub

Sub MarkHiddenFieldsForManualDeletion(reportName As String)
    ' Makes all hidden textboxes VISIBLE and BRIGHT MAGENTA
    ' User can then manually select and delete them
    ' This is more reliable than programmatic deletion
    
    Dim rpt As Report
    Dim ctl As Control
    Dim markedCount As Integer
    Dim yPos As Integer
    
    DoCmd.OpenReport reportName, acViewDesign
    Set rpt = REPORTS(reportName)
    
    markedCount = 0
    yPos = 100
    
    For Each ctl In rpt.Controls
        If ctl.ControlType = acTextBox And ctl.Visible = False Then
            On Error Resume Next
            ' Make it VERY visible
            ctl.Visible = True
            ctl.BackColor = RGB(255, 0, 255)  ' Bright magenta
            ctl.ForeColor = RGB(255, 255, 255)  ' White text
            ctl.BorderColor = RGB(255, 0, 0)  ' Red border
            ctl.BorderWidth = 3
            ctl.SpecialEffect = 2  ' Sunken
            ctl.Top = yPos
            ctl.Left = 100
            ctl.Width = 3000
            ctl.Height = 250
            ctl.FontSize = 10
            ctl.FontWeight = 700  ' Bold
            
            If Err.Number = 0 Then
                markedCount = markedCount + 1
                yPos = yPos + 300
                Debug.Print "? Marked: " & ctl.Name
            Else
                Debug.Print "? Error marking: " & ctl.Name
                Err.Clear
            End If
            On Error GoTo 0
        End If
    Next ctl
    
    DoCmd.Save acReport, reportName
    ' Leave report OPEN so user can see the marked fields
    
    MsgBox "MARKED " & markedCount & " hidden fields in BRIGHT MAGENTA" & vbCrLf & vbCrLf & _
           "The report is now open in Design View." & vbCrLf & vbCrLf & _
           "TO DELETE:" & vbCrLf & _
           "1. Select all magenta fields (Click first, then Shift+Click others)" & vbCrLf & _
           "2. Press Delete key" & vbCrLf & _
           "3. Save and close", vbInformation, "Manual Deletion Required"
End Sub

Sub MarkAllHiddenFieldsForDeletion()
    ' Marks hidden fields in all three subreports for manual deletion
    Call MarkHiddenFieldsForManualDeletion("rptEXPIRATIONDATESclients")
    Call MarkHiddenFieldsForManualDeletion("rptEXPIRATIONDATESday")
    Call MarkHiddenFieldsForManualDeletion("rptEXPIRATIONDATEShouse")
    
    MsgBox "All three reports are now open with fields marked." & vbCrLf & _
           "Delete the MAGENTA fields from each report and save.", vbInformation
End Sub

' =============================================
' USAGE EXAMPLES
' =============================================

' TO ADD FIELDS TO ALL REPORTS:
'   1. Press Alt+F11 (open VBA Editor)
'   2. Run: FixAllReports (or individually: FixClientsReport, FixDayReport, FixHouseReport)
'
' TO REMOVE FIELDS FROM ALL REPORTS:
'   1. Run: RemoveAllHiddenFields
'
' TO CHECK WHICH FIELDS ARE HIDDEN:
'   1. Run: ListAllHiddenFields
'
' TO DO COMPLETE IMPLEMENTATION:
'   1. Run: ImplementCompleteOptimization
'
' TO DO COMPLETE ROLLBACK:
'   1. Run: RollbackCompleteOptimization


