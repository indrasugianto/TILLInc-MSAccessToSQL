' Module Name: Report_rptEXPIRATIONDATESday_Original
' Module Type: Document Module
' Lines of Code: 195
' Extracted: 2026-02-04 13:03:36

Option Compare Database
Option Explicit

Private Sub Detail_Format(Cancel As Integer, FormatCount As Integer)
'On Error GoTo ShowMeError
On Error GoTo 0
    Dim FN As String, LN As String, FindComma As Long

    LastVehicleChecklistCompletedFmt.Visible = False: LastVehicleChecklistCompletedTxt.Visible = False
    Select Case Format(LastVehicleChecklistCompleted, "YYYY-MM-DD")
        Case ExpMissing, ExpOptional, ExpNA, ExpPending: Call SetExpirationFieldProperties(LastVehicleChecklistCompletedTxt, LastVehicleChecklistCompleted)
        Case Else
            LastVehicleChecklistCompletedFmt.Visible = Not IsEmpty(LastVehicleChecklistCompleted) And Not IsNull(LastVehicleChecklistCompleted) ' And LengthN(10, LastVehicleChecklistCompleted)
            If LastVehicleChecklistCompletedFmt.Visible Then
                If (Int(Now) - DateValue(LastVehicleChecklistCompletedFmt)) >= Trig_Day_LVC_Red Then
                    LastVehicleChecklistCompletedFmt.ForeColor = RGB(255, 0, 0): LastVehicleChecklistCompletedFmt.FontWeight = 700
                    LastVehicleChecklistCompletedFmt.BorderColor = RGB(255, 0, 0): LastVehicleChecklistCompletedFmt.BorderStyle = 1
                Else
                    LastVehicleChecklistCompletedFmt.ForeColor = RGB(0, 0, 0): LastVehicleChecklistCompletedFmt.FontWeight = 400: LastVehicleChecklistCompletedFmt.BorderStyle = 0
                End If
            End If
    End Select

    DAYStaffTrainedInPrivacyBeforeFmt.Visible = False: DAYStaffTrainedInPrivacyBeforeTxt.Visible = False
    Select Case Format(DAYStaffTrainedInPrivacyBefore, "YYYY-MM-DD")
        Case ExpMissing, ExpOptional, ExpNA, ExpPending: Call SetExpirationFieldProperties(DAYStaffTrainedInPrivacyBeforeTxt, DAYStaffTrainedInPrivacyBefore)
        Case Else
            DAYStaffTrainedInPrivacyBeforeFmt.Visible = Not IsEmpty(DAYStaffTrainedInPrivacyBefore) And Not IsNull(DAYStaffTrainedInPrivacyBefore) '  And LengthN(10, DAYStaffTrainedInPrivacyBefore)
            If DAYStaffTrainedInPrivacyBeforeFmt.Visible Then
                If (DateValue(DAYStaffTrainedInPrivacyBeforeFmt) - Int(Now)) < Trig_Day_STP_Red Then
                    DAYStaffTrainedInPrivacyBeforeFmt.ForeColor = RGB(255, 0, 0): DAYStaffTrainedInPrivacyBeforeFmt.FontWeight = 700
                    DAYStaffTrainedInPrivacyBeforeFmt.BorderColor = RGB(255, 0, 0): DAYStaffTrainedInPrivacyBeforeFmt.BorderStyle = 1
                ElseIf (DateValue(DAYStaffTrainedInPrivacyBeforeFmt) - Int(Now)) <= Trig_Day_STP_Green Then
                    DAYStaffTrainedInPrivacyBeforeFmt.ForeColor = RGB(18, 94, 40): DAYStaffTrainedInPrivacyBeforeFmt.FontWeight = 700
                    DAYStaffTrainedInPrivacyBeforeFmt.BorderColor = RGB(18, 94, 40): DAYStaffTrainedInPrivacyBeforeFmt.BorderStyle = 1
                Else
                    DAYStaffTrainedInPrivacyBeforeFmt.ForeColor = RGB(0, 0, 0): DAYStaffTrainedInPrivacyBeforeFmt.FontWeight = 400: DAYStaffTrainedInPrivacyBeforeFmt.BorderStyle = 0
                End If
            End If
    End Select

    DAYAllPlansReviewedByStaffBeforeFmt.Visible = False: DAYAllPlansReviewedByStaffBeforeTxt.Visible = False
    Select Case Format(DAYAllPlansReviewedByStaffBefore, "YYYY-MM-DD")
        Case ExpMissing, ExpOptional, ExpNA, ExpPending: Call SetExpirationFieldProperties(DAYAllPlansReviewedByStaffBeforeTxt, DAYAllPlansReviewedByStaffBefore)
        Case Else
            DAYAllPlansReviewedByStaffBeforeFmt.Visible = Not IsEmpty(DAYAllPlansReviewedByStaffBefore) And Not IsNull(DAYAllPlansReviewedByStaffBefore) '  And LengthN(10, DAYAllPlansReviewedByStaffBefore)
            If DAYAllPlansReviewedByStaffBeforeFmt.Visible Then
                If (DateValue(DAYAllPlansReviewedByStaffBeforeFmt) - Int(Now)) < Trig_Day_APRS_Red Then
                    DAYAllPlansReviewedByStaffBeforeFmt.ForeColor = RGB(255, 0, 0): DAYAllPlansReviewedByStaffBeforeFmt.FontWeight = 700
                    DAYAllPlansReviewedByStaffBeforeFmt.BorderColor = RGB(255, 0, 0): DAYAllPlansReviewedByStaffBeforeFmt.BorderStyle = 1
                ElseIf (DateValue(DAYAllPlansReviewedByStaffBeforeFmt) - Int(Now)) <= Trig_Day_APRS_Green Then
                    DAYAllPlansReviewedByStaffBeforeFmt.ForeColor = RGB(18, 94, 40): DAYAllPlansReviewedByStaffBeforeFmt.FontWeight = 700
                    DAYAllPlansReviewedByStaffBeforeFmt.BorderColor = RGB(18, 94, 40): DAYAllPlansReviewedByStaffBeforeFmt.BorderStyle = 1
                Else
                    DAYAllPlansReviewedByStaffBeforeFmt.ForeColor = RGB(0, 0, 0): DAYAllPlansReviewedByStaffBeforeFmt.FontWeight = 400: DAYAllPlansReviewedByStaffBeforeFmt.BorderStyle = 0
                End If
            End If
    End Select

    DAYQtrlySafetyChecklistDueByFmt.Visible = False: DAYQtrlySafetyChecklistDueByTxt.Visible = False
    Select Case Format(DAYQtrlySafetyChecklistDueBy, "YYYY-MM-dd")
        Case ExpMissing, ExpOptional, ExpNA, ExpPending: Call SetExpirationFieldProperties(DAYQtrlySafetyChecklistDueByTxt, DAYQtrlySafetyChecklistDueBy)
        Case Else
            DAYQtrlySafetyChecklistDueByFmt.Visible = Not IsEmpty(DAYQtrlySafetyChecklistDueBy) And Not IsNull(DAYQtrlySafetyChecklistDueBy) '  And LengthN(10, DAYQtrlySafetyChecklistDueBy)
            If DAYQtrlySafetyChecklistDueByFmt.Visible Then
                If (DateValue(DAYQtrlySafetyChecklistDueByFmt) - Int(Now)) < Trig_Day_QSR_Red Then
                    DAYQtrlySafetyChecklistDueByFmt.ForeColor = RGB(255, 0, 0): DAYQtrlySafetyChecklistDueByFmt.FontWeight = 700
                    DAYQtrlySafetyChecklistDueByFmt.BorderColor = RGB(255, 0, 0): DAYQtrlySafetyChecklistDueByFmt.BorderStyle = 1
                ElseIf (DateValue(DAYQtrlySafetyChecklistDueByFmt) - Int(Now)) <= Trig_Day_QSR_Green Then
                    DAYQtrlySafetyChecklistDueByFmt.ForeColor = RGB(18, 94, 40): DAYQtrlySafetyChecklistDueByFmt.FontWeight = 700
                    DAYQtrlySafetyChecklistDueByFmt.BorderColor = RGB(18, 94, 40): DAYQtrlySafetyChecklistDueByFmt.BorderStyle = 1
                Else
                    DAYQtrlySafetyChecklistDueByFmt.ForeColor = RGB(0, 0, 0): DAYQtrlySafetyChecklistDueByFmt.FontWeight = 400: DAYQtrlySafetyChecklistDueByFmt.BorderStyle = 0
                End If
            End If
    End Select

    HumanRightsOfficerFmt.Visible = True
    If Not IsNull(HumanRightsOfficer) And Not IsEmpty(HumanRightsOfficer) Then
        FindComma = InStr(1, HumanRightsOfficer, ",", vbTextCompare)
        If FindComma > 0 Then
            HumanRightsOfficerFmt = CorrectProperNames(StrConv(HumanRightsOfficer, vbProperCase))
            LN = Left(HumanRightsOfficer, FindComma - 1): FN = Mid(HumanRightsOfficer, FindComma + 2, 255)
            FN = SpecialNames(FN)
            LN = SpecialNames(LN)
            HumanRightsOfficerFmt = FN & " " & LN
            If FindComma > 0 Then HumanRightsOfficerFmt = Left(HumanRightsOfficerFmt, FindComma) & StrConv(Mid(HumanRightsOfficerFmt, FindComma + 1, 1), vbUpperCase) & Mid(HumanRightsOfficerFmt, FindComma + 2, 256)
            HumanRightsOfficerFmt.FontWeight = 400
        Else
            HumanRightsOfficerFmt = Null
        End If
    End If
    Call CheckBlankField(HumanRightsOfficerFmt)
    
    HROTrainsStaffBeforeFmt.Visible = False: HROTrainsStaffBeforeTxt.Visible = False
    Select Case Format(HROTrainsStaffBefore, "YYYY-MM-dd")
        Case ExpMissing, ExpOptional, ExpNA, ExpPending: Call SetExpirationFieldProperties(HROTrainsStaffBeforeTxt, HROTrainsStaffBefore)
        Case Else
            HROTrainsStaffBeforeFmt.Visible = Not IsEmpty(HROTrainsStaffBefore) And Not IsNull(HROTrainsStaffBefore) '  And LengthN(10, HROTrainsStaffBefore)
            If HROTrainsStaffBeforeFmt.Visible Then
                If (DateValue(HROTrainsStaffBeforeFmt) - Int(Now)) < Trig_Day_HROTS_Red Then
                    HROTrainsStaffBeforeFmt.ForeColor = RGB(255, 0, 0): HROTrainsStaffBeforeFmt.FontWeight = 700
                    HROTrainsStaffBeforeFmt.BorderColor = RGB(255, 0, 0): HROTrainsStaffBeforeFmt.BorderStyle = 1
                ElseIf (DateValue(HROTrainsStaffBeforeFmt) - Int(Now)) <= Trig_Day_HROTS_Green Then
                    HROTrainsStaffBeforeFmt.ForeColor = RGB(18, 94, 40): HROTrainsStaffBeforeFmt.FontWeight = 700
                    HROTrainsStaffBeforeFmt.BorderColor = RGB(18, 94, 40): HROTrainsStaffBeforeFmt.BorderStyle = 1
                Else
                    HROTrainsStaffBeforeFmt.ForeColor = RGB(0, 0, 0): HROTrainsStaffBeforeFmt.FontWeight = 400: HROTrainsStaffBeforeFmt.BorderStyle = 0
                End If
            End If
    End Select

    HROTrainsIndividualsBeforeFmt.Visible = False: HROTrainsIndividualsBeforeTxt.Visible = False
    Select Case Format(HROTrainsIndividualsBefore, "YYYY-MM-dd")
        Case ExpMissing, ExpOptional, ExpNA, ExpPending: Call SetExpirationFieldProperties(HROTrainsIndividualsBeforeTxt, HROTrainsIndividualsBefore)
        Case Else
            HROTrainsIndividualsBeforeFmt.Visible = Not IsEmpty(HROTrainsIndividualsBefore) And Not IsNull(HROTrainsIndividualsBefore) '  And LengthN(10, HROTrainsIndividualsBefore)
            If HROTrainsIndividualsBeforeFmt.Visible Then
                If (DateValue(HROTrainsIndividualsBeforeFmt) - Int(Now)) < Trig_Day_HROTI_Red Then
                    HROTrainsIndividualsBeforeFmt.ForeColor = RGB(255, 0, 0): HROTrainsIndividualsBeforeFmt.FontWeight = 700
                    HROTrainsIndividualsBeforeFmt.BorderColor = RGB(255, 0, 0): HROTrainsIndividualsBeforeFmt.BorderStyle = 1
                ElseIf (DateValue(HROTrainsIndividualsBeforeFmt) - Int(Now)) <= Trig_Day_HROTI_Green Then
                    HROTrainsIndividualsBeforeFmt.ForeColor = RGB(18, 94, 40): HROTrainsIndividualsBeforeFmt.FontWeight = 700
                    HROTrainsIndividualsBeforeFmt.BorderColor = RGB(18, 94, 40): HROTrainsIndividualsBeforeFmt.BorderStyle = 1
                Else
                    HROTrainsIndividualsBeforeFmt.ForeColor = RGB(0, 0, 0): HROTrainsIndividualsBeforeFmt.FontWeight = 400: HROTrainsIndividualsBeforeFmt.BorderStyle = 0
                End If
            End If
    End Select
    
    FireSafetyOfficerFmt.Visible = True
    If Not IsNull(FireSafetyOfficer) And Not IsEmpty(FireSafetyOfficer) Then
        FindComma = InStr(1, FireSafetyOfficer, ",", vbTextCompare)
        If FindComma > 0 Then
            FireSafetyOfficerFmt = CorrectProperNames(StrConv(FireSafetyOfficer, vbProperCase))
            LN = Left(FireSafetyOfficer, FindComma - 1): FN = Mid(FireSafetyOfficer, FindComma + 2, 255)
            FN = SpecialNames(FN)
            LN = SpecialNames(LN)
            FireSafetyOfficerFmt = CorrectProperNames(StrConv(FN & " " & LN, vbProperCase)): FindComma = InStr(1, FireSafetyOfficerFmt, "-", vbTextCompare)
            If FindComma > 0 Then FireSafetyOfficerFmt = Left(FireSafetyOfficerFmt, FindComma) & StrConv(Mid(FireSafetyOfficerFmt, FindComma + 1, 1), vbUpperCase) & Mid(FireSafetyOfficerFmt, FindComma + 2, 256)
            FireSafetyOfficerFmt.FontWeight = 400
        Else
            FireSafetyOfficerFmt = Null
        End If
    End If
    Call CheckBlankField(FireSafetyOfficerFmt)
    
    FSOTrainsStaffBeforeFmt.Visible = False: FSOTrainsStaffBeforeTxt.Visible = False
    Select Case Format(FSOTrainsStaffBefore, "YYYY-MM-dd")
        Case ExpMissing, ExpOptional, ExpNA, ExpPending: Call SetExpirationFieldProperties(FSOTrainsStaffBeforeTxt, FSOTrainsStaffBefore)
        Case Else
            FSOTrainsStaffBeforeFmt.Visible = Not IsEmpty(FSOTrainsStaffBefore) And Not IsNull(FSOTrainsStaffBefore) '  And LengthN(10, FSOTrainsStaffBefore)
            If FSOTrainsStaffBeforeFmt.Visible Then
                If (DateValue(FSOTrainsStaffBeforeFmt) - Int(Now)) < Trig_Day_FSOTS_Red Then
                    FSOTrainsStaffBeforeFmt.ForeColor = RGB(255, 0, 0): FSOTrainsStaffBeforeFmt.FontWeight = 700
                    FSOTrainsStaffBeforeFmt.BorderColor = RGB(255, 0, 0): FSOTrainsStaffBeforeFmt.BorderStyle = 1
                ElseIf (DateValue(FSOTrainsStaffBeforeFmt) - Int(Now)) <= Trig_Day_FSOTS_Green Then
                    FSOTrainsStaffBeforeFmt.ForeColor = RGB(18, 94, 40): FSOTrainsStaffBeforeFmt.FontWeight = 700
                    FSOTrainsStaffBeforeFmt.BorderColor = RGB(18, 94, 40): FSOTrainsStaffBeforeFmt.BorderStyle = 1
                Else
                    FSOTrainsStaffBeforeFmt.ForeColor = RGB(0, 0, 0): FSOTrainsStaffBeforeFmt.FontWeight = 400: FSOTrainsStaffBeforeFmt.BorderStyle = 0
                End If
            End If
    End Select

    FSOTrainsIndividualsBeforeFmt.Visible = False: FSOTrainsIndividualsBeforeTxt.Visible = False:
    Select Case Format(FSOTrainsIndividualsBefore, "YYYY-MM-dd")
        Case ExpMissing, ExpOptional, ExpNA, ExpPending: Call SetExpirationFieldProperties(FSOTrainsIndividualsBeforeTxt, FSOTrainsIndividualsBefore)
        Case Else
            FSOTrainsIndividualsBeforeFmt.Visible = Not IsEmpty(FSOTrainsIndividualsBefore) And Not IsNull(FSOTrainsIndividualsBefore) '  And LengthN(10, FSOTrainsIndividualsBefore)
            If FSOTrainsIndividualsBeforeFmt.Visible Then
                If (DateValue(FSOTrainsIndividualsBeforeFmt) - Int(Now)) < Trig_Day_FSOTI_Red Then
                    FSOTrainsIndividualsBeforeFmt.ForeColor = RGB(255, 0, 0): FSOTrainsIndividualsBeforeFmt.FontWeight = 700
                    FSOTrainsIndividualsBeforeFmt.BorderColor = RGB(255, 0, 0): FSOTrainsIndividualsBeforeFmt.BorderStyle = 1
                ElseIf (DateValue(FSOTrainsIndividualsBeforeFmt) - Int(Now)) <= Trig_Day_FSOTI_Green Then
                    FSOTrainsIndividualsBeforeFmt.ForeColor = RGB(18, 94, 40): FSOTrainsIndividualsBeforeFmt.FontWeight = 700
                    FSOTrainsIndividualsBeforeFmt.BorderColor = RGB(18, 94, 40): FSOTrainsIndividualsBeforeFmt.BorderStyle = 1
                Else
                    FSOTrainsIndividualsBeforeFmt.ForeColor = RGB(0, 0, 0): FSOTrainsIndividualsBeforeFmt.FontWeight = 400: FSOTrainsIndividualsBeforeFmt.BorderStyle = 0
                End If
            End If
    End Select
    Exit Sub
ShowMeError:
    MsgBox "Error # " & Str(Err.Number) & " was generated by " & Me.Name & Chr(13) & Err.Description, vbOKOnly, "Error", Err.HelpFile, Err.HelpContext
End Sub

Private Sub CheckBlankField(DataField As Variant)
'On Error GoTo ShowMeError
On Error GoTo 0
    If IsEmpty(DataField) Or IsNull(DataField) Or Len(DataField) = 0 Then DataField.BackColor = RGB(255, 0, 0) Else DataField.BackColor = RGB(255, 255, 255)
    Exit Sub
ShowMeError:
    MsgBox "Error # " & Str(Err.Number) & " was generated by " & Me.Name & Chr(13) & Err.Description, vbOKOnly, "Error", Err.HelpFile, Err.HelpContext
End Sub
