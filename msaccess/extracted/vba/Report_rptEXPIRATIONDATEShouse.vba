' Module Name: Report_rptEXPIRATIONDATEShouse
' Module Type: Document Module
' Lines of Code: 213
' Extracted: 1/29/2026 4:12:25 PM

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
                If (Int(Now) - DateValue(LastVehicleChecklistCompletedFmt)) >= Trig_Res_LVC_Red Then
                    LastVehicleChecklistCompletedFmt.ForeColor = RGB(255, 0, 0): LastVehicleChecklistCompletedFmt.FontWeight = 700
                    LastVehicleChecklistCompletedFmt.BorderColor = RGB(255, 0, 0): LastVehicleChecklistCompletedFmt.BorderStyle = 1
                Else
                    LastVehicleChecklistCompletedFmt.ForeColor = RGB(0, 0, 0): LastVehicleChecklistCompletedFmt.FontWeight = 400: LastVehicleChecklistCompletedFmt.BorderStyle = 0
                End If
            End If
    End Select
    
    MostRecentAsleepFireDrillFmt.Visible = False: MostRecentAsleepFireDrillTxt.Visible = False
    Select Case Format(MostRecentAsleepFireDrill, "YYYY-MM-DD")
        Case ExpMissing, ExpOptional, ExpNA, ExpPending: Call SetExpirationFieldProperties(MostRecentAsleepFireDrillTxt, MostRecentAsleepFireDrill)
        Case Else
            MostRecentAsleepFireDrillFmt.Visible = Not IsEmpty(MostRecentAsleepFireDrill) And Not IsNull(MostRecentAsleepFireDrill) ' And LengthN(10, MostRecentAsleepFireDrill)
            If MostRecentAsleepFireDrillFmt.Visible Then
                If (DateAdd("m", 14, DateValue(MostRecentAsleepFireDrillFmt)) - Int(Now)) < Trig_Res_MRFD_Red Then
                    MostRecentAsleepFireDrillFmt.ForeColor = RGB(255, 0, 0): MostRecentAsleepFireDrillFmt.FontWeight = 700 ' Red
                    MostRecentAsleepFireDrillFmt.BorderColor = RGB(255, 0, 0): MostRecentAsleepFireDrillFmt.BorderStyle = 1
                ElseIf (DateAdd("m", 14, DateValue(MostRecentAsleepFireDrillFmt)) - Int(Now)) <= Trig_Res_MRFD_Green Then
                    MostRecentAsleepFireDrillFmt.ForeColor = RGB(18, 94, 40): MostRecentAsleepFireDrillFmt.FontWeight = 700
                    MostRecentAsleepFireDrillFmt.BorderColor = RGB(18, 94, 40): MostRecentAsleepFireDrillFmt.BorderStyle = 1 ' Green
                Else
                    MostRecentAsleepFireDrillFmt.ForeColor = RGB(0, 0, 0): MostRecentAsleepFireDrillFmt.FontWeight = 400: MostRecentAsleepFireDrillFmt.BorderStyle = 0
                End If
            End If
    End Select
    
    NextRecentAsleepFireDrillFmt.Visible = Not IsEmpty(NextRecentAsleepFireDrill) And Not IsNull(NextRecentAsleepFireDrill) ' And LengthN(10, NextRecentAsleepFireDrill): NextRecentAsleepFireDrillFmt.FontWeight = 400: NextRecentAsleepFireDrillFmt.BorderStyle = 0
    
    HousePlansReviewedByStaffBeforeFmt.Visible = False: HousePlansReviewedByStaffBeforeTxt.Visible = False
    Select Case Format(HousePlansReviewedByStaffBefore, "YYYY-MM-DD")
        Case ExpMissing, ExpOptional, ExpNA, ExpPending: Call SetExpirationFieldProperties(HousePlansReviewedByStaffBeforeTxt, HousePlansReviewedByStaffBefore)
        Case Else
            HousePlansReviewedByStaffBeforeFmt.Visible = Not IsEmpty(HousePlansReviewedByStaffBefore) And Not IsNull(HousePlansReviewedByStaffBefore) ' And LengthN(10, HousePlansReviewedByStaffBefore)
            If HousePlansReviewedByStaffBeforeFmt.Visible Then
                If (DateValue(HousePlansReviewedByStaffBeforeFmt) - Int(Now)) < Trig_Res_HPR_Red Then
                    HousePlansReviewedByStaffBeforeFmt.ForeColor = RGB(255, 0, 0): HousePlansReviewedByStaffBeforeFmt.FontWeight = 700
                    HousePlansReviewedByStaffBeforeFmt.BorderColor = RGB(255, 0, 0): HousePlansReviewedByStaffBeforeFmt.BorderStyle = 1
                ElseIf (DateValue(HousePlansReviewedByStaffBeforeFmt) - Int(Now)) <= Trig_Res_HPR_Green Then
                    HousePlansReviewedByStaffBeforeFmt.ForeColor = RGB(18, 94, 40): HousePlansReviewedByStaffBeforeFmt.FontWeight = 700
                    HousePlansReviewedByStaffBeforeFmt.BorderColor = RGB(18, 94, 40): HousePlansReviewedByStaffBeforeFmt.BorderStyle = 1
                Else
                    HousePlansReviewedByStaffBeforeFmt.ForeColor = RGB(0, 0, 0): HousePlansReviewedByStaffBeforeFmt.FontWeight = 400: HousePlansReviewedByStaffBeforeFmt.BorderStyle = 0
                End If
            End If
    End Select

    HouseSafetyPlanExpiresFmt.Visible = False: HouseSafetyPlanExpiresTxt.Visible = False
    Select Case Format(HouseSafetyPlanExpires, "YYYY-MM-DD")
        Case ExpMissing, ExpOptional, ExpNA, ExpPending: Call SetExpirationFieldProperties(HouseSafetyPlanExpiresTxt, HouseSafetyPlanExpires)
        Case Else
            HouseSafetyPlanExpiresFmt.Visible = Not IsEmpty(HouseSafetyPlanExpires) And Not IsNull(HouseSafetyPlanExpires) ' And LengthN(10, HouseSafetyPlanExpires)
            If HouseSafetyPlanExpiresFmt.Visible Then
                If (DateValue(HouseSafetyPlanExpiresFmt) - Int(Now)) < Trig_Res_HSPE_Red Then
                    HouseSafetyPlanExpiresFmt.ForeColor = RGB(255, 0, 0): HouseSafetyPlanExpiresFmt.FontWeight = 700
                    HouseSafetyPlanExpiresFmt.BorderColor = RGB(255, 0, 0): HouseSafetyPlanExpiresFmt.BorderStyle = 1
                ElseIf (DateValue(HouseSafetyPlanExpiresFmt) - Int(Now)) <= Trig_Res_HSPE_Green Then
                    HouseSafetyPlanExpiresFmt.ForeColor = RGB(18, 94, 40): HouseSafetyPlanExpiresFmt.FontWeight = 700
                    HouseSafetyPlanExpiresFmt.BorderColor = RGB(18, 94, 40): HouseSafetyPlanExpiresFmt.BorderStyle = 1
                Else
                    HouseSafetyPlanExpiresFmt.ForeColor = RGB(0, 0, 0): HouseSafetyPlanExpiresFmt.FontWeight = 400: HouseSafetyPlanExpiresFmt.BorderStyle = 0
                End If
            End If
    End Select
    
    MAPChecklistCompletedFmt.Visible = False: MAPChecklistCompletedTxt.Visible = False
    Select Case Format(MAPChecklistCompleted, "YYYY-MM-DD")
        Case ExpMissing, ExpOptional, ExpNA, ExpPending: Call SetExpirationFieldProperties(MAPChecklistCompletedTxt, MAPChecklistCompleted)
        Case Else
            MAPChecklistCompletedFmt.Visible = Not IsEmpty(MAPChecklistCompleted) And Not IsNull(MAPChecklistCompleted) ' And LengthN(10, MAPChecklistCompleted)
            If MAPChecklistCompletedFmt.Visible Then
                If (Int(Now) - DateValue(MAPChecklistCompletedFmt)) >= Trig_Res_MAP_Red Then
                    MAPChecklistCompletedFmt.ForeColor = RGB(255, 0, 0): MAPChecklistCompletedFmt.FontWeight = 700
                    MAPChecklistCompletedFmt.BorderColor = RGB(255, 0, 0): MAPChecklistCompletedFmt.BorderStyle = 1
                Else
                    MAPChecklistCompletedFmt.ForeColor = RGB(0, 0, 0): MAPChecklistCompletedFmt.FontWeight = 400: MAPChecklistCompletedFmt.BorderStyle = 0
                End If
            End If
    End Select

    HumanRightsOfficerFmt.Visible = True
    If Not IsNull(HumanRightsOfficer) And Not IsEmpty(HumanRightsOfficer) Then
        FindComma = InStr(1, HumanRightsOfficer, ",", vbTextCompare)
        If FindComma > 0 Then
            HumanRightsOfficerFmt = CorrectProperNames(StrConv(HumanRightsOfficerFmt, vbProperCase))
            LN = Left(HumanRightsOfficer, FindComma - 1): FN = Mid(HumanRightsOfficer, FindComma + 2, 255)
            FN = SpecialNames(FN)
            LN = SpecialNames(LN)
            HumanRightsOfficerFmt = FN & " " & LN
            FindComma = InStr(1, HumanRightsOfficerFmt, "-", vbTextCompare)
            If FindComma > 0 Then HumanRightsOfficerFmt = Left(HumanRightsOfficerFmt, FindComma) & StrConv(Mid(HumanRightsOfficerFmt, FindComma + 1, 1), vbUpperCase) & Mid(HumanRightsOfficerFmt, FindComma + 2, 256)
            HumanRightsOfficerFmt.FontWeight = 400
        Else
            HumanRightsOfficerFmt = Null
        End If
    End If
    Call CheckBlankField(HumanRightsOfficerFmt)
    
    HROTrainsStaffBeforeFmt.Visible = False: HROTrainsStaffBeforeTxt.Visible = False
    Select Case Format(HROTrainsStaffBefore, "YYYY-MM-DD")
        Case ExpMissing, ExpOptional, ExpNA, ExpPending: Call SetExpirationFieldProperties(HROTrainsStaffBeforeTxt, HROTrainsStaffBefore)
        Case Else
            HROTrainsStaffBeforeFmt.Visible = Not IsEmpty(HROTrainsStaffBefore) And Not IsNull(HROTrainsStaffBefore) ' And LengthN(10, HROTrainsStaffBefore)
            If HROTrainsStaffBeforeFmt.Visible Then
                If (DateValue(HROTrainsStaffBeforeFmt) - Int(Now)) < Trig_Res_HROTS_Red Then
                    HROTrainsStaffBeforeFmt.ForeColor = RGB(255, 0, 0): HROTrainsStaffBeforeFmt.FontWeight = 700
                    HROTrainsStaffBeforeFmt.BorderColor = RGB(255, 0, 0): HROTrainsStaffBeforeFmt.BorderStyle = 1
                ElseIf (DateValue(HROTrainsStaffBeforeFmt) - Int(Now)) <= Trig_Res_HROTS_Green Then
                    HROTrainsStaffBeforeFmt.ForeColor = RGB(18, 94, 40): HROTrainsStaffBeforeFmt.FontWeight = 700
                    HROTrainsStaffBeforeFmt.BorderColor = RGB(18, 94, 40): HROTrainsStaffBeforeFmt.BorderStyle = 1
                Else
                    HROTrainsStaffBeforeFmt.ForeColor = RGB(0, 0, 0): HROTrainsStaffBeforeFmt.FontWeight = 400: HROTrainsStaffBeforeFmt.BorderStyle = 0
                End If
            End If
    End Select

    HROTrainsIndividualsBeforeFmt.Visible = False: HROTrainsIndividualsBeforeTxt.Visible = False
    Select Case Format(HROTrainsIndividualsBefore, "YYYY-MM-DD")
        Case ExpMissing, ExpOptional, ExpNA, ExpPending: Call SetExpirationFieldProperties(HROTrainsIndividualsBeforeTxt, HROTrainsIndividualsBefore)
        Case Else
            HROTrainsIndividualsBeforeFmt.Visible = Not IsEmpty(HROTrainsIndividualsBefore) And Not IsNull(HROTrainsIndividualsBefore) ' And LengthN(10, HROTrainsIndividualsBefore)
            If HROTrainsIndividualsBeforeFmt.Visible Then
                If (DateValue(HROTrainsIndividualsBeforeFmt) - Int(Now)) < Trig_Res_HROTI_Red Then
                    HROTrainsIndividualsBeforeFmt.ForeColor = RGB(255, 0, 0): HROTrainsIndividualsBeforeFmt.FontWeight = 700
                    HROTrainsIndividualsBeforeFmt.BorderColor = RGB(255, 0, 0): HROTrainsIndividualsBeforeFmt.BorderStyle = 1
                ElseIf (DateValue(HROTrainsIndividualsBeforeFmt) - Int(Now)) <= Trig_Res_HROTI_Green Then
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
            FireSafetyOfficerFmt = FN & " " & LN
            If FindComma > 0 Then FireSafetyOfficerFmt = Left(FireSafetyOfficerFmt, FindComma) & StrConv(Mid(FireSafetyOfficerFmt, FindComma + 1, 1), vbUpperCase) & Mid(FireSafetyOfficerFmt, FindComma + 2, 256)
            FireSafetyOfficerFmt.FontWeight = 400
        Else
            FireSafetyOfficerFmt = Null
        End If
    End If
    Call CheckBlankField(FireSafetyOfficerFmt)
    
    FSOTrainsStaffBeforeFmt.Visible = False: FSOTrainsStaffBeforeTxt.Visible = False
    Select Case Format(FSOTrainsStaffBefore, "YYYY-MM-DD")
        Case ExpMissing, ExpOptional, ExpNA, ExpPending: Call SetExpirationFieldProperties(FSOTrainsStaffBeforeTxt, FSOTrainsStaffBefore)
        Case Else
            FSOTrainsStaffBeforeFmt.Visible = Not IsEmpty(FSOTrainsStaffBefore) And Not IsNull(FSOTrainsStaffBefore) ' And LengthN(10, FSOTrainsStaffBefore)
            If FSOTrainsStaffBeforeFmt.Visible Then
                If (DateValue(FSOTrainsStaffBeforeFmt) - Int(Now)) < Trig_Res_FSOTS_Red Then
                    FSOTrainsStaffBeforeFmt.ForeColor = RGB(255, 0, 0): FSOTrainsStaffBeforeFmt.FontWeight = 700
                    FSOTrainsStaffBeforeFmt.BorderColor = RGB(255, 0, 0): FSOTrainsStaffBeforeFmt.BorderStyle = 1
                ElseIf (DateValue(FSOTrainsStaffBeforeFmt) - Int(Now)) <= Trig_Res_FSOTS_Green Then
                    FSOTrainsStaffBeforeFmt.ForeColor = RGB(18, 94, 40): FSOTrainsStaffBeforeFmt.FontWeight = 700
                    FSOTrainsStaffBeforeFmt.BorderColor = RGB(18, 94, 40): FSOTrainsStaffBeforeFmt.BorderStyle = 1
                Else
                    FSOTrainsStaffBeforeFmt.ForeColor = RGB(0, 0, 0): FSOTrainsStaffBeforeFmt.FontWeight = 400: FSOTrainsStaffBeforeFmt.BorderStyle = 0
                End If
            End If
    End Select

    FSOTrainsIndividualsBeforeFmt.Visible = False: FSOTrainsIndividualsBeforeTxt.Visible = False:
    Select Case Format(FSOTrainsIndividualsBefore, "YYYY-MM-DD")
        Case ExpMissing, ExpOptional, ExpNA, ExpPending: Call SetExpirationFieldProperties(FSOTrainsIndividualsBeforeTxt, FSOTrainsIndividualsBefore)
        Case Else
            FSOTrainsIndividualsBeforeFmt.Visible = Not IsEmpty(FSOTrainsIndividualsBefore) And Not IsNull(FSOTrainsIndividualsBefore) '  And LengthN(10, FSOTrainsIndividualsBefore)
            If FSOTrainsIndividualsBeforeFmt.Visible Then
                If (DateValue(FSOTrainsIndividualsBeforeFmt) - Int(Now)) < Trig_Res_FSOTI_Red Then
                    FSOTrainsIndividualsBeforeFmt.ForeColor = RGB(255, 0, 0): FSOTrainsIndividualsBeforeFmt.FontWeight = 700
                    FSOTrainsIndividualsBeforeFmt.BorderColor = RGB(255, 0, 0): FSOTrainsIndividualsBeforeFmt.BorderStyle = 1
                ElseIf (DateValue(FSOTrainsIndividualsBeforeFmt) - Int(Now)) <= Trig_Res_FSOTI_Green Then
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