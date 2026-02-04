' Module Name: Report_rptEXPIRATIONDATESstaff_Original
' Module Type: Document Module
' Lines of Code: 248
' Extracted: 2026-02-04 13:03:36

Option Compare Database
Option Explicit

Public DateThreeMonthsOut As Date, ThisYear As Date, NextYear As Date

Private Sub Detail_Format(Cancel As Integer, FormatCount As Integer)
'On Error GoTo ShowMeError
On Error GoTo 0
    Dim FindDash As Long

    FullName = RTrim(StrConv(FirstName, 3)) & " " & RTrim(StrConv(LastName, 3)): FindDash = InStr(1, FullName, "-", vbTextCompare)
    If FindDash > 0 Then FullName = Left(FullName, FindDash) & StrConv(Mid(FullName, FindDash + 1, 1), vbUpperCase) & Mid(FullName, FindDash + 2, 256)

    BBPTxt.Visible = False: BBPFmt.Visible = False
    Select Case Format(BBP, "YYYY-MM-DD")
        Case ExpMissing, ExpOptional, ExpNA, ExpPending: Call SetExpirationFieldProperties(BBPTxt, BBP)
        Case Else
            BBPFmt.Visible = Not IsEmpty(BBP) And Not IsNull(BBP) ' And LengthN(10, BBP)
            If BBPFmt.Visible Then
                If (BBPFmt - Int(Now)) < Trig_Staff_BBP_Red Then
                    BBPFmt.ForeColor = RGB(255, 0, 0): BBPFmt.FontWeight = 700: BBPFmt.BorderColor = RGB(255, 0, 0): BBPFmt.BorderStyle = 1
                ElseIf (BBPFmt - Int(Now)) <= Trig_Staff_BBP_Green Then
                    BBPFmt.ForeColor = RGB(18, 94, 40): BBPFmt.FontWeight = 700: BBPFmt.BorderColor = RGB(18, 94, 40): BBPFmt.BorderStyle = 1
                Else
                    BBPFmt.ForeColor = RGB(0, 0, 0): BBPFmt.FontWeight = 400: BBPFmt.BorderStyle = 0
                End If
            End If
    End Select

    BIPTxt.Visible = False: BIPFmt.Visible = False
    Select Case Format(BackInjuryPrevention, "YYYY-MM-DD")
        Case ExpMissing, ExpOptional, ExpNA, ExpPending: Call SetExpirationFieldProperties(BIPTxt, BackInjuryPrevention)
        Case Else
            BIPFmt.Visible = Not IsEmpty(BackInjuryPrevention) And Not IsNull(BackInjuryPrevention) ' And LengthN(10, BackInjuryPrevention)
            If BIPFmt.Visible Then
                If (BIPFmt - Int(Now)) < Trig_Staff_BIP_Red Then
                    BIPFmt.ForeColor = RGB(255, 0, 0): BIPFmt.FontWeight = 700: BIPFmt.BorderColor = RGB(255, 0, 0): BIPFmt.BorderStyle = 1
                ElseIf (BIPFmt - Int(Now)) <= Trig_Staff_BIP_Green Then
                    BIPFmt.ForeColor = RGB(18, 94, 40): BIPFmt.FontWeight = 700: BIPFmt.BorderColor = RGB(18, 94, 40): BIPFmt.BorderStyle = 1
                Else
                    BIPFmt.ForeColor = RGB(0, 0, 0): BIPFmt.FontWeight = 400: BIPFmt.BorderStyle = 0
                End If
            End If
    End Select
    
    CPRTxt.Visible = False: CPRFmt.Visible = False
    Select Case Format(CPR, "YYYY-MM-DD")
        Case ExpMissing, ExpOptional, ExpNA, ExpPending: Call SetExpirationFieldProperties(CPRTxt, CPR)
        Case Else
            CPRFmt.Visible = Not IsEmpty(CPR) And Not IsNull(CPR) ' And LengthN(10, CPR)
            If CPRFmt.Visible Then
                If (CPRFmt - Int(Now)) < Trig_Staff_CPR_Red Then
                    CPRFmt.ForeColor = RGB(255, 0, 0): CPRFmt.FontWeight = 700: CPRFmt.BorderColor = RGB(255, 0, 0): CPRFmt.BorderStyle = 1
                ElseIf (CPRFmt - Int(Now)) <= Trig_Staff_CPR_Green Then
                    CPRFmt.ForeColor = RGB(18, 94, 40): CPRFmt.FontWeight = 700: CPRFmt.BorderColor = RGB(18, 94, 40): CPRFmt.BorderStyle = 1
                Else
                    CPRFmt.ForeColor = RGB(0, 0, 0): CPRFmt.FontWeight = 400: CPRFmt.BorderStyle = 0
                End If
            End If
    End Select
    
    DefensiveDrivingTxt.Visible = False: DefensiveDrivingFmt.Visible = False
    Select Case Format(DefensiveDriving, "YYYY-MM-DD")
        Case ExpMissing, ExpOptional, ExpNA, ExpPending: Call SetExpirationFieldProperties(DefensiveDrivingTxt, DefensiveDriving)
        Case Else:                           If DefensiveDrivingFmt = "Done" Then DefensiveDrivingFmt.Visible = True Else DefensiveDrivingFmt.Visible = False
    End Select
    
    DriversLicenseTxt.Visible = False: DriversLicenseFmt.Visible = False
    Select Case Format(DriversLicense, "YYYY-MM-DD")
        Case ExpMissing, ExpOptional, ExpNA, ExpPending: Call SetExpirationFieldProperties(DriversLicenseTxt, DriversLicense)
        Case Else
            DriversLicenseFmt.Visible = Not IsEmpty(DriversLicense) And Not IsNull(DriversLicense) ' And LengthN(10, DriversLicense)
            If DriversLicenseFmt.Visible Then
                If (DriversLicenseFmt - Int(Now)) < Trig_Staff_DL_Red Then
                    DriversLicenseFmt.ForeColor = RGB(255, 0, 0): DriversLicenseFmt.FontWeight = 700: DriversLicenseFmt.BorderColor = RGB(255, 0, 0): DriversLicenseFmt.BorderStyle = 1
                ElseIf (DriversLicenseFmt - Int(Now)) <= Trig_Staff_DL_Green Then
                    DriversLicenseFmt.ForeColor = RGB(18, 94, 40): DriversLicenseFmt.FontWeight = 700: DriversLicenseFmt.BorderColor = RGB(18, 94, 40): DriversLicenseFmt.BorderStyle = 1
                Else
                    DriversLicenseFmt.ForeColor = RGB(0, 0, 0): DriversLicenseFmt.FontWeight = 400: DriversLicenseFmt.BorderStyle = 0
                End If
            End If
    End Select
    
    FirstAidTxt.Visible = False: FirstAidFmt.Visible = False
    Select Case Format(FirstAid, "YYYY-MM-DD")
        Case ExpMissing, ExpOptional, ExpNA, ExpPending: Call SetExpirationFieldProperties(FirstAidTxt, FirstAid)
        Case Else
            FirstAidFmt.Visible = Not IsEmpty(FirstAid) And Not IsNull(FirstAid) ' And LengthN(10, FirstAid)
            If FirstAidFmt.Visible Then
                If (FirstAidFmt - Int(Now)) < Trig_Staff_FA_Red Then
                    FirstAidFmt.ForeColor = RGB(255, 0, 0): FirstAidFmt.FontWeight = 700: FirstAidFmt.BorderColor = RGB(255, 0, 0): FirstAidFmt.BorderStyle = 1
                ElseIf (FirstAidFmt - Int(Now)) <= Trig_Staff_FA_Green Then
                    FirstAidFmt.ForeColor = RGB(18, 94, 40): FirstAidFmt.FontWeight = 700: FirstAidFmt.BorderColor = RGB(18, 94, 40): FirstAidFmt.BorderStyle = 1:
                Else
                    FirstAidFmt.ForeColor = RGB(0, 0, 0): FirstAidFmt.FontWeight = 400: FirstAidFmt.BorderStyle = 0
                End If
            End If
    End Select
    
    PBSTxt.Visible = False: PBSFmt.Visible = False
    Select Case Format(PBS, "YYYY-MM-DD")
        Case ExpMissing, ExpOptional, ExpNA, ExpPending: Call SetExpirationFieldProperties(PBSTxt, PBS)
        Case Else:                           If PBSFmt = "Done" Then PBSFmt.Visible = True Else PBSFmt.Visible = False
    End Select
    
    SafetyCaresTxt.Visible = False: SafetyCaresFmt.Visible = False
    Select Case Format(SafetyCares, "YYYY-MM-DD")
        Case ExpMissing, ExpOptional, ExpNA, ExpPending: Call SetExpirationFieldProperties(SafetyCaresTxt, SafetyCares)
        Case Else
            SafetyCaresFmt.Visible = Not IsEmpty(SafetyCares) And Not IsNull(SafetyCares) ' And LengthN(10, SafetyCares)
            If SafetyCaresFmt.Visible Then
                If (SafetyCaresFmt - Int(Now)) < Trig_Staff_SC_Red Then
                    SafetyCaresFmt.ForeColor = RGB(255, 0, 0): SafetyCaresFmt.FontWeight = 700: SafetyCaresFmt.BorderColor = RGB(255, 0, 0): SafetyCaresFmt.BorderStyle = 1
                ElseIf (SafetyCaresFmt - Int(Now)) <= Trig_Staff_SC_Green Then
                    SafetyCaresFmt.ForeColor = RGB(18, 94, 40): SafetyCaresFmt.FontWeight = 700: SafetyCaresFmt.BorderColor = RGB(18, 94, 40): SafetyCaresFmt.BorderStyle = 1
                Else
                    SafetyCaresFmt.ForeColor = RGB(0, 0, 0): SafetyCaresFmt.FontWeight = 400: SafetyCaresFmt.BorderStyle = 0
                End If
            End If
    End Select
    
    TBTxt.Visible = False: TBFmt.Visible = False
    Select Case Format(TB, "YYYY-MM-DD")
        Case ExpMissing, ExpOptional, ExpCompleted, ExpNA: Call SetExpirationFieldProperties(TBTxt, TB)
        Case ExpPending
            Call SetExpirationFieldProperties(TBTxt, TB)
            If Location = "Hollis" Or DLookup("ABI", "tblLocations", "GPName='" & Location & "'") = True Or DLookup("Department", "tblLocations", "GPName='" & Location & "'") = "Day Services" Then
                TBTxt.Visible = True: TBTxt.Caption = "Pending": TBTxt.ForeColor = RGB(255, 0, 0): TBTxt.FontWeight = 700: TBTxt.BorderStyle = 1: TBTxt.BorderColor = RGB(255, 0, 0)
            End If
        Case Else
            TBFmt.Visible = Not IsEmpty(TB) And Not IsNull(TB) ' And LengthN(10, TB)
            If TBFmt.Visible Then
                If (TBFmt - Int(Now)) < Trig_Staff_TB_Red Then
                    TBFmt.ForeColor = RGB(255, 0, 0): TBFmt.FontWeight = 700: TBFmt.BorderColor = RGB(255, 0, 0): TBFmt.BorderStyle = 1
                ElseIf (TBFmt - Int(Now)) <= Trig_Staff_TB_Green Then
                    TBFmt.ForeColor = RGB(18, 94, 40): TBFmt.FontWeight = 700: TBFmt.BorderColor = RGB(18, 94, 40): TBFmt.BorderStyle = 1
                Else
                    TBFmt.ForeColor = RGB(0, 0, 0): TBFmt.FontWeight = 400: TBFmt.BorderStyle = 0
                End If
            End If
    End Select

    WheelchairSafetyTxt.Visible = False: WheelchairSafetyFmt.Visible = False
    Select Case Format(WheelchairSafety, "YYYY-MM-DD")
        Case ExpMissing, ExpOptional, ExpNA, ExpPending: Call SetExpirationFieldProperties(WheelchairSafetyTxt, WheelchairSafety)
        Case Else:                                       If WheelchairSafetyFmt = "Done" Then WheelchairSafetyFmt.Visible = True Else WheelchairSafetyFmt.Visible = False
    End Select
    
    WorkplaceViolenceTxt.Visible = False: WorkplaceViolenceFmt.Visible = False
    Select Case Format(WorkplaceViolence, "YYYY-MM-DD")
        Case ExpMissing, ExpOptional, ExpNA, ExpPending: Call SetExpirationFieldProperties(WorkplaceViolenceTxt, WorkplaceViolence)
        Case Else
            WorkplaceViolenceFmt.Visible = Not IsEmpty(WorkplaceViolence) And Not IsNull(WorkplaceViolence) ' And LengthN(10, WorkplaceViolence)
            If WorkplaceViolenceFmt.Visible Then
                If (WorkplaceViolenceFmt - Int(Now)) < Trig_Staff_WV_Red Then
                    WorkplaceViolenceFmt.ForeColor = RGB(255, 0, 0): WorkplaceViolenceFmt.FontWeight = 700: WorkplaceViolenceFmt.BorderColor = RGB(255, 0, 0): WorkplaceViolenceFmt.BorderStyle = 1
                ElseIf (WorkplaceViolenceFmt - Int(Now)) <= Trig_Staff_WV_Green Then
                    WorkplaceViolenceFmt.ForeColor = RGB(18, 94, 40): WorkplaceViolenceFmt.FontWeight = 700: WorkplaceViolenceFmt.BorderColor = RGB(18, 94, 40): WorkplaceViolenceFmt.BorderStyle = 1
                Else
                    WorkplaceViolenceFmt.ForeColor = RGB(0, 0, 0): WorkplaceViolenceFmt.FontWeight = 400: WorkplaceViolenceFmt.BorderStyle = 0
                End If
            End If
    End Select

    ProfessionalLicensesTxt.Visible = False: ProfessionalLicensesFmt.Visible = False
    Select Case Format(ProfessionalLicenses, "YYYY-MM-DD")
        Case ExpMissing, ExpOptional, ExpNA, ExpPending: Call SetExpirationFieldProperties(ProfessionalLicensesTxt, ProfessionalLicenses)
        Case Else
            ProfessionalLicensesFmt.Visible = Not IsEmpty(ProfessionalLicenses) And Not IsNull(ProfessionalLicenses) ' And LengthN(10, ProfessionalLicenses)
            If ProfessionalLicensesFmt.Visible Then
                If (ProfessionalLicensesFmt - Int(Now)) < Trig_Staff_PL_Red Then
                    ProfessionalLicensesFmt.ForeColor = RGB(255, 0, 0): ProfessionalLicensesFmt.FontWeight = 700: ProfessionalLicensesFmt.BorderColor = RGB(255, 0, 0): ProfessionalLicensesFmt.BorderStyle = 1
                ElseIf (ProfessionalLicensesFmt - Int(Now)) <= Trig_Staff_PL_Green Then
                    ProfessionalLicensesFmt.ForeColor = RGB(18, 94, 40): ProfessionalLicensesFmt.FontWeight = 700: ProfessionalLicensesFmt.BorderColor = RGB(18, 94, 40): ProfessionalLicensesFmt.BorderStyle = 1
                Else
                    ProfessionalLicensesFmt.ForeColor = RGB(0, 0, 0): ProfessionalLicensesFmt.FontWeight = 400: ProfessionalLicensesFmt.BorderStyle = 0
                End If
            End If
    End Select
   
    MapCertLabel.Visible = True: MapCertLine.Visible = True: MAPCertTxt.Visible = False: MAPCertFmt.Visible = False
    Select Case Format(MAPCert, "YYYY-MM-DD")
        Case ExpMissing, ExpOptional, ExpNA, ExpPending: Call SetExpirationFieldProperties(MAPCertTxt, MAPCert)
        Case Else
            MAPCertFmt.Visible = Not IsEmpty(MAPCert) And Not IsNull(MAPCert) ' And LengthN(10, MAPCert)
            If MAPCertFmt.Visible Then
                MAPCertFmt = DateValue(MAPCert)
                If (MAPCertFmt - Int(Now)) < Trig_Staff_MAP_Red Then
                    MAPCertFmt.ForeColor = RGB(255, 0, 0): MAPCertFmt.FontWeight = 700: MAPCertFmt.BorderColor = RGB(255, 0, 0): MAPCertFmt.BorderStyle = 1
                ElseIf (MAPCertFmt - Int(Now)) <= Trig_Staff_MAP_Green Then
                    MAPCertFmt.ForeColor = RGB(18, 94, 40): MAPCertFmt.FontWeight = 700: MAPCertFmt.BorderColor = RGB(18, 94, 40): MAPCertFmt.BorderStyle = 1
                Else
                    MAPCertFmt.ForeColor = RGB(0, 0, 0): MAPCertFmt.FontWeight = 400: MAPCertFmt.BorderStyle = 0
                End If
            Else
                If Left(Location, 4) = "DED-" And _
                   Not (Left(Location, 7) = "DED-C10" Or _
                        Left(Location, 6) = "DED-C1" Or _
                        Left(Location, 6) = "DED-C2" Or _
                        Left(Location, 6) = "DED-C3" Or _
                        Left(Location, 6) = "DED-C4" Or _
                        Left(Location, 6) = "DED-C5" Or _
                        Left(Location, 6) = "DED-C6" Or _
                        Left(Location, 6) = "DED-C7" Or _
                        Left(Location, 6) = "DED-C8" Or _
                        Left(Location, 6) = "DED-C9") _
                Then ' It's N/A.
                    MAPCertFmt.Visible = False
                    MAPCertTxt.Visible = True
                    MAPCertTxt.Caption = "N/A"
                    MAPCertTxt.ForeColor = RGB(0, 0, 0): MAPCertTxt.FontWeight = 400: MAPCertTxt.BorderStyle = 0
                Else
                    MAPCertFmt = DateAdd("d", 120, DateValue(AdjustedStartDate)): MAPCertFmt.Visible = True: MAPCertFmt.ForeColor = RGB(111, 49, 152): MAPCertFmt.FontWeight = 700: MAPCertFmt.BorderColor = RGB(111, 49, 152): MAPCertFmt.BorderStyle = 1
                End If
            End If
    End Select
        
    EvalDueByFmt.Visible = Not IsEmpty(EvalDueBy) And Not IsNull(EvalDueBy) ' And LengthN(8, EvalDueBy)
    If EvalDueByFmt.Visible Then
        If (EvalDueByFmt - Int(Now)) < Trig_Staff_EVL_Red Then
            EvalDueByFmt.ForeColor = RGB(255, 0, 0): EvalDueByFmt.FontWeight = 700: EvalDueByFmt.BorderColor = RGB(255, 0, 0): EvalDueByFmt.BorderStyle = 1
        ElseIf (EvalDueByFmt - Int(Now)) <= Trig_Staff_EVL_Green Then
            EvalDueByFmt.ForeColor = RGB(18, 94, 40): EvalDueByFmt.FontWeight = 700: EvalDueByFmt.BorderColor = RGB(18, 94, 40): EvalDueByFmt.BorderStyle = 1
        Else
            EvalDueByFmt.ForeColor = RGB(0, 0, 0): EvalDueByFmt.FontWeight = 400: EvalDueByFmt.BorderStyle = 0
        End If
        If ThreeMonthEvaluation Then
            If (EvalDueByFmt - Int(Now)) < Trig_Staff_3MO_Red Then
                EvalDueByFmt.ForeColor = RGB(255, 0, 0): EvalDueByFmt.FontWeight = 700: EvalDueByFmt.BorderColor = RGB(0, 0, 255): EvalDueByFmt.BorderStyle = 1
            Else
                EvalDueByFmt.ForeColor = RGB(0, 0, 255): EvalDueByFmt.FontWeight = 700: EvalDueByFmt.BorderColor = RGB(0, 0, 255): EvalDueByFmt.BorderStyle = 1
            End If
        End If
    End If
   
'   LastSupervisionFmt.Visible = Not IsEmpty(LastSupervision) And Not IsNull(LastSupervision) ' And LengthN(8, LastSupervision): LastSupervisionFmt.FontWeight = 400
    LastSupervisionFmt.Visible = False
    If LastSupervisionFmt.Visible Then
        If (LastSupervisionFmt - Int(Now)) < Trig_Staff_SUP_Red Then
            LastSupervisionFmt.ForeColor = RGB(255, 0, 0): LastSupervisionFmt.FontWeight = 700: LastSupervisionFmt.BorderColor = RGB(255, 0, 0): LastSupervisionFmt.BorderStyle = 1
        Else
            LastSupervisionFmt.ForeColor = RGB(0, 0, 0): LastSupervisionFmt.FontWeight = 400: LastSupervisionFmt.BorderStyle = 0
        End If
    End If
    Exit Sub
ShowMeError:
    MsgBox "Error # " & Str(Err.Number) & " was generated by " & Me.Name & Chr(13) & Err.Description, vbOKOnly, "Error", Err.HelpFile, Err.HelpContext
End Sub
