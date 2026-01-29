' Module Name: Form_frmLocations
' Module Type: Document Module
' Lines of Code: 428
' Extracted: 1/29/2026 4:12:23 PM

Option Compare Database
Option Explicit

Dim LastRecordCityTown As Variant, LastRecordLocationName As Variant, FirstRecordCityTown As Variant, FirstRecordLocationName As Variant, NavigationError As Boolean

Private Sub BigCloseForm_Click()
    Me.Visible = False
    Form_frmMainMenu.SetFocus
End Sub

Private Sub ClearAllFilters_Click()
    SelectCityTown = ""
    SelectLocation = ""
    FilterCriteria = "<Show all>"
    Me.FilterOn = False
    Me.Refresh
End Sub

Private Sub Form_Load()
    Dim rs As DAO.Recordset
    
    NavigationError = False
    NewRecordInProgress = False
    ' Who Can delete records?
    Button_DeleteRecord.Visible = DCount("Action", "catUserPermissions", "User='" & Form_frmMainMenu.UserName & "' AND Action='Can delete records'") > 0
    ' Count records.
    Set rs = Me.RecordsetClone
    rs.MoveFirst
    FirstRecordCityTown = rs!CityTown: FirstRecordLocationName = rs!LocationName
    rs.MoveLast
    LastRecordCityTown = rs!CityTown:  LastRecordLocationName = rs!LocationName
    Set rs = Nothing
End Sub

Private Sub Form_Current()
On Error GoTo ShowMeError
    Dim NameParts() As String
    
    If NavigationError Then Call Button_FirstRecord_Click
    CityTown.Enabled = False: CityTown.Locked = True: LocationName.Enabled = False: LocationName.Locked = True: ABI.Visible = False
    Button_FirstRecord.Visible = True: Button_PreviousRecord.Visible = True: Button_LastRecord.Visible = True: Button_NextRecord.Visible = True
    ResettlementCityTown.Visible = False
    ResettlementLocation.Visible = False
    If IsNull(State) Or Len(State) <= 0 Then County.Visible = False Else County.Visible = True
    If CityTown = FirstRecordCityTown And LocationName = FirstRecordLocationName Then
        Button_FirstRecord.Visible = False: Button_PreviousRecord.Visible = False
    ElseIf CityTown = LastRecordCityTown And LocationName = LastRecordLocationName Then
        Button_LastRecord.Visible = False:  Button_NextRecord.Visible = False
    End If
    ResCapacity.Visible = False: ResTILLOwned.Visible = False: DDSArea.Visible = False: DDSRegion.Visible = False
    ' Who can see the manager logonID?
    If Me.NewRecord Then
        ManagerLogonID.Visible = True
    Else
        ManagerLogonID.Visible = DCount("Action", "catUserPermissions", _
            "User='" & Form_frmMainMenu.UserName & "' AND Action='Can see and edit Locations LogonID'") > 0 And _
            (Department = "Residential Services" Or Department = "Day Services")
    End If
    ' Who can see expirations?
    Call WhoCanSeeExpirations
    ' Validate Expirations.
    Call ValidateExpirations
    ' Update residential-specific fields.
    If Department = "Residential Services" Or (Department = "Individualized Support Options" And CityTown <> "Dedham") Then
        ResettlementCityTown.Visible = True: ResettlementLocation.Visible = True: ResTILLOwned.Visible = True
    End If
    If Department = "Residential Services" Then
        ABI.Visible = True: DDSArea.Visible = True: DDSRegion.Visible = True: Call Cluster_AfterUpdate: EditPrimaryContact.Visible = False
        EditSecondaryContact.Visible = False: EditTertiaryContact.Visible = False: Cluster.Visible = True: CommmunityRNName.Visible = True
        LPN1Name.Visible = True: LPN2Name.Visible = True
        
        NameParts = Split(DLookup("IndexedName", "tblPeople", "OfficeCityTown = '" & [CityTown] & "' AND OfficeLocationName = '" & [LocationName] & "' AND (StaffTitle = 'Residence Manager' OR StaffTitle = 'Site Coordinator')"), "/")
        StaffPrimaryContactLastName = NameParts(0): StaffPrimaryContactFirstName = NameParts(1): StaffPrimaryContactMiddleInitial = NameParts(2)
        StaffPrimaryContactIndexedName = DLookup("IndexedName", "tblPeople", "OfficeCityTown = """ & [CityTown] & """ AND OfficeLocationName = """ & [LocationName] & """")
        
        StaffSecondaryContactLastName = DLookup("ClusterManagerLastName", "catClusters", "ClusterID = " & [Cluster])
        StaffSecondaryContactFirstName = DLookup("ClusterManagerFirstName", "catClusters", "ClusterID = " & [Cluster])
        StaffSecondaryContactMiddleInitial = DLookup("ClusterManagerMiddleInitial", "catClusters", "ClusterID = " & [Cluster])
        StaffSecondaryContactIndexedName = DLookup("IndexedName", "tblPeople", "FirstName='" & StaffSecondaryContactFirstName & "' AND LastName='" & StaffSecondaryContactLastName & "'")
        
        StaffTertiaryContactLastName = DLookup("ClusterDirectorLastName", "catClusters", "ClusterID = " & [Cluster])
        StaffTertiaryContactFirstName = DLookup("ClusterDirectorFirstName", "catClusters", "ClusterID = " & [Cluster])
        StaffTertiaryContactMiddleInitial = DLookup("ClusterDirectorMiddleInitial", "catClusters", "ClusterID = " & [Cluster])
        StaffTertiaryContactIndexedName = DLookup("IndexedName", "tblPeople", "FirstName='" & StaffTertiaryContactFirstName & "' AND LastName='" & StaffTertiaryContactLastName & "'")
        
        Leader.Visible = True: ResCapacity.Visible = True: DID.Visible = False
        Me.Refresh
    Else
        Leader.Visible = False: Cluster.Visible = False: EditPrimaryContact.Visible = True: EditSecondaryContact.Visible = True: EditTertiaryContact.Visible = True
        DID.Visible = True: CommmunityRNName.Visible = False: LPN1Name.Visible = False: LPN2Name.Visible = False
    End If
    
    If Department = "Individualized Support Options" Then
        DDSArea.Visible = True: DDSRegion.Visible = True
    End If
    
    Call CheckOpenContact(PrimaryContact): Call CheckOpenContact(SecondaryContact): Call CheckOpenContact(TertiaryContact)
    SelectCityTown.SetFocus
    Me.Caption = "Place: " & CityTown & " - " & LocationName
    If Not AddressValidated Then If Len(Address) > 0 And Len(City) > 0 And Len(State) > 0 Then Call ValidateAddress(Address, City, State, ZIP, AddressValidated, Me)
    Exit Sub
ShowMeError:
    If Err.Number = 94 Then Call Button_FirstRecord_Click
End Sub

Private Sub WhoCanSeeExpirations()
    ExpirationsHeader.Visible = DCount("Action", "catUserPermissions", "User='" & Form_frmMainMenu.UserName & "' AND Action='Can see Location Expirations'") > 0
    LastVehicleChecklistCompleted.Visible = ExpirationsHeader.Visible
    MostRecentAsleepFireDrill.Visible = ExpirationsHeader.Visible
    NextRecentAsleepFireDrill.Visible = ExpirationsHeader.Visible
    MAPChecklistCompleted.Visible = ExpirationsHeader.Visible
    HumanRightsOfficer.Visible = ExpirationsHeader.Visible
    HROTrainsStaffBefore.Visible = ExpirationsHeader.Visible
    HROTrainsIndividualsBefore.Visible = ExpirationsHeader.Visible
    FireSafetyOfficer.Visible = ExpirationsHeader.Visible
    FSOTrainsStaffBefore.Visible = ExpirationsHeader.Visible
    FSOTrainsIndividualsBefore.Visible = ExpirationsHeader.Visible
    
    If Department = "Day Services" Then
        HouseSafetyPlanExpires.Visible = False: HousePlansReviewedByStaffBefore.Visible = False
        DAYStaffTrainedInPrivacyBefore.Visible = ExpirationsHeader.Visible: DAYAllPlansReviewedByStaffBefore.Visible = ExpirationsHeader.Visible
        DAYQtrlySafetyChecklistDueBy.Visible = ExpirationsHeader.Visible
    Else
        HouseSafetyPlanExpires.Visible = ExpirationsHeader.Visible: HousePlansReviewedByStaffBefore.Visible = ExpirationsHeader.Visible
        DAYStaffTrainedInPrivacyBefore.Visible = False: DAYAllPlansReviewedByStaffBefore.Visible = False
        DAYQtrlySafetyChecklistDueBy.Visible = False
    End If
End Sub

Private Sub FindLocationRecord()
    Dim rst As DAO.Recordset
    
    Set rst = Me.RecordsetClone
    rst.FindFirst "CityTown = '" & SelectCityTown & "' AND LocationName = '" & SelectLocation & "'"
    Me.Bookmark = rst.Bookmark
    SelectCityTown = "": SelectLocation = ""
End Sub

Private Function EditContact(WhichContact As String) As Boolean
    EditContact = True: ContactControl = WhichContact: DoCmd.OpenForm "frmLocationsSelectStaff"
End Function

Private Sub CheckOpenContact(Field As Variant)
    If Field = "*** OPEN ***" Then Field.BackColor = RGB(255, 153, 0) Else Field.BackColor = RGB(255, 255, 0)
End Sub

Private Function ValidateExpirations() As Boolean
    If IsNull(LastVehicleChecklistCompleted) Or Len(LastVehicleChecklistCompleted) <= 0 _
        Then LastVehicleChecklistCompleted.BackColor = RGB(255, 0, 0) _
        Else LastVehicleChecklistCompleted.BackColor = RGB(255, 255, 255)
    If IsNull(MostRecentAsleepFireDrill) Or Len(MostRecentAsleepFireDrill) <= 0 _
        Then MostRecentAsleepFireDrill.BackColor = RGB(255, 0, 0) _
        Else MostRecentAsleepFireDrill.BackColor = RGB(255, 255, 255)
    If IsNull(NextRecentAsleepFireDrill) Or Len(NextRecentAsleepFireDrill) <= 0 _
        Then NextRecentAsleepFireDrill.BackColor = RGB(255, 0, 0) _
        Else NextRecentAsleepFireDrill.BackColor = RGB(255, 255, 255)
    If IsNull(HousePlansReviewedByStaffBefore) Or Len(HousePlansReviewedByStaffBefore) <= 0 _
        Then HousePlansReviewedByStaffBefore.BackColor = RGB(255, 0, 0) _
        Else HousePlansReviewedByStaffBefore.BackColor = RGB(255, 255, 255)
    If IsNull(HouseSafetyPlanExpires) Or Len(HouseSafetyPlanExpires) <= 0 _
        Then HouseSafetyPlanExpires.BackColor = RGB(255, 0, 0) _
        Else HouseSafetyPlanExpires.BackColor = RGB(255, 255, 255)
    If IsNull(DAYStaffTrainedInPrivacyBefore) Or Len(DAYStaffTrainedInPrivacyBefore) <= 0 _
        Then DAYStaffTrainedInPrivacyBefore.BackColor = RGB(255, 0, 0) _
        Else DAYStaffTrainedInPrivacyBefore.BackColor = RGB(255, 255, 255)
    If IsNull(DAYAllPlansReviewedByStaffBefore) Or Len(DAYAllPlansReviewedByStaffBefore) <= 0 _
        Then DAYAllPlansReviewedByStaffBefore.BackColor = RGB(255, 0, 0) _
        Else DAYAllPlansReviewedByStaffBefore.BackColor = RGB(255, 255, 255)
    If IsNull(DAYQtrlySafetyChecklistDueBy) Or Len(DAYQtrlySafetyChecklistDueBy) <= 0 _
        Then DAYQtrlySafetyChecklistDueBy.BackColor = RGB(255, 0, 0) _
        Else DAYQtrlySafetyChecklistDueBy.BackColor = RGB(255, 255, 255)
    If IsNull(MAPChecklistCompleted) Or Len(MAPChecklistCompleted) <= 0 _
        Then MAPChecklistCompleted.BackColor = RGB(255, 0, 0) _
        Else MAPChecklistCompleted.BackColor = RGB(255, 255, 255)
    If IsNull(HumanRightsOfficer) Or Len(HumanRightsOfficer) <= 0 _
        Then HumanRightsOfficer.BackColor = RGB(255, 0, 0) _
        Else HumanRightsOfficer.BackColor = RGB(255, 255, 255)
    If IsNull(HROTrainsStaffBefore) Or Len(HROTrainsStaffBefore) <= 0 _
        Then HROTrainsStaffBefore.BackColor = RGB(255, 0, 0) _
        Else HROTrainsStaffBefore.BackColor = RGB(255, 255, 255)
    If IsNull(HROTrainsIndividualsBefore) Or Len(HROTrainsIndividualsBefore) <= 0 _
        Then HROTrainsIndividualsBefore.BackColor = RGB(255, 0, 0) _
        Else HROTrainsIndividualsBefore.BackColor = RGB(255, 255, 255)
    If IsNull(FireSafetyOfficer) Or Len(FireSafetyOfficer) <= 0 _
        Then FireSafetyOfficer.BackColor = RGB(255, 0, 0) _
        Else FireSafetyOfficer.BackColor = RGB(255, 255, 255)
    If IsNull(FSOTrainsStaffBefore) Or Len(FSOTrainsStaffBefore) <= 0 _
        Then FSOTrainsStaffBefore.BackColor = RGB(255, 0, 0) _
        Else FSOTrainsStaffBefore.BackColor = RGB(255, 255, 255)
    If IsNull(FSOTrainsIndividualsBefore) Or Len(FSOTrainsIndividualsBefore) <= 0 _
        Then FSOTrainsIndividualsBefore.BackColor = RGB(255, 0, 0) _
        Else FSOTrainsIndividualsBefore.BackColor = RGB(255, 255, 255)
    Me.Refresh
End Function

Private Function MoveToRecord(Pointer As Integer) As Boolean
On Error GoTo ShowMeError
    MoveToRecord = True: NavigationError = False: SelectCityTown = "": SelectLocation = ""
    DoCmd.GoToRecord , , Pointer
    Exit Function
ShowMeError:
'   MsgBox "Can't go any further.", vbOKOnly, "ERROR!"
    MsgBox "Error # " & Str(Err.Number) & " was generated by " & Me.Name & Chr(13) & Err.Description, vbOKOnly, "Error", Err.HelpFile, Err.HelpContext
    NavigationError = True
    Call Button_FirstRecord_Click
    MoveToRecord = False
End Function

Private Function ValidateNewAddress() As Boolean
    ValidateNewAddress = True
    If NewRecordInProgress Then
        If AddressValidated Then Exit Function
        If Len(Address) > 0 And Len(City) > 0 And Len(State) > 0 Then Call ValidateAddress(Address, City, State, ZIP, AddressValidated, Me)
    End If
End Function

' Focus actions: Usually a validation check.

Private Function CheckAddress() As Boolean
    CheckAddress = True
    If AddressValidated Then Exit Function
    If NewRecordInProgress Then
        If Len(Address) > 0 And Len(City) > 0 And Len(State) > 0 Then
            Call ValidateAddress(Address, City, State, ZIP, AddressValidated, Me)
            DDSArea = DLookup("Area", "catDDSAreasAndRegions", "CityTown='" & CityTown & "'")
            DDSRegion = DLookup("Region", "catDDSAreasAndRegions", "CityTown='" & CityTown & "'")
            If ABI Then
                Select Case DDSRegion
                    Case "DDS Metro Region":        DDSRegion = "DDS Metro ABI": DDSArea = "Metro ABI"
                    Case "DDS Northeast Region":    DDSRegion = "DDS Northeast ABI": DDSArea = "Northeast ABI"
                End Select
            End If
        End If
    End If
End Function

Private Sub Department_LostFocus()
    If Department = "Residential Services" Then
        Cluster.Visible = True: Cluster.ValidationRule = "Is Not Null": Leader.Visible = True
        Call Cluster_AfterUpdate: CommmunityRNName.Visible = True: LPN1Name.Visible = True: LPN2Name.Visible = True
    Else
        Cluster.Visible = False: Cluster.Value = Null: Cluster.ValidationRule = "": Leader.Visible = False: CommmunityRNName.Visible = False
        LPN1Name.Visible = False: LPN2Name.Visible = False
    End If
End Sub

Private Sub Cluster_LostFocus()
    StaffSecondaryContactLastName = DLookup("ClusterManagerLastName", "catClusters", "ClusterID = " & Cluster)
    StaffSecondaryContactFirstName = DLookup("ClusterManagerFirstName", "catClusters", "ClusterID = " & Cluster)
    StaffSecondaryContactMiddleInitial = DLookup("ClusterManagerMiddleInitial", "catClusters", "ClusterID = " & Cluster)
    StaffTertiaryContactLastName = DLookup("ClusterDirectorLastName", "catClusters", "ClusterID = " & Cluster)
    StaffTertiaryContactFirstName = DLookup("ClusterDirectorFirstName", "catClusters", "ClusterID = " & Cluster)
    StaffTertiaryContactMiddleInitial = DLookup("ClusterDirectorMiddleInitial", "catClusters", "ClusterID = " & Cluster)
End Sub

' AfterUpdate Actions: Usually entails validating expirations and logging changes into the Change log.

Private Sub Address_AfterUpdate()
    If IsNull(Address) Or Len(Address) < 0 Then Address.BackColor = RGB(255, 0, 0) Else Address.BackColor = RGB(255, 255, 255): AddressValidated = False
End Sub

Private Sub City_AfterUpdate()
    If IsNull(City) Or Len(City) < 0 Then City.BackColor = RGB(255, 0, 0) Else City.BackColor = RGB(255, 255, 255): AddressValidated = False
End Sub

Private Sub Cluster_AfterUpdate()
    If IsNull(Cluster) Or Len(Cluster) < 0 Then Cluster.BackColor = RGB(255, 0, 0) Else Cluster.BackColor = RGB(255, 255, 255)
End Sub

Private Function ParseNurseName(NurseName As Variant, NurseIndexedName As Variant) As Boolean
    Dim FN As Variant, LN As Variant, Ptr As Long
    
    ParseNurseName = True
    If IsNull(NurseName) Then
        NurseIndexedName = Null
    Else
        Ptr = InStr(1, NurseName, " "): FN = Left(NurseName, Ptr - 1): LN = Mid(NurseName, Ptr + 1, 255)
        NurseIndexedName = DLookup("IndexedName", "tblPeople", "LastName='" & LN & "' AND FirstName='" & FN & "'")
    End If
End Function

Private Sub CostCenter_AfterUpdate()
    If IsNull(CostCenter) Or Len(CostCenter) < 0 Then CostCenter.BackColor = RGB(255, 0, 0) Else CostCenter.BackColor = RGB(255, 255, 255)
End Sub

Private Sub EmailAddress_AfterUpdate()
    If IsNull(EmailAddress) Or Len(EmailAddress) < 0 Then EmailAddress.BackColor = RGB(255, 0, 0) Else EmailAddress.BackColor = RGB(255, 255, 255)
End Sub

Private Function ExpirationFieldsAfterUpdate(WhichField As String) As Boolean
    ExpirationFieldsAfterUpdate = True
    Select Case WhichField
        Case "FSOTrainsIndividualsBefore":          If ValidateDateString(FSOTrainsIndividualsBefore) Then Call ValidateExpirations
        Case "FSOTrainsStaffBefore":                If ValidateDateString(FSOTrainsStaffBefore) Then Call ValidateExpirations
        Case "HousePlansReviewedByStaffBefore":     If ValidateDateString(HousePlansReviewedByStaffBefore) Then Call ValidateExpirations
        Case "HouseSafetyPlanExpires":              If ValidateDateString(HouseSafetyPlanExpires) Then Call ValidateExpirations
        Case "HROTrainsIndividualsBefore":          If ValidateDateString(HROTrainsIndividualsBefore) Then Call ValidateExpirations
        Case "HROTrainsStaffBefore":                If ValidateDateString(HROTrainsStaffBefore) Then Call ValidateExpirations
        Case "LastVehicleChecklistCompleted":       If ValidateDateString(LastVehicleChecklistCompleted) Then Call ValidateExpirations
        Case "MAPChecklistCompleted":               If ValidateDateString(MAPChecklistCompleted) Then Call ValidateExpirations
        Case "MostRecentAsleepFireDrill":           If ValidateDateString(MostRecentAsleepFireDrill) Then Call ValidateExpirations
        Case "DAYStaffTrainedInPrivacyBefore":      If ValidateDateString(DAYStaffTrainedInPrivacyBefore) Then Call ValidateExpirations
        Case "DAYAllPlansReviewedByStaffBefore":    If ValidateDateString(DAYAllPlansReviewedByStaffBefore) Then Call ValidateExpirations
        Case "DAYQtrlySafetyChecklistDueBy":        If ValidateDateString(DAYQtrlySafetyChecklistDueBy) Then Call ValidateExpirations
        Case "NextRecentAsleepFireDrill":           If ValidateDateString(NextRecentAsleepFireDrill) Then Call ValidateExpirations
    End Select
End Function

Private Sub PhoneNumber_AfterUpdate()
    If IsNull(PhoneNumber) Or Len(PhoneNumber) < 0 Then PhoneNumber.BackColor = RGB(255, 0, 0) Else PhoneNumber.BackColor = RGB(255, 255, 255)
End Sub

Private Sub ResettlementCityTown_AfterUpdate()
    CityPicker = ResettlementCityTown
End Sub

Private Sub State_AfterUpdate()
    If IsNull(State) Or Len(State) < 0 Then
        State.BackColor = RGB(255, 0, 0): County.Visible = False
    Else
        State.BackColor = RGB(255, 255, 255): County.Visible = True
    End If
    AddressValidated = False
End Sub

Private Sub ZIP_AfterUpdate()
    If IsNull(ZIP) Or Len(ZIP) < 0 Then ZIP.BackColor = RGB(255, 0, 0) Else ZIP.BackColor = RGB(255, 255, 255)
    AddressValidated = False
End Sub

Private Sub SelectCityTown_AfterUpdate()
On Error GoTo ShowMeError
    Call DropTempTables
    TILLDataBase.Execute "SELECT LocationName INTO temptbl0 FROM tblLocations WHERE CityTown='" & [SelectCityTown] & "';", dbSeeChanges: Call BriefDelay
    Call BriefDelay
    
    SelectLocation = "": SelectLocation.Requery
    If DCount("LocationName", "temptbl0") = 1 Then
        SelectLocation = DLookup("LocationName", "temptbl0")
        Call FindLocationRecord
        Call BriefDelay
    End If
    Exit Sub
ShowMeError:
    MsgBox "Error # " & Str(Err.Number) & " was generated by " & Me.Name & Chr(13) & Err.Description, vbOKOnly, "Error", Err.HelpFile, Err.HelpContext
End Sub

Private Sub SelectLocation_AfterUpdate()
    If Not (IsNull(SelectLocation) And IsNull(SelectCityTown)) Then
        Call FindLocationRecord
    End If
End Sub

Private Sub Department_AfterUpdate()
    If Department = "Residential Services" Then
        Cluster.Visible = True: Call Cluster_AfterUpdate: Cluster.ValidationRule = "Is Not Null"
        CommmunityRNName.Visible = True: LPN1Name.Visible = True: LPN2Name.Visible = True
    Else
        Cluster.Visible = False: Cluster.Value = Null: Cluster.ValidationRule = ""
        CommmunityRNName.Visible = False: LPN1Name.Visible = False: LPN2Name.Visible = False
    End If
End Sub

' Click actions.

Private Sub Button_FirstRecord_Click()
    SelectCityTown = "": SelectLocation = ""
    DoCmd.GoToRecord , "", acFirst
End Sub

Private Sub Button_NewRecord_Click()
    Call MoveToRecord(acNewRec)
    CityTown.Enabled = True: CityTown.Locked = False: LocationName.Enabled = True: LocationName.Locked = False
    CityTown.SetFocus
    Me.Caption = "New Place": RecordAddedDate = Format(Now(), "mm/dd/yyyy"): RecordAddedBy = Form_frmMainMenu.UserName
    NewRecordInProgress = True
End Sub

Private Sub FindRecord_Click()
    DoCmd.GoToControl Screen.PreviousControl.Name
    Err.Clear
    DoCmd.RunCommand acCmdFind
End Sub

Private Sub FindNextRecord_Click()
    DoCmd.FindNext
End Sub

Private Sub Button_UndoRecord_Click()
    DoCmd.RunCommand acCmdUndo
End Sub

Private Sub Button_DeleteRecord_Click()
    If MsgBox("Do you really want to delete the entry for " & CityTown & " - " & LocationName & ")?  This action will delete the Location record and all other supporting records in all other tables.", vbYesNo, "Confirm Deletion") = vbYes Then
' **** THERE'S A BUG IN THE NEXT STATEMENT! ***
        TILLDataBase.Execute "INSERT INTO tblDELETEDLocations ( CityTown, LocationName, RecordDeletedDate, RecordDeletedBy, GPName, Department, CostCenter, Cluster, ABI, NumClients, ResCapacity, EmailAddress, Address, City, State, ZIP, AddressValidated, PhoneNumber, DID, SpeedDial, FaxNumber, DDSArea, DDSRegion, " & _
            "StaffPrimaryContactIndexedName, StaffPrimaryContactLastName, StaffPrimaryContactFirstName, StaffPrimaryContactMiddleInitial, Leader, LOTR, ManagerLogonID, StaffSecondaryContactIndexedName, StaffSecondaryContactLastName, StaffSecondaryContactFirstName, StaffSecondaryContactMiddleInitial, StaffTertiaryContactIndexedName, StaffTertiaryContactLastName, StaffTertiaryContactFirstName, StaffTertiaryContactMiddleInitial, " & _
            "LastVehicleChecklistCompleted, MostRecentAsleepFireDrill, NextRecentAsleepFireDrill, DAYStaffTrainedInPrivacyBefore, DAYAllPlansReviewedByStaffBefore, DAYQtrlySafetyChecklistDueBy, HouseSafetyPlanExpires, HousePlansReviewedByStaffBefore, MAPChecklistCompleted, HumanRightsOfficer, HROTrainsStaffBefore, HROTrainsIndividualsBefore, " & _
            "FireSafetyOfficer, FSOTrainsStaffBefore, FSOTrainsIndividualsBefore, ReportExpirations, ResettlementCityTown, ResettlementLocation, Comments ) " & _
            "SELECT tblLocations.CityTown, tblLocations.LocationName, Format([Date],'mm/dd/yyyy') AS RecordDeletedDate, [Form_frmMainMenu].[UserName] AS RecordDeletedBy, tblLocations.GPName, tblLocations.Department, tblLocations.CostCenter, tblLocations.Cluster, tblLocations.ABI, tblLocations.NumClients, tblLocations.ResCapacity, tblLocations.EmailAddress, tblLocations.Address, tblLocations.City, tblLocations.State, tblLocations.ZIP, tblLocations.AddressValidated, tblLocations.PhoneNumber, tblLocations.DID, tblLocations.SpeedDial, tblLocations.FaxNumber, tblLocations.DDSArea, tblLocations.DDSRegion, " & _
            "tblLocations.StaffPrimaryContactIndexedName, tblLocations.StaffPrimaryContactLastName, tblLocations.StaffPrimaryContactFirstName, tblLocations.StaffPrimaryContactMiddleInitial, tblLocations.Leader, tblLocations.LOTR, tblLocations.ManagerLogonID, tblLocations.StaffSecondaryContactIndexedName, tblLocations.StaffSecondaryContactLastName, tblLocations.StaffSecondaryContactFirstName, tblLocations.StaffSecondaryContactMiddleInitial, tblLocations.StaffTertiaryContactIndexedName, tblLocations.StaffTertiaryContactLastName , tblLocations.StaffTertiaryContactFirstName, tblLocations.StaffTertiaryContactMiddleInitial, " & _
            "tblLocations.LastVehicleChecklistCompleted, tblLocations.MostRecentAsleepFireDrill, tblLocations.NextRecentAsleepFireDrill, tblLocations.DAYStaffTrainedInPrivacyBefore, tblLocations.DAYAllPlansReviewedByStaffBefore, tblLocations.DAYQtrlySafetyChecklistDueBy, tblLocations.HouseSafetyPlanExpires, tblLocations.HousePlansReviewedByStaffBefore, tblLocations.MAPChecklistCompleted, tblLocations.HumanRightsOfficer, tblLocations.HROTrainsStaffBefore, tblLocations.HROTrainsIndividualsBefore, " & _
            "tblLocations.FireSafetyOfficer, tblLocations.FSOTrainsStaffBefore, tblLocations.FSOTrainsIndividualsBefore, tblLocations.ReportExpirations, tblLocations.ResettlementCityTown, tblLocations.ResettlementLocation, tblLocations.Comments " & _
            "FROM tblLocations " & _
            "WHERE tblLocations.CityTown=[Forms]![frmLocations]![CityTown] AND tblLocations.LocationName=[Forms]![frmLocations]![LocationName];", dbSeeChanges: Call BriefDelay
        TILLDataBase.Execute "INSERT INTO tblDELETEDLocationsContacts ( CityTown, Location, IDX, Purpose, RecordDeletedDate, RecordDeletedBy, Global, Contact, PhoneNumber, MobilePhone, Address ) " & _
            "SELECT tblLocationsContacts.CityTown, tblLocationsContacts.Location, tblLocationsContacts.IDX, tblLocationsContacts.Purpose, Format([Date],""mm/dd/yyyy"") AS RecordDeletedDate, [Form_frmMainMenu].[UserName] AS RecordDeletedBy, tblLocationsContacts.Global, tblLocationsContacts.Contact, tblLocationsContacts.PhoneNumber, tblLocationsContacts.MobilePhone, tblLocationsContacts.Address " & _
            "FROM tblLocationsContacts " & _
            "WHERE tblLocations.CityTown=[Forms]![frmLocations]![CityTown] AND tblLocations.LocationName=[Forms]![frmLocations]![LocationName];", dbSeeChanges: Call BriefDelay
        TILLDataBase.Execute "DELETE * FROM tblLocations WHERE tblLocations.CityTown=[Forms]![frmLocations]![CityTown] AND tblLocations.LocationName=[Forms]![frmLocations]![LocationName];", dbSeeChanges: Call BriefDelay
        TILLDataBase.Execute "DELETE * FROM tblLocationsContacts WHERE tblLocations.CityTown=[Forms]![frmLocations]![CityTown] AND tblLocations.LocationName=[Forms]![frmLocations]![LocationName];", dbSeeChanges: Call BriefDelay
        DoCmd.GoToRecord , , acFirst
    End If
End Sub

Private Sub Report_Click()
    Me.Requery
    Call ExecReport("rptPROGCONT")
End Sub

Private Sub FilterCriteria_AfterUpdate()
    With Form_frmLocations
        Select Case FilterCriteria
            Case "<Show all>": Me.Filter = "": Me.FilterOn = False
            Case Else:         Me.Filter = "Department = '" & [FilterCriteria] & "'": Me.FilterOn = True
        End Select
    End With
End Sub