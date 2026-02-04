' Module Name: Form_frmPeopleScheduleStaffChanges
' Module Type: Document Module
' Lines of Code: 191
' Extracted: 2026-02-04 13:03:36

Option Compare Database
Option Explicit

Private Sub AddOfficeCityTown_AfterUpdate()
    Call BriefDelay
    If Not IsNull(AddOfficeCityTown) Or Len(AddOfficeCityTown) > 0 Then
        AddOfficeLocationName.RowSource = "SELECT DISTINCT tblLocations.LocationName FROM tblLocations WHERE tblLocations.CityTown = '" & AddOfficeCityTown & "' ORDER BY tblLocations.LocationName;"
    Else
        AddOfficeCityTown.SetFocus
    End If
End Sub

Private Sub CancelAddition_Click()
    Call Form_Current
End Sub

Private Sub CancelChange_Click()
    Call Form_Current
End Sub

Private Sub CancelDeletion_Click()
    Call Form_Current
End Sub

Private Sub DateOfChange_AfterUpdate()
    If DateValue(DateOfChange) < Int(Now) Then
        MsgBox "Change date is in past.  Please correct.", vbOKOnly, "Error!"
        DateOfChange = Null: DateOfChange.SetFocus
    End If
End Sub

Private Sub DeleteDate_AfterUpdate()
    If DateValue(DeleteDate) < Int(Now) Then
        MsgBox "Delete date is in past.  Please correct.", vbOKOnly, "Error!"
        DeleteDate = Null: DeleteDate.SetFocus
    End If
End Sub

Private Sub DeletePerson_AfterUpdate()
    DeleteIndexedName = DLookup("IndexedName", "qryPeopleAllPeopleRecords", "Person=""" & DeletePerson & """")
    DeleteFirstName = DLookup("FirstName", "tblPeople", "IndexedName=""" & DeleteIndexedName & """")
    DeleteMidInitial = DLookup("MiddleInitial", "tblPeople", "IndexedName=""" & DeleteIndexedName & """")
    DeleteLastName = DLookup("LastName", "tblPeople", "IndexedName=""" & DeleteIndexedName & """")
End Sub

Private Sub Form_Current()
    OfficeCityTown = Null: OfficeLocationName = Null: OfficeCityTown.SetFocus: CurrentFullName.Visible = False: NewFirstName.Visible = False: NewMidInitial.Visible = False
    NewLastName.Visible = False: DateOfChange.Visible = False: PositionIsOpen = False: PositionIsOpen.Visible = False: OpenPosition.Visible = False
    MakeChange.Visible = False: CancelChange.Visible = False

    DeletePerson = Null: DeleteIndexedName = Null: DeleteFirstName = Null: DeleteMidInitial = Null: DeleteLastName = Null: DeleteDate = Null
    AddOfficeCityTown = Null: AddOfficeLocationName = Null: AddFirstName = Null: AddMidInitial = Null: AddLastName = Null: AddJobTitle = Null
    AddDepartment = Null: AddDID = Null: AddHasPhone = False: AddExtPhone = Null: AddEmailAddress = Null: AddDate = Null
End Sub

Private Sub MakeAddition_Click()
    Dim DateOfChangeString As String
    ' Validate all data.
    If IsNull(AddDate) Then
        MsgBox "Add date is missing.", vbOKOnly, "Error!": AddDate.SetFocus
        Exit Sub
    End If
    ' Confirm to proceed.
    If MsgBox(AddFirstName & " " & AddLastName & " will be added on " & Format(AddDate, "Long Date") & "...OK?", vbYesNo, "Confirm?") = vbNo Then
        Call Form_Current
    End If
    ' Add the record.
    DateOfChangeString = Format(DateValue(AddDate), "mm/dd/yyyy")
    If IsNull(AddExtPhone) Or Len(AddExtPhone) <= 0 Then
        If IsNull(AddDID) Then
            TILLDataBase.Execute "INSERT INTO tblPeopleScheduledStaffChanges ( Action, AddOfficeCityTown, AddOfficeLocationName, AddFirstName, AddMidInitial, AddLastName, AddJobTitle, AddDepartment, AddHasPhone, AddEmailAddress, DateOfChange, Cancelled, Applied ) " & _
                "SELECT 'Add' AS Action, '" & AddOfficeCityTown & "' AS AddOfficeCityTown, '" & AddOfficeLocationName & "' AS AddOfficeLocationName, '" & AddFirstName & "' AS AddFirstName, '" & AddMidInitial & "' AS AddMidInitial, '" & AddLastName & "' AS AddLastName, '" & _
                AddJobTitle & "' AS AddJobtitle, '" & AddDepartment & "' AS AddDepartment, " & AddHasPhone & " AS AddHasPhone, '" & AddEmailAddress & "' AS AddEmailAddress, '" & _
                DateOfChangeString & "' AS DateOfChange, False AS Cancelled, False AS Applied ", dbSeeChanges: Call BriefDelay
        Else
            TILLDataBase.Execute "INSERT INTO tblPeopleScheduledStaffChanges ( Action, AddOfficeCityTown, AddOfficeLocationName, AddFirstName, AddMidInitial, AddLastName, AddJobTitle, AddDepartment, AddDID, AddHasPhone, AddEmailAddress, DateOfChange, Cancelled, Applied ) " & _
                "SELECT 'Add' AS Action, '" & AddOfficeCityTown & "' AS AddOfficeCityTown, '" & AddOfficeLocationName & "' AS AddOfficeLocationName, '" & AddFirstName & "' AS AddFirstName, '" & AddMidInitial & "' AS AddMidInitial, '" & AddLastName & "' AS AddLastName, '" & _
                AddJobTitle & "' AS AddJobtitle, '" & AddDepartment & "' AS AddDepartment, " & AddDID & " AS AddDID, " & AddHasPhone & " AS AddHasPhone, '" & AddEmailAddress & "' AS AddEmailAddress, '" & _
                DateOfChangeString & "' AS DateOfChange, False AS Cancelled, False AS Applied ", dbSeeChanges: Call BriefDelay
        End If
    Else
        If IsNull(AddDID) Then
            TILLDataBase.Execute "INSERT INTO tblPeopleScheduledStaffChanges ( Action, AddOfficeCityTown, AddOfficeLocationName, AddFirstName, AddMidinitial, AddLastName, AddJobTitle, AddDepartment, AddHasPhone, AddExtPhone, AddEmailAddress, DateOfChange, Cancelled, Applied ) " & _
                "SELECT 'Add' AS Action, '" & AddOfficeCityTown & "' AS AddOfficeCityTown, '" & AddOfficeLocationName & "' AS AddOfficeLocationName, '" & AddFirstName & "' AS AddFirstName, '" & AddMidInitial & "' AS AddMidInitial, '" & AddLastName & "' AS AddLastName, '" & _
                AddJobTitle & "' AS AddJobtitle, '" & AddDepartment & "' AS AddDepartment, " & AddHasPhone & " AS AddHasPhone, '" & AddExtPhone & "' AS AddExtPhone, '" & AddEmailAddress & "' AS AddEmailAddress, '" & _
                DateOfChangeString & "' AS DateOfChange, False AS Cancelled, False AS Applied ", dbSeeChanges: Call BriefDelay
        Else
            TILLDataBase.Execute "INSERT INTO tblPeopleScheduledStaffChanges ( Action, AddOfficeCityTown, AddOfficeLocationName, AddFirstName, AddMidinitial, AddLastName, AddJobTitle, AddDepartment, AddDID, AddHasPhone, AddExtPhone, AddEmailAddress, DateOfChange, Cancelled, Applied ) " & _
                "SELECT 'Add' AS Action, '" & AddOfficeCityTown & "' AS AddOfficeCityTown, '" & AddOfficeLocationName & "' AS AddOfficeLocationName, '" & AddFirstName & "' AS AddFirstName, '" & AddMidInitial & "' AS AddMidInitial, '" & AddLastName & "' AS AddLastName, '" & _
                AddJobTitle & "' AS AddJobtitle, '" & AddDepartment & "' AS AddDepartment, " & AddDID & " AS AddDID, " & AddHasPhone & " AS AddHasPhone, '" & AddExtPhone & "' AS AddExtPhone, '" & AddEmailAddress & "' AS AddEmailAddress, '" & _
                DateOfChangeString & "' AS DateOfChange, False AS Cancelled, False AS Applied ", dbSeeChanges: Call BriefDelay
        
        End If
    End If
    MsgBox "Addition scheduled.", vbOKOnly, "Scheduled."
    Call Form_Current
End Sub

Private Sub MakeChange_Click()
    Dim DateOfChangeString As String
    ' Validate all data.
    If IsNull(NewFirstName) Then
        MsgBox "New first name is missing.", vbOKOnly, "Error!": NewFirstName.SetFocus
        Exit Sub
    End If
    If IsNull(NewLastName) Then
        MsgBox "New last name is missing.", vbOKOnly, "Error!": NewLastName.SetFocus
        Exit Sub
    End If
    If IsNull(DateOfChange) Then
        MsgBox "Change date is missing.", vbOKOnly, "Error!": DateOfChange.SetFocus
        Exit Sub
    End If
    ' Confirm to proceed.
    If MsgBox(NewFirstName & " " & NewMidInitial & " " & NewLastName & " will be assigned to " & OfficeCityTown & "-" & OfficeLocationName & " on " & Format(DateOfChange, "Long Date") & "...OK?", vbYesNo, "Confirm?") = vbNo Then
        Call Form_Current
    End If
    ' Add the record.
    DateOfChangeString = Format(DateValue(DateOfChange), "mm/dd/yyyy")
    TILLDataBase.Execute "INSERT INTO tblPeopleScheduledStaffChanges ( Action, OfficeCityTown, OfficeLocationName, CurrentFirstName, CurrentMidInitial, CurrentLastName, NewFirstName, NewMidInitial, NewLastName, DateOfChange, Cancelled, Applied ) " & _
        "SELECT 'Change' AS Action, '" & OfficeCityTown & "' AS OfficeCityTown, '" & OfficeLocationName & "' AS OfficeLocationName, '" & _
        CurrentFirstName & "' AS CurrentFirstName, '" & CurrentMidInitial & "' AS CurrentMidinitial, '" & CurrentLastName & "' AS CurrentLastName, '" & NewFirstName & "' AS NewFirstName, '" & NewMidInitial & "' AS NewMidInitial, '" & NewLastName & "' AS NewLastName, '" & DateOfChangeString & "' AS DateOfChange, False AS Cancelled, False AS Applied ", dbSeeChanges: Call BriefDelay
    MsgBox "Change scheduled.", vbOKOnly, "Scheduled."
    Call Form_Current
End Sub

Private Sub MakeDeletion_Click()
    Dim DateOfChangeString As String
    ' Validate all data.
    If IsNull(DeleteIndexedName) Then
        MsgBox "No name selected.", vbOKOnly, "Error!": DeletePerson.SetFocus
        Exit Sub
    End If
    If IsNull(DeleteDate) Then
        MsgBox "Deletion date is missing.", vbOKOnly, "Error!": DeleteDate.SetFocus
        Exit Sub
    End If
    ' Confirm to proceed.
    If MsgBox(DeleteFirstName & " " & DeleteMidInitial & " " & DeleteLastName & " will be deleted on " & Format(DeleteDate, "Long Date") & "...OK?", vbYesNo, "Confirm?") = vbNo Then
        Call Form_Current
    End If
    ' Add the record.
    DateOfChangeString = Format(DateValue(DeleteDate), "mm/dd/yyyy")
    TILLDataBase.Execute "INSERT INTO tblPeopleScheduledStaffChanges ( Action, DeleteFirstName, DeleteMidInitial, DeleteLastName, DeleteIndexedName, DateOfChange, Cancelled, Applied ) " & _
        "SELECT ""Delete"" AS Action, """ & DeleteFirstName & """ AS DeleteFirstName, """ & DeleteMidInitial & """ AS DeleteMidinitial, """ & DeleteLastName & """ AS DeleteLastName, """ & DeleteIndexedName & """ AS DeleteIndexedName, """ & DateOfChangeString & """ AS DateOfChange, False AS Cancelled, False AS Applied ", dbSeeChanges: Call BriefDelay
    MsgBox "Deletion scheduled.", vbOKOnly, "Scheduled."
    Call Form_Current
End Sub

Private Sub OfficeCityTown_AfterUpdate()
    OfficeLocationName.Requery
    If IsNull(OfficeLocationName) Or Len(OfficeLocationName) <= 0 Then Exit Sub
    Call OfficeFilled
End Sub

Private Sub OfficeLocationName_AfterUpdate()
    If IsNull(OfficeCityTown) Or Len(OfficeCityTown) <= 0 Then Exit Sub
    Call OfficeFilled
End Sub

Private Sub OfficeFilled()
    Dim Criteria As Variant
    
    Criteria = "OfficeCityTown = '" & OfficeCityTown & "' AND OfficeLocationName = '" & OfficeLocationName & _
        "' AND ((Department = 'Residential Services' AND (StaffTitle = 'Residence Manager' OR StaffTitle = 'Site Coordinator')) OR (Department = 'Day Services' AND (StaffTitle = 'Day Hab Manager' OR StaffTitle = 'TILL Central Site Supervisor')))"
    CurrentFirstName = DLookup("FirstName", "tblPeople", Criteria): CurrentLastName = DLookup("LastName", "tblPeople", Criteria)
    
    If Len(CurrentFirstName) <= 0 And Len(CurrentLastName) <= 0 Then
        MsgBox "There is no staff entry for this location.  You will need to add a new record.", vbOKOnly, "Error!"
        OfficeCityTown = Null: OfficeLocationName = Null: OfficeCityTown.SetFocus
        Exit Sub
    End If
    
    If CurrentFirstName = "TBD" Then CurrentFullName = "*** OPEN ***" Else CurrentFullName = CurrentFirstName & " " & CurrentLastName
    CurrentFullName.Visible = True: NewFirstName.Visible = True: NewFirstName = Null: NewFirstName.SetFocus: NewMidInitial.Visible = True: NewMidInitial = Null: NewLastName.Visible = True
    NewLastName = Null: PositionIsOpen.Visible = True: PositionIsOpen = False: DateOfChange.Visible = True: DateOfChange = Null
    MakeChange.Visible = True: CancelChange.Visible = True
End Sub

Private Sub PositionIsOpen_Click()
    If PositionIsOpen Then
        OpenPosition.Visible = True: NewFirstName.Visible = False: NewFirstName = "TBD": NewLastName.Visible = False
        If OfficeLocationName = "Day Hab" Or Left(OfficeLocationName, 12) = "TILL Central" Then NewLastName = OfficeCityTown & " " & OfficeLocationName Else NewLastName = OfficeLocationName
    Else
        NewFirstName.Visible = True: NewFirstName = Null: NewFirstName.SetFocus: NewLastName.Visible = True: NewLastName = Null
    End If
End Sub

Private Sub ShowScheduled_Click()
    DoCmd.OpenForm "frmPeopleShowScheduleStaffChanges"
End Sub
