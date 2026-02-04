' Module Name: Form_frmPeopleClientsSelectLocation
' Module Type: Document Module
' Lines of Code: 109
' Extracted: 2026-02-04 13:03:35

Option Compare Database
Option Explicit

Private Department As String

Private Sub Form_Current()
    SelectedCityTown.SetFocus
End Sub

Private Sub Form_Load()
    Select Case Me.OpenArgs
        Case "Res": Department = "Residential Services"
        Case "CLO": Department = "Individualized Support Options"
        Case "Day": Department = "Day Services"
        Case "Voc": Department = "Vocational Services"
        Case Else
    End Select
    Form_frmPeople.DeptCriteria = Department
    SelectedCityTown.RowSource = "SELECT DISTINCT tblLocations.CityTown FROM tblLocations WHERE tblLocations.Department = """ & Department & """ ORDER BY tblLocations.CityTown;"
    SelectedCityTown.Requery
End Sub

Private Sub SelectedCityTown_AfterUpdate()
    Form_frmPeople.CityPicker = SelectedCityTown
    SelectedLocation.RowSource = "SELECT tblLocations.LocationName, tblLocations.CityTown FROM tblLocations WHERE tblLocations.CityTown = """ & Form_frmPeople.CityPicker & """ And tblLocations.Department = """ & Department & """ ORDER BY tblLocations.LocationName;"
    SelectedLocation.Requery
End Sub

Private Sub SelectOK_Click()
On Error GoTo ShowMeError
    Select Case Me.OpenArgs
        Case "Res"
            If DLookup("ABI", "tblLocations", "ABI=True AND CityTown= '" & SelectedCityTown & "'") Then
                Form_frmPeopleClientsServiceResidential.Funding = Null
            Else
                Form_frmPeopleClientsServiceResidential.Funding = DLookup("Area", "catDDSAreasAndRegions", "CityTown='" & SelectedCityTown & "'")
            End If
            Call UpdateChangeLog("RESFundingSource", Form_frmPeopleClientsServiceResidential.Funding)
            
            DoCmd.Close acForm, "frmPeopleClientsServiceResidential"
            TILLDataBase.Execute "UPDATE qrytblPeopleClientsResidentialServices SET qrytblPeopleClientsResidentialServices.CityTown = """ & SelectedCityTown & """, qrytblPeopleClientsResidentialServices.Location = """ & SelectedLocation & """" & _
                "WHERE qrytblPeopleClientsResidentialServices.IndexedName=""" & Form_frmPeople.IndexedName & """;"
            Call UpdateChangeLog("RESCityTown", SelectedCityTown)
            Call UpdateChangeLog("RESLocation", SelectedLocation)
            Form_frmPeople.ResLocation = SelectedCityTown & "-" & SelectedLocation
            DoCmd.OpenForm "frmPeopleClientsServiceResidential", , , "IndexedName=""" & Form_frmPeople.IndexedName & """"
        Case "CLO"
            DoCmd.Close acForm, "frmPeopleClientsServiceCLO"
            TILLDataBase.Execute "UPDATE qrytblPeopleClientsCLOServices SET qrytblPeopleClientsCLOServices.CityTown = """ & SelectedCityTown & """, qrytblPeopleClientsCLOServices.Location = """ & SelectedLocation & """" & _
                "WHERE qrytblPeopleClientsCLOServices.IndexedName=""" & Form_frmPeople.IndexedName & """;"
            Call UpdateChangeLog("CLOCityTown", SelectedCityTown)
            Call UpdateChangeLog("CLOLocation", SelectedLocation)
            Form_frmPeople.CLOLocation = SelectedCityTown & "-" & SelectedLocation
            DoCmd.OpenForm "frmPeopleClientsServiceCLO", , , "IndexedName=""" & Form_frmPeople.IndexedName & """"
        Case "Day"
            DoCmd.Close acForm, "frmPeopleClientsServiceDay"
            TILLDataBase.Execute "UPDATE qrytblPeopleClientsDayServices SET qrytblPeopleClientsDayServices.CityTown = """ & SelectedCityTown & """, qrytblPeopleClientsDayServices.LocationName = """ & SelectedLocation & """" & _
                "WHERE qrytblPeopleClientsDayServices.IndexedName=""" & Form_frmPeople.IndexedName & """;"
            Call UpdateChangeLog("DAYCityTown", SelectedCityTown)
            Call UpdateChangeLog("DAYLocation", SelectedLocation)
            If SelectedLocation = "Day Hab" Then Form_frmPeople.DeptCriteria = "Day Habilitation" Else Form_frmPeople.DeptCriteria = "TILL Central"
            Form_frmPeople.DayLocation = SelectedCityTown & "-" & SelectedLocation
            DoCmd.OpenForm "frmPeopleClientsServiceDay", , , "IndexedName=""" & Form_frmPeople.IndexedName & """"
        Case "Voc"
            DoCmd.Close acForm, "frmPeopleClientsServiceVocational"
            TILLDataBase.Execute "UPDATE qrytblPeopleClientsVocationalServices SET qrytblPeopleClientsVocationalServices.CityTown = """ & SelectedCityTown & """, qrytblPeopleClientsVocationalServices.Location = """ & SelectedLocation & """" & _
                "WHERE qrytblPeopleClientsVocationalServices.IndexedName=""" & Form_frmPeople.IndexedName & """;"
            Call UpdateChangeLog("VOCCityTown", SelectedCityTown)
            Call UpdateChangeLog("VOCLocation", SelectedLocation)
            Form_frmPeople.VocLocation = SelectedCityTown & "-" & SelectedLocation
            DoCmd.OpenForm "frmPeopleClientsServiceVocational", , , "IndexedName=""" & Form_frmPeople.IndexedName & """"
        Case Else
    End Select
    
    Select Case Me.OpenArgs
        Case "Res", "CLO"
            With Form_frmPeople
                .PhysicalAddress = DLookup("Address", "tblLocations", "CityTown = """ & [SelectedCityTown] & """ AND LocationName = """ & [SelectedLocation] & """")
                .PhysicalAddressValidated = True
                Call UpdateChangeLog("PhysicalAddress", .PhysicalAddress)
                .PhysicalCity = DLookup("City", "tblLocations", "CityTown = """ & [SelectedCityTown] & """ AND LocationName = """ & [SelectedLocation] & """")
                Call UpdateChangeLog("PhysicalCity", .PhysicalCity)
                .PhysicalState = DLookup("State", "tblLocations", "CityTown = """ & [SelectedCityTown] & """ AND LocationName = """ & [SelectedLocation] & """")
                Call UpdateChangeLog("PhysicalState", .PhysicalState)
                .PhysicalZIP = DLookup("ZIP", "tblLocations", "CityTown = """ & [SelectedCityTown] & """ AND LocationName = """ & [SelectedLocation] & """")
                Call UpdateChangeLog("PhysicalZIP", .PhysicalZIP)
                .MailingAddress = DLookup("Address", "tblLocations", "CityTown = """ & [SelectedCityTown] & """ AND LocationName = """ & [SelectedLocation] & """")
                .MailingAddressValidated = True
                Call UpdateChangeLog("MailingAddress", .MailingAddress)
                .MailingCity = DLookup("City", "tblLocations", "CityTown = """ & [SelectedCityTown] & """ AND LocationName = """ & [SelectedLocation] & """")
                Call UpdateChangeLog("MailingCity", .MailingCity)
                .MailingState = DLookup("State", "tblLocations", "CityTown = """ & [SelectedCityTown] & """ AND LocationName = """ & [SelectedLocation] & """")
                Call UpdateChangeLog("MailingState", .MailingState)
                .MailingZIP = DLookup("ZIP", "tblLocations", "CityTown = """ & [SelectedCityTown] & """ AND LocationName = """ & [SelectedLocation] & """")
                Call UpdateChangeLog("MailingState", .MailingState)
                .HomePhone = DLookup("PhoneNumber", "tblLocations", "CityTown = """ & [SelectedCityTown] & """ AND LocationName = """ & [SelectedLocation] & """")
                Call UpdateChangeLog("HomePhone", .HomePhone)
                Call UpdateCounty
'               If .Dirty Then .Dirty = False
            End With
        Case Else
    End Select
    
    DoCmd.Close acForm, "frmPeopleClientsSelectLocation"
    Exit Sub
ShowMeError:
    MsgBox "Error # " & Str(Err.Number) & " was generated by " & Me.Name & Chr(13) & Err.Description, vbOKOnly, "Error", Err.HelpFile, Err.HelpContext
End Sub

