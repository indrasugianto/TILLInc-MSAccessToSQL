' Module Name: Form_frmRptExpMAILINGSANDSPREADSHEETS
' Module Type: Document Module
' Lines of Code: 1084
' Extracted: 2026-02-04 13:03:35

Option Compare Database
Option Explicit

Private Sub ChooseCompanies_Click()
    If ChooseCompanies Then
        ChooseCompanies = True
        ChoosePeople = False
        SelectDonor.RowSource = "SELECT DISTINCT qryDonations.CompanyOrganization AS FullName, qryDonations.IndexedName, qryDonations.LastName, qryDonations.FirstName FROM qryDonations WHERE qryDonations.IndexedName Like '///*';"
    Else
        ChooseCompanies = False
        ChoosePeople = True
        SelectDonor.RowSource = "SELECT DISTINCT [LastName] & ', ' & [FirstName] AS FullName, qryDonations.IndexedName FROM qryDonations WHERE qryDonations.IndexedName Not Like '///*' ORDER BY [LastName] & ', ' & [FirstName];"
    End If
    SelectDonor = ""
    SelectDonor.Requery
End Sub

Private Sub ChoosePeople_Click()
    If ChoosePeople Then
        ChoosePeople = True
        ChooseCompanies = False
        SelectDonor.RowSource = "SELECT DISTINCT [LastName] & ', ' & [FirstName] AS FullName, qryDonations.IndexedName FROM qryDonations WHERE qryDonations.IndexedName Not Like '///*' ORDER BY [LastName] & ', ' & [FirstName];"
    Else
        ChoosePeople = False
        ChooseCompanies = True
        SelectDonor.RowSource = "SELECT DISTINCT qryDonations.CompanyOrganization AS FullName, qryDonations.IndexedName, qryDonations.LastName, qryDonations.FirstName FROM qryDonations WHERE qryDonations.IndexedName Like '///*';"
    End If
    SelectDonor = ""
    SelectDonor.Requery
End Sub

Private Sub DayClientProfiles_Click()
    Call BasicExport("SELECT tblPeople.LastName, tblPeople.FirstName, tblPeopleClientsDemographics.DateOfBirth, tblPeopleClientsDayServices.CityTown, tblPeopleClientsDayServices.LocationName AS Program, tblPeopleClientsDayServices.Profile, tblPeopleClientsDayServices.Severity, tblPeopleClientsDayServices.ScoresUpdatedWhen AS DateProfileUpdated, tblPeopleClientsDayServices.ScoresUpdatedWho AS ProfileUpdatedBy INTO temptbl " & _
        "FROM (tblPeople INNER JOIN tblPeopleClientsDemographics ON tblPeople.IndexedName = tblPeopleClientsDemographics.IndexedName) INNER JOIN tblPeopleClientsDayServices ON tblPeopleClientsDemographics.IndexedName = tblPeopleClientsDayServices.IndexedName " & _
        "WHERE tblPeople.IsClientDay=True AND tblPeople.IsDeceased=False AND tblPeopleClientsDayServices.Inactive=False;", "DayClientProfiles")
End Sub

Private Sub DAYCLIENTSONLY_Click()
    Call BasicExport("SELECT tblPeople.LastName, tblPeople.FirstName, tblPeople.MiddleInitial, tblPeople.MailingAddress, tblPeople.MailingCity, tblPeople.MailingState, tblPeople.MailingZIP, tblPeopleClientsVendors.ResidentialVendor, tblPeopleClientsVendors.ResVendorAddress, tblPeopleClientsVendors.ResVendorCity, tblPeopleClientsVendors.ResVendorState, tblPeopleClientsVendors.ResVendorZIP, tblPeopleClientsVendors.ResidentialVendorPhoneNumber, tblPeopleClientsVendors.ResVendorLocation, tblPeopleClientsVendors.LivingWithParentOrGuardian, tblPeopleClientsVendors.LivingIndependently, tblPeople.DayLocation, tblPeople.VocLocation INTO temptbl " & _
        "FROM ((((tblPeople LEFT JOIN tblPeopleClientsDayServices ON tblPeople.IndexedName = tblPeopleClientsDayServices.IndexedName) INNER JOIN tblPeopleClientsVendors ON tblPeople.IndexedName = tblPeopleClientsVendors.IndexedName) LEFT JOIN tblPeopleClientsCLOServices ON tblPeople.IndexedName = tblPeopleClientsCLOServices.IndexedName) LEFT JOIN tblPeopleClientsVocationalServices ON tblPeople.IndexedName = tblPeopleClientsVocationalServices.IndexedName) LEFT JOIN tblPeopleClientsResidentialServices ON tblPeople.IndexedName = tblPeopleClientsResidentialServices.IndexedName " & _
        "WHERE (((tblPeople.IsDeceased)=False) AND ((tblPeople.NoMailings)=False) AND ((tblPeople.IsClientDay)=True) AND ((tblPeopleClientsDayServices.Inactive)=False) AND ((tblPeople.IsCilentCLO)=False) AND ((tblPeople.IsClientRes)=False)) OR (((tblPeople.IsDeceased)=False) AND ((tblPeople.NoMailings)=False) AND ((tblPeople.IsClientDay)=True) AND ((tblPeopleClientsDayServices.Inactive)=False) AND ((tblPeople.IsCilentCLO)=True) AND ((tblPeopleClientsCLOServices.Inactive)=True) AND ((tblPeople.IsClientRes)=False)) OR (((tblPeople.IsDeceased)=False) AND ((tblPeople.NoMailings)=False) AND ((tblPeople.IsClientDay)=True) AND ((tblPeopleClientsDayServices.Inactive)=False) " & _
        "AND ((tblPeople.IsCilentCLO)=False) AND ((tblPeople.IsClientRes)=True) AND ((tblPeopleClientsResidentialServices.Inactive)=True)) OR (((tblPeople.IsDeceased)=False) AND ((tblPeople.NoMailings)=False) AND ((tblPeople.IsClientDay)=True) AND ((tblPeopleClientsDayServices.Inactive)=False) AND ((tblPeople.IsCilentCLO)=True) AND ((tblPeopleClientsCLOServices.Inactive)=True) AND ((tblPeople.IsClientRes)=True) AND ((tblPeopleClientsResidentialServices.Inactive)=True)) OR (((tblPeople.IsDeceased)=False) AND ((tblPeople.NoMailings)=False) AND ((tblPeople.IsClientVocat)=True) AND ((tblPeopleClientsVocationalServices.Inactive)=False) AND ((tblPeople.IsCilentCLO)=False) AND ((tblPeople.IsClientRes)=False)) OR (((tblPeople.IsDeceased)=False) " & _
        "AND ((tblPeople.NoMailings)=False) AND ((tblPeople.IsClientVocat)=True) AND ((tblPeopleClientsVocationalServices.Inactive)=False) AND ((tblPeople.IsCilentCLO)=True) AND ((tblPeopleClientsCLOServices.Inactive)=True) AND ((tblPeople.IsClientRes)=False)) OR (((tblPeople.IsDeceased)=False) AND ((tblPeople.NoMailings)=False) AND ((tblPeople.IsClientVocat)=True) AND ((tblPeopleClientsVocationalServices.Inactive)=False) AND ((tblPeople.IsCilentCLO)=False) AND ((tblPeople.IsClientRes)=True) AND ((tblPeopleClientsResidentialServices.Inactive)=True)) OR (((tblPeople.IsDeceased)=False) AND ((tblPeople.NoMailings)=False) AND ((tblPeople.IsClientVocat)=True) AND ((tblPeopleClientsVocationalServices.Inactive)=False) AND ((tblPeople.IsCilentCLO)=True) AND ((tblPeopleClientsCLOServices.Inactive)=True) AND ((tblPeople.IsClientRes)=True) AND ((tblPeopleClientsResidentialServices.Inactive)=True));", "DayClientsNotRes")
End Sub

Private Sub DonationsEndDate_AfterUpdate()
    DonationsEndDateNumeric = DateValue(DonationsEndDate)
End Sub

Private Sub DonationsStartDate_AfterUpdate()
    DonationsStartDateNumeric = DateValue(DonationsStartDate)
End Sub

Private Sub Form_Load()
    DNRAPPEAL.Visible = DCount("Action", "catUserPermissions", "User='" & Form_frmMainMenu.UserName & "' AND Action='Can access Donor Appeal'") > 0
    
    ChoosePeople = True: ChooseCompanies = False
    SelectDonor.RowSource = "SELECT DISTINCT [LastName] & ', ' & [FirstName] AS FullName, tblPeople.IndexedName FROM tblPeopleDonors RIGHT JOIN tblPeople ON tblPeopleDonors.IndexedName = tblPeople.IndexedName WHERE (((tblPeople.LastName) Is Not Null) AND ((tblPeople.FirstName) Is Not Null And (tblPeople.FirstName) Is Not Null) AND ((Len([FirstName]))>0) AND ((Len([LastName]))>0) AND ((tblPeople.IsDonor)=True)) ORDER BY [LastName] & ', ' & [FirstName];"
    DonationsStartDate = "01/01/2000": DonationsStartDateNumeric = DateValue(DonationsStartDate): DonationsEndDate = Format(Date, "mm/dd/yyyy"): DonationsEndDateNumeric = Date
    
End Sub

Private Sub Form_Current()
    SelectedCityTownDay.SetFocus
End Sub

Private Sub AllDayLocations_Click()
On Error GoTo ShowMeError
    Call DropTempTables
    
    SysCmdResult = SysCmd(4, "Creating client table.")
    TILLDataBase.Execute "CREATE TABLE temptbl0 (ClientIndexedName CHAR(160), ClientLastName CHAR(25), ClientFirstName CHAR(25), ClientMiddleInitial CHAR(1), FamilyIndexedName CHAR(160), Relationship CHAR(25), Guardian BIT, PrimaryContact BIT, IsClientCLO BIT, IsClientRes BIT, IsDeceased BIT, NoMailings BIT, FamilyInactive BIT);", dbSeeChanges: Call BriefDelay
    TILLDataBase.Execute "INSERT INTO temptbl0 (ClientIndexedName, ClientLastName, ClientFirstName, ClientMiddleInitial, FamilyIndexedName, Relationship, Guardian, PrimaryContact, IsClientCLO, IsClientRes, IsDeceased, NoMailings, FamilyInactive ) " & _
        "SELECT tblPeople.IndexedName, tblPeople.LastName, tblPeople.FirstName, tblPeople.MiddleInitial, tblPeopleFamily.IndexedName, tblPeopleFamily.Relationship, tblPeopleFamily.Guardian, tblPeopleFamily.PrimaryContact, tblPeople.IsCilentCLO, tblPeople.IsClientRes, tblPeople.IsDeceased, tblPeople.NoMailings, tblPeopleFamily.Inactive " & _
        "FROM ((tblPeople LEFT JOIN tblPeopleFamily ON tblPeople.IndexedName = tblPeopleFamily.ClientIndexedName) LEFT JOIN tblPeopleClientsDayServices ON tblPeople.IndexedName = tblPeopleClientsDayServices.IndexedName) LEFT JOIN tblPeopleClientsVocationalServices ON tblPeople.IndexedName = tblPeopleClientsVocationalServices.IndexedName " & _
        "WHERE (tblPeopleFamily.PrimaryContact=True AND tblPeople.IsClientDay=True   AND tblPeopleClientsDayServices.Inactive=False) OR " & _
              "(tblPeopleFamily.Guardian=True       AND tblPeople.IsClientDay=True   AND tblPeopleClientsDayServices.Inactive=False       ) OR " & _
              "(tblPeopleFamily.PrimaryContact=True AND tblPeople.IsClientVocat=True AND tblPeopleClientsVocationalServices.Inactive=False) OR " & _
              "(tblPeopleFamily.Guardian=True       AND tblPeople.IsClientVocat=True AND tblPeopleClientsVocationalServices.Inactive=False);", dbSeeChanges: Call BriefDelay
    
    TILLDataBase.Execute "DELETE temptbl0.* FROM temptbl0 WHERE temptbl0.IsDeceased=True OR temptbl0.NoMailings=True OR temptbl0.FamilyInactive=True;", dbSeeChanges: Call BriefDelay
    
    SysCmdResult = SysCmd(4, "Creating family table.")
    TILLDataBase.Execute "CREATE TABLE temptbl1 (ClientIndexedName CHAR(160), FamilyIndexedName CHAR(160), Location CHAR(50), ClientLastName CHAR(25), ClientFirstName CHAR(25), ClientMiddleInitial CHAR(1), Relationship CHAR(25), PrimaryContact BIT, Guardian BIT, LN CHAR(25), FN CHAR(25), MI CHAR(1), CO CHAR(40), FamilyFamiliarGreeting CHAR(40), MailingAddress CHAR(50), MailingCity CHAR(25), MailingState CHAR(2), MailingZIP CHAR(10), GoGreen BIT, EmailAddress CHAR(50));", dbSeeChanges: Call BriefDelay
    TILLDataBase.Execute "ALTER TABLE temptbl1 ADD CONSTRAINT temptbl1constraint PRIMARY KEY (ClientIndexedName, FamilyIndexedName);", dbSeeChanges: Call BriefDelay
    TILLDataBase.Execute "INSERT INTO temptbl1 (ClientIndexedName, FamilyIndexedName, Location, ClientLastName, ClientFirstName, ClientMiddleInitial, Relationship, PrimaryContact, Guardian, LN, FN, MI, CO, FamilyFamiliarGreeting, MailingAddress, MailingCity, MailingState, MailingZIP, GoGreen, EmailAddress ) " & _
        "SELECT ClientIndexedName, FamilyIndexedName, tblPeopleClientsVendors.DayVendorLocation, ClientLastName, ClientFirstName, ClientMiddleInitial, Relationship, PrimaryContact, Guardian, tblPeople.LastName AS LN, tblPeople.FirstName AS FN, " & _
        "tblPeople.MiddleInitial AS MI, tblPeople.CompanyOrganization AS CO, tblPeople.FamiliarGreeting AS FamilyFamiliarGreeting, tblPeople.MailingAddress AS MailingAddress, tblPeople.MailingCity AS MailingCity, tblPeople.MailingState AS MailingState, tblPeople.MailingZIP AS MailingZIP, tblPeople.GoGreen, tblPeople.EmailAddress " & _
        "FROM (temptbl0 INNER JOIN tblPeopleClientsVendors ON temptbl0.ClientIndexedName = tblPeopleClientsVendors.IndexedName) INNER JOIN tblPeople ON temptbl0.FamilyIndexedName = tblPeople.IndexedName " & _
        "WHERE tblPeople.IsDeceased=False AND Len(tblPeople.MailingAddress) > 0;", dbSeeChanges: Call BriefDelay
    
    SysCmdResult = SysCmd(4, "Building export table.")
    TILLDataBase.Execute "CREATE TABLE temptbl2 (MailingAddress CHAR(50), MailingCity CHAR(25), MailingState CHAR(2), MailingZIP CHAR(10), ClientIndexedName CHAR(160), Location CHAR(50), LN1 CHAR(25), " & _
        "FN1 CHAR(25), MI1 CHAR(1), CO1 CHAR(50), FamilyFamiliarGreeting1 CHAR(30),  FamilyMemberIsGuardian1 BIT, FamilyGoGreen1 BIT, EmailAddress1 CHAR(50), LN2 CHAR(25), FN2 CHAR(25), MI2 CHAR(1), " & _
        "CO2 CHAR(50), FamilyFamiliarGreeting2 CHAR(30), FamilyMemberIsGuardian2 BIT, FamilyGoGreen2 BIT, EmailAddress2 CHAR(50), LN3 CHAR(25), FN3 CHAR(25), MI3 CHAR(1), CO3 CHAR(50), FamilyFamiliarGreeting3 CHAR(30), FamilyMemberIsGuardian3 BIT, FamilyGoGreen3 BIT, EmailAddress3 CHAR(50), LN4 CHAR(25), FN4 CHAR(25), MI4 CHAR(1), CO4 CHAR(50), FamilyFamiliarGreeting4 CHAR(30), FamilyMemberIsGuardian4 BIT, FamilyGoGreen4 BIT, EmailAddress4 CHAR(50) );", dbSeeChanges: Call BriefDelay
    TILLDataBase.Execute "ALTER TABLE temptbl2 ADD CONSTRAINT temptbl2constraint PRIMARY KEY (MailingAddress, MailingCity, MailingState, MailingZIP, ClientIndexedName);", dbSeeChanges: Call BriefDelay
    
    TILLDataBase.Execute "INSERT INTO temptbl2 (MailingAddress, MailingCity, MailingState, MailingZIP, ClientIndexedName, Location) " & _
        "SELECT MailingAddress, MailingCity, MailingState, MailingZIP, ClientIndexedName, Location FROM temptbl1;", dbSeeChanges: Call BriefDelay
    
    Call ParseFamilyMembersAndThenExport(3, "temptbl2", "DayProgramNewsletters-AllLocations")
    SysCmdResult = SysCmd(5)
    Exit Sub
ShowMeError:
    MsgBox "Error # " & Str(Err.Number) & " was generated by " & Me.Name & Chr(13) & Err.Description, vbOKOnly, "Error", Err.HelpFile, Err.HelpContext
End Sub

Private Sub AllResidenceLocations_Click()
On Error GoTo ShowMeError
    Call DropTempTables
    
    SysCmdResult = SysCmd(4, "Creating client table.")
    TILLDataBase.Execute "CREATE TABLE temptbl0 (ClientIndexedName CHAR(160), ClientLastName CHAR(25), ClientFirstName CHAR(25), ClientMiddleInitial CHAR(1), FamilyIndexedName CHAR(160), Guardian BIT);", dbSeeChanges: Call BriefDelay
    TILLDataBase.Execute "INSERT INTO temptbl0 (ClientIndexedName, ClientLastName, ClientFirstName, ClientMiddleInitial, FamilyIndexedName, Guardian) " & _
        "SELECT tblPeople.IndexedName, tblPeople.LastName, tblPeople.FirstName, tblPeople.MiddleInitial, tblPeopleFamily.IndexedName, tblPeopleFamily.Guardian " & _
        "FROM ((tblPeople LEFT JOIN tblPeopleFamily ON tblPeople.IndexedName = tblPeopleFamily.ClientIndexedName) LEFT JOIN tblPeopleClientsResidentialServices ON tblPeople.IndexedName = tblPeopleClientsResidentialServices.IndexedName) LEFT JOIN tblPeopleClientsCLOServices ON tblPeople.IndexedName = tblPeopleClientsCLOServices.IndexedName " & _
        "WHERE tblPeople.NoMailings=False AND tblPeople.IsDeceased=False AND tblPeopleFamily.Inactive=False AND " & _
             "((tblPeopleFamily.PrimaryContact=True AND tblPeople.IsClientRes=True AND tblPeopleClientsResidentialServices.Inactive=False) OR " & _
              "(tblPeopleFamily.Guardian=True       AND tblPeople.IsClientRes=True AND tblPeopleClientsResidentialServices.Inactive=False) OR " & _
              "(tblPeopleFamily.PrimaryContact=True AND tblPeople.IsCilentCLO=True AND tblPeopleClientsCLOServices.Inactive=False) OR " & _
              "(tblPeopleFamily.Guardian=True       AND tblPeople.IsCilentCLO=True AND tblPeopleClientsCLOServices.Inactive=False));", dbSeeChanges: Call BriefDelay
    SysCmdResult = SysCmd(4, "Adding vendor information.")
    TILLDataBase.Execute "CREATE TABLE temptbl1 (ClientIndexedName CHAR(160), Location CHAR(50), ClientLastName CHAR(25), ClientFirstName CHAR(25), ClientMiddleInitial CHAR(1), FamilyIndexedName CHAR(160), Relationship CHAR(25), PrimaryContact BIT, Guardian BIT, LN CHAR(25), FN CHAR(25), MI CHAR(1), CO CHAR(40), FamilyFamiliarGreeting CHAR(50), MailingAddress CHAR(50), MailingCity CHAR(25), MailingState CHAR(2), MailingZIP CHAR(10), GoGreen BIT, EmailAddress CHAR(50));", dbSeeChanges: Call BriefDelay
    TILLDataBase.Execute "INSERT INTO temptbl1 (ClientIndexedName, Location, ClientLastName, ClientFirstName, ClientMiddleInitial, FamilyIndexedName, Guardian, LN, FN, MI, CO, FamilyFamiliarGreeting, MailingAddress, MailingCity, MailingState, MailingZIP, GoGreen, EmailAddress) " & _
        "SELECT ClientIndexedName, tblPeopleClientsVendors.ResVendorLocation, ClientLastName, ClientFirstName, ClientMiddleInitial, FamilyIndexedName, Guardian, tblPeople.LastName AS LN, tblPeople.FirstName AS FN, " & _
        "tblPeople.MiddleInitial AS MI, tblPeople.CompanyOrganization AS CO, tblPeople.FamiliarGreeting AS FamilyFamiliarGreeting, tblPeople.MailingAddress AS MailingAddress, tblPeople.MailingCity AS MailingCity, tblPeople.MailingState AS MailingState, tblPeople.MailingZIP AS MailingZIP, tblPeople.GoGreen, tblPeople.EmailAddress " & _
        "FROM (temptbl0 INNER JOIN tblPeopleClientsVendors ON temptbl0.ClientIndexedName = tblPeopleClientsVendors.IndexedName) INNER JOIN tblPeople ON temptbl0.FamilyIndexedName = tblPeople.IndexedName " & _
        "WHERE tblPeople.IsDeceased=False;", dbSeeChanges: Call BriefDelay
    SysCmdResult = SysCmd(4, "Building export table.")
    TILLDataBase.Execute "CREATE TABLE temptbl2 (MailingAddress CHAR(50), MailingCity CHAR(25), MailingState CHAR(2), MailingZIP CHAR(10), ClientIndexedName CHAR(160), Location CHAR(50), LN1 CHAR(25), " & _
        "FN1 CHAR(25), MI1 CHAR(1), CO1 CHAR(50), FamilyFamiliarGreeting1 CHAR(30), FamilyMemberIsGuardian1 BIT, FamilyGoGreen1 BIT, EmailAddress1 CHAR(50), LN2 CHAR(25), FN2 CHAR(25), MI2 CHAR(1), " & _
        "CO2 CHAR(50), FamilyFamiliarGreeting2 CHAR(30), FamilyMemberIsGuardian2 BIT, FamilyGoGreen2 BIT, EmailAddress2 CHAR(50), LN3 CHAR(25), FN3 CHAR(25), MI3 CHAR(1), CO3 CHAR(50), FamilyFamiliarGreeting3 CHAR(30), FamilyMemberIsGuardian3 BIT, FamilyGoGreen3 BIT, EmailAddress3 CHAR(50), LN4 CHAR(25), FN4 CHAR(25), MI4 CHAR(1), CO4 CHAR(50), FamilyFamiliarGreeting4 CHAR(30), FamilyMemberIsGuardian4 BIT, FamilyGoGreen4 BIT, EmailAddress4 CHAR(50) );", dbSeeChanges: Call BriefDelay
    TILLDataBase.Execute "ALTER TABLE temptbl2 ADD CONSTRAINT temptbl2constraint PRIMARY KEY (MailingAddress, MailingCity, MailingState, MailingZIP, ClientIndexedName);", dbSeeChanges: Call BriefDelay
    TILLDataBase.Execute "INSERT INTO temptbl2 (MailingAddress, MailingCity, MailingState, MailingZIP, ClientIndexedName, Location) " & _
        "SELECT MailingAddress, MailingCity, MailingState, MailingZIP, ClientIndexedName, Location FROM temptbl1;", dbSeeChanges: Call BriefDelay
    Call ParseFamilyMembersAndThenExport(0, "temptbl2", "ResProgramNewsletters-AllLocations")
    SysCmdResult = SysCmd(5)
    Exit Sub
ShowMeError:
    MsgBox "Error # " & Str(Err.Number) & " was generated by " & Me.Name & Chr(13) & Err.Description, vbOKOnly, "Error", Err.HelpFile, Err.HelpContext
End Sub

Private Sub AutismSupport_Click()
On Error GoTo ShowMeError
    Call DropTempTables
    SysCmdResult = SysCmd(4, "Collecting autism client information.")
    TILLDataBase.Execute "SELECT tblPeople.IndexedName AS ClientIndexedName, tblPeople.LastName AS ClientLastName, tblPeople.FirstName AS ClientFirstName, " & _
        "tblPeople.MiddleInitial AS ClientMiddleInitial, tblPeopleFamily.IndexedName AS FamilyIndexedName, tblPeopleFamily.Relationship, tblPeopleFamily.Guardian, " & _
        "tblPeopleFamily.PrimaryContact, tblPeopleFamily.Surrogate, tblPeopleClientsAutismServices.Age " & _
        "INTO temptbl0 " & _
        "FROM (tblPeople LEFT JOIN tblPeopleFamily ON tblPeople.IndexedName = tblPeopleFamily.ClientIndexedName) " & _
        "LEFT JOIN tblPeopleClientsAutismServices ON tblPeople.IndexedName = tblPeopleClientsAutismServices.IndexedName " & _
        "WHERE tblPeopleFamily.Inactive=False " & _
        "AND tblPeople.IsClientAutism=True " & _
        "AND tblPeople.IsDeceased=False " & _
        "AND tblPeople.NoMailings=False " & _
        "AND tblPeopleClientsAutismServices.Inactive=False;", dbSeeChanges: Call BriefDelay
    
    SysCmdResult = SysCmd(4, "Collecting autism client family information.")
    TILLDataBase.Execute "SELECT ClientIndexedName, ClientLastName, ClientFirstName, ClientMiddleInitial, Age, FamilyIndexedName, Relationship, " & _
        "tblPeople.LastName AS LN, tblPeople.FirstName AS FN, tblPeople.MiddleInitial AS MI, " & _
        "tblPeople.CompanyOrganization AS CO, tblPeople.FamiliarGreeting AS FamilyFamiliarGreeting, " & _
        "tblPeople.MailingAddress AS MailingAddress, tblPeople.MailingCity AS MailingCity, tblPeople.MailingState AS MailingState, " & _
        "tblPeople.MailingZIP AS MailingZIP, tblPeople.GoGreen, tblPeople.EmailAddress INTO temptbl1 " & _
        "FROM ((temptbl0 INNER JOIN tblPeopleClientsVendors ON temptbl0.ClientIndexedName = tblPeopleClientsVendors.IndexedName) " & _
        "INNER JOIN tblPeople ON temptbl0.FamilyIndexedName = tblPeople.IndexedName) " & _
        "WHERE tblPeople.IsDeceased=FALSE " & _
        "AND (Guardian=True OR PrimaryContact=True OR Surrogate=True);", dbSeeChanges: Call BriefDelay
    
    SysCmdResult = SysCmd(4, "Building export table.")
    TILLDataBase.Execute "CREATE TABLE temptbl2 (MailingAddress CHAR(50), MailingCity CHAR(25), MailingState CHAR(2), MailingZIP CHAR(10), ClientIndexedName CHAR(160), Age INT, LN1 CHAR(25)," & _
        "FN1 CHAR(25), MI1 CHAR(1), CO1 CHAR(50), FamilyFamiliarGreeting1 CHAR(30), FamilyMemberIsGuardian1 BIT, FamilyGoGreen1 BIT, EmailAddress1 CHAR(50), LN2 CHAR(25), FN2 CHAR(25), MI2 CHAR(1), " & _
        "CO2 CHAR(50), FamilyFamiliarGreeting2 CHAR(30), FamilyMemberIsGuardian2 BIT, FamilyGoGreen2 BIT, EmailAddress2 CHAR(50), LN3 CHAR(25), FN3 CHAR(25), MI3 CHAR(1), CO3 CHAR(50), FamilyFamiliarGreeting3 CHAR(30), FamilyMemberIsGuardian3 BIT, FamilyGoGreen3 BIT, EmailAddress3 CHAR(50), LN4 CHAR(25), FN4 CHAR(25), MI4 CHAR(1), CO4 CHAR(50), FamilyFamiliarGreeting4 CHAR(30), FamilyMemberIsGuardian4 BIT, FamilyGoGreen4 BIT, EmailAddress4 CHAR(50));", dbSeeChanges: Call BriefDelay
    
    TILLDataBase.Execute "ALTER TABLE temptbl2 ADD CONSTRAINT temptbl2constraint PRIMARY KEY (MailingAddress,MailingCity,MailingState,MailingZIP,ClientIndexedName);", dbSeeChanges: Call BriefDelay
    
    ProgressMessages = ProgressMessages & "Inserting autism client information into export table." & vbCrLf: ProgressMessages.Requery
    TILLDataBase.Execute "INSERT INTO temptbl2 (MailingAddress, MailingCity, MailingState, MailingZIP, ClientIndexedName, Age)" & _
        "SELECT MailingAddress, MailingCity, MailingState, MailingZIP, ClientIndexedName, Age FROM temptbl1;", dbSeeChanges: Call BriefDelay

    Call ParseFamilyMembersAndThenExport(0, "temptbl2", "AutismSupportMailing")
    SysCmdResult = SysCmd(5)
    Exit Sub
ShowMeError:
    MsgBox "Error # " & Str(Err.Number) & " was generated by " & Me.Name & Chr(13) & Err.Description, vbOKOnly, "Error", Err.HelpFile, Err.HelpContext
End Sub

Private Sub CLIENTMAILINGS_Click()
    SysCmdResult = SysCmd(4, "Creating export table.")
    Call BasicExport("SELECT tblPeople.LastName, tblPeople.FirstName, tblPeople.MiddleInitial, tblPeopleClientsDemographics.LegalStatus, " & _
        "tblPeople.PhysicalAddress, tblPeople.PhysicalCity, tblPeople.PhysicalState, tblPeople.PhysicalZIP, tblPeople.MailingAddress, tblPeople.MailingCity, " & _
        "tblPeople.MailingState, tblPeople.MailingZIP, tblPeopleClientsDemographics.CountyOfResidence, tblPeople.HomePhone, tblPeople.MobilePhone, tblPeopleClientsDemographics.DateOfBirth, tblPeopleClientsDemographics.Race, tblPeopleClientsDemographics.Gender, tblPeopleClientsDemographics.SocialSecurityNumber, tblPeopleClientsDemographics.RepPayeeIsTILL, " & _
        "IIf([tblPeopleClientsDayServices]![Inactive],'Inactive',IIf([tblPeople]![IsClientDay],'Yes','No')) AS ClientDay, IIf([tblPeople]![IsClientDay],[tblPeopleClientsDayServices]![CityTown] & '-' & [tblPeopleClientsDayServices]![LocationName]) AS DayLocation, IIf([tblPeopleClientsResidentialServices]![Inactive],'Inactive',IIf([tblPeople]![IsClientRes],'Yes','No')) AS ClientRes, IIf([tblPeople]![IsClientRes],[tblPeopleClientsResidentialServices]![CityTown] & '-' & [tblPeopleClientsResidentialServices]![Location]) AS ResLocation, " & _
        "IIf([tblPeopleClientsCLOServices]![Inactive],'Inactive',IIf([tblPeople]![IsCilentCLO],'Yes','No')) AS ClientCLO, IIf([tblPeople]![IsCilentCLO],[tblPeopleClientsCLOServices]![CityTown] & '-' & [tblPeopleClientsCLOServices]![Location]) AS CLOLocation, " & _
        "IIf([tblPeopleClientsVocationalServices]![Inactive],'Inactive',IIf([tblPeople]![IsClientVocat],'Yes','No')) AS ClientVocat, IIf([tblPeople]![IsClientVocat],[tblPeopleClientsVocationalServices]![CityTown] & '-' & [tblPeopleClientsVocationalServices]![Location]) AS VocLocation, " & _
        "IIf([tblPeopleClientsSharedLivingServices]![Inactive],'Inactive',IIf([tblPeople]![IsClientSharedLiving],'Yes','No')) AS ClientSL, IIf([tblPeopleClientsAFCServices]![Inactive],'Inactive',IIf([tblPeople]![IsClientAFC],'Yes','No')) AS ClientAFC, " & _
        "IIf([tblPeopleClientsIndividualSupportServices]![Inactive],'Inactive',IIf([tblPeople]![IsClientIndiv],'Yes','No')) AS ClientISS, IIf([tblPeopleClientsNHDay]![Inactive],'Inactive',IIf([tblPeople]![IsClientNHDay],'Yes','No')) AS ClientNHDay, " & _
        "IIf([tblPeopleClientsNHRes]![Inactive],'Inactive',IIf([tblPeople]![IsClientNHRes],'Yes','No')) AS ClientNHRes, IIf([tblPeopleClientsAutismServices]![Inactive],'Inactive',IIf([tblPeople]![IsClientAutism],'Yes','No')) AS ClientAutism, " & _
        "IIf([tblPeopleClientsPCAServices]![Inactive],'Inactive',IIf([tblPeople]![IsClientPCA],'Yes','No')) AS ClientPCA, IIf([tblPeopleClientsSpringboardServices]![Inactive],'Inactive',IIf([tblPeople]![IsClientSpring],'Yes','No')) AS ClientSpringboard, " & _
        "IIf([tblPeopleClientsTRASEServices]![Inactive],'Inactive',IIf([tblPeople]![IsClientTRASE],'Yes','No')) AS ClientTRASE, IIf([tblPeopleClientsCommunityConnectionsServices]![Inactive],'Inactive',IIf([tblPeople]![IsClientCommunityConnections],'Yes','No')) AS ClientCommunityConnections INTO temptbl " & _
        "FROM ((((((((((((((tblPeople LEFT JOIN tblPeopleClientsDemographics ON tblPeople.IndexedName = tblPeopleClientsDemographics.IndexedName) LEFT JOIN tblPeopleClientsDayServices ON tblPeopleClientsDemographics.IndexedName = tblPeopleClientsDayServices.IndexedName) LEFT JOIN tblPeopleClientsCLOServices ON tblPeopleClientsDemographics.IndexedName = tblPeopleClientsCLOServices.IndexedName) LEFT JOIN tblPeopleClientsResidentialServices ON tblPeopleClientsDemographics.IndexedName = tblPeopleClientsResidentialServices.IndexedName) LEFT JOIN tblPeopleClientsVocationalServices ON tblPeopleClientsDemographics.IndexedName = tblPeopleClientsVocationalServices.IndexedName) LEFT JOIN tblPeopleClientsSharedLivingServices ON tblPeopleClientsDemographics.IndexedName = tblPeopleClientsSharedLivingServices.IndexedName) LEFT JOIN tblPeopleClientsAFCServices ON tblPeopleClientsDemographics.IndexedName = tblPeopleClientsAFCServices.IndexedName) " & _
        "LEFT JOIN tblPeopleClientsIndividualSupportServices ON tblPeopleClientsDemographics.IndexedName = tblPeopleClientsIndividualSupportServices.IndexedName) LEFT JOIN tblPeopleClientsNHDay ON tblPeopleClientsDemographics.IndexedName = tblPeopleClientsNHDay.IndexedName) LEFT JOIN tblPeopleClientsNHRes ON tblPeopleClientsDemographics.IndexedName = tblPeopleClientsNHRes.IndexedName) LEFT JOIN tblPeopleClientsAutismServices ON tblPeopleClientsDemographics.IndexedName = tblPeopleClientsAutismServices.IndexedName) LEFT JOIN tblPeopleClientsPCAServices ON tblPeopleClientsDemographics.IndexedName = tblPeopleClientsPCAServices.IndexedName) LEFT JOIN tblPeopleClientsSpringboardServices ON tblPeopleClientsDemographics.IndexedName = tblPeopleClientsSpringboardServices.IndexedName) LEFT JOIN tblPeopleClientsTRASEServices ON tblPeopleClientsDemographics.IndexedName = tblPeopleClientsTRASEServices.IndexedName) " & _
        "LEFT JOIN tblPeopleClientsCommunityConnectionsServices ON tblPeopleClientsDemographics.IndexedName = tblPeopleClientsCommunityConnectionsServices.IndexedName " & _
        "WHERE tblPeople.NoMailings=False AND tblPeople.IsDeceased=False AND tblPeople.ClientCompletelyInactive=False AND " & _
        "((tblPeople.IsClientDay=True AND tblPeopleClientsDayServices.Inactive=False) OR (tblPeople.IsClientRes=True AND tblPeopleClientsResidentialServices.Inactive=False) OR " & _
         "(tblPeople.IsClientVocat=True AND tblPeopleClientsVocationalServices.Inactive=False) OR (tblPeople.IsClientSharedLiving=True AND tblPeopleClientsSharedLivingServices.Inactive=False) OR " & _
         "(tblPeople.IsClientAFC=True AND tblPeopleClientsAFCServices.Inactive=False) OR (tblPeople.IsCilentCLO=True AND tblPeopleClientsCLOServices.Inactive=False) OR " & _
         "(tblPeople.IsClientIndiv=True AND tblPeopleClientsIndividualSupportServices.Inactive=False) OR (tblPeople.IsClientNHDay=True AND tblPeopleClientsNHDay.Inactive=False) OR " & _
         "(tblPeople.IsClientNHRes=True AND tblPeopleClientsNHRes.Inactive=False) OR (tblPeople.IsClientAutism=True AND tblPeopleClientsAutismServices.Inactive=False) OR " & _
         "(tblPeople.IsClientPCA=True AND tblPeopleClientsPCAServices.Inactive=False) OR (tblPeople.IsClientSpring=True AND tblPeopleClientsSpringboardServices.Inactive=False) OR " & _
         "(tblPeople.IsClientTRASE=True AND tblPeopleClientsTRASEServices.Inactive=False) OR (tblPeople.IsClientCommunityConnections=True AND tblPeopleClientsCommunityConnectionsServices.Inactive=False));", "ClientMailing")
    SysCmdResult = SysCmd(5)
End Sub

Private Sub CLIENTFAMILIES_Click()
On Error GoTo ShowMeError
    Call DropTempTables
    SysCmdResult = SysCmd(4, "Gathering family members.")
    TILLDataBase.Execute "SELECT tblPeople.TILLEGram, tblPeople.LastName, tblPeople.FirstName, tblPeople.CompanyOrganization, tblPeople.MiddleInitial, tblPeople.FamiliarGreeting, tblPeople.MailingAddress, tblPeople.MailingCity, tblPeople.MailingState, tblPeople.MailingZIP, tblPeople.IsFamilyGuardian, tblPeople.IsDonor, tblPeople.IsInterestedParty, tblPeople.IsConsultant, tblPeople.InterestedPartyCategory, tblPeople.InterestedPartyInactive, tblPeople.GoGreen, tblPeople.EmailAddress, tblPeopleFamily.Relationship , tblPeopleFamily.ClientIndexedName, tblPeopleFamily.ClientLastName, tblPeopleFamily.ClientFirstName, tblPeopleFamily.ClientMiddleInitial, tblPeopleFamily.Guardian, tblPeopleFamily.PrimaryContact " & _
        "INTO temptbl0 " & _
        "FROM tblPeople LEFT JOIN tblPeopleFamily ON tblPeople.IndexedName = tblPeopleFamily.IndexedName " & _
        "WHERE tblPeopleFamily.Inactive = False AND tblPeople.IsFamilyGuardian = True AND tblPeople.IsDeceased = False AND tblPeople.NoMailings = False " & _
        "ORDER BY tblPeople.IsInterestedParty DESC , tblPeople.IsConsultant DESC , tblPeople.IndexedName;", dbSeeChanges: Call BriefDelay
    
    SysCmdResult = SysCmd(4, "Gathering clients.")
    DoCmd.SetWarnings False
    DoCmd.OpenQuery "qryClientFamilyMailing", acViewNormal
    Call BriefDelay(2)
    DoCmd.SetWarnings True
    
    SysCmdResult = SysCmd(4, "Creating export table.")
    TILLDataBase.Execute "CREATE TABLE temptbl2 (MailingAddress CHAR(50), MailingCity CHAR(25), MailingState CHAR(2), MailingZIP CHAR(10), ClientIndexedName CHAR(160), LN1 CHAR(25)," & _
                                       "FN1 CHAR(25), MI1 CHAR(1), CO1 CHAR(50), FamilyFamiliarGreeting1 CHAR(30), FamilyMemberIsGuardian1 BIT, FamilyGoGreen1 BIT, EmailAddress1 CHAR(50), LN2 CHAR(25), FN2 CHAR(25), MI2 CHAR(1), " & _
                                       "CO2 CHAR(50), FamilyFamiliarGreeting2 CHAR(30), FamilyMemberIsGuardian2 BIT, FamilyGoGreen2 BIT, EmailAddress2 CHAR(50), LN3 CHAR(25), FN3 CHAR(25), MI3 CHAR(1), CO3 CHAR(50), FamilyFamiliarGreeting3 CHAR(30), FamilyMemberIsGuardian3 BIT, FamilyGoGreen3 BIT, EmailAddress3 CHAR(50), LN4 CHAR(25), " & _
                                       "FN4 CHAR(25), MI4 CHAR(1), CO4 CHAR(50), FamilyFamiliarGreeting4 CHAR(30), FamilyMemberIsGuardian4 BIT, FamilyGoGreen4 BIT, EmailAddress4 CHAR(50), TILLEGram BIT, ClientLastName CHAR(25), ClientFirstName CHAR(25), ClientMiddleInitial CHAR(1), " & _
                                       "IsClientDay CHAR(8), IsClientVocat CHAR(8), DayProgram CHAR(60), IsClientRes CHAR(8), IsClientCLO CHAR(8), ResProgram CHAR(60), IsClientSharedLiving CHAR(8), IsClientAFC CHAR(8), IsClientNHDay CHAR(8), IsClientNHRes BIT, IsClientIndiv CHAR(8), IsClientAutism CHAR(8), IsClientPCA CHAR(8), " & _
                                       "IsClientSpring CHAR(8), IsClientTRASE CHAR(8), IsClientCommunityConnections CHAR(8), LivingWithParentOrGuardian BIT, LivingIndependently BIT);", dbSeeChanges: Call BriefDelay
    
    TILLDataBase.Execute "ALTER TABLE temptbl2 ADD CONSTRAINT temptbl2constraint PRIMARY KEY (MailingAddress,MailingCity,MailingState,MailingZIP,ClientIndexedName);", dbSeeChanges: Call BriefDelay
    SysCmdResult = SysCmd(4, "Pre-fill export table.")
    TILLDataBase.Execute "INSERT INTO temptbl2 (MailingAddress, MailingCity, MailingState, MailingZIP, ClientIndexedName, TILLEGram, " & _
        "ClientLastName, ClientFirstName, ClientMiddleInitial, IsClientDay, IsClientVocat, DayProgram, IsClientRes, ResProgram, IsClientSharedLiving, " & _
        "IsClientAFC, IsClientNHDay, IsClientNHRes, IsClientCLO, IsClientIndiv, IsClientAutism, IsClientPCA, IsClientSpring, IsClientTRASE, " & _
        "IsClientCommunityConnections, LivingWithParentOrGuardian, LivingIndependently ) " & _
        "SELECT MailingAddress, MailingCity, MailingState, MailingZIP, ClientIndexedName, TILLEGram, " & _
        "ClientLastName, ClientFirstName, ClientMiddleInitial, IsClientDay, IsClientVocat, DayVendorLocation, IsClientRes, ResVendorLocation, IsClientSharedLiving, " & _
        "IsClientAFC, IsClientNHDay, IsClientNHRes, IsClientCLO, IsClientIndiv, IsClientAutism , IsClientPCA, IsClientSpring, IsClientTRASE, " & _
        "IsClientCommunityConnections, LivingWithParentOrGuardian, LivingIndependently " & vbCrLf & _
        "FROM temptbl1 " & _
        "WHERE Left(MailingAddress,1) <> ' '", dbSeeChanges: Call BriefDelay(2)
    
    Call ParseFamilyMembersAndThenExport(1, "temptbl2", "ClientFamilyMailing")
    SysCmdResult = SysCmd(5)
    Exit Sub
ShowMeError:
    Call DropTempTables
    SysCmdResult = SysCmd(5)
    MsgBox "Error # " & Str(Err.Number) & " was generated by " & Me.Name & Chr(13) & Err.Description, vbOKOnly, "Error", Err.HelpFile, Err.HelpContext
End Sub

Private Sub ExpirationsByLocation_Click()
    SysCmdResult = SysCmd(4, "Creating export table.")
    Call BasicExport("SELECT tblLocations.Department, tblLocations.CityTown, tblLocations.LocationName, tblLocations.Cluster, " & _
        "IIf(IsNull(tblLocations.LastVehicleChecklistCompleted),Null,DateValue(tblLocations.LastVehicleChecklistCompleted)) AS LastVehicleChecklistCompleted, " & _
        "IIf(IsNull(tblLocations.MostRecentAsleepFireDrill),Null,DateValue(tblLocations.MostRecentAsleepFireDrill)) AS MostRecentAsleepFireDrill, " & _
        "IIf(IsNull(tblLocations.NextRecentAsleepFireDrill),Null,DateValue(tblLocations.NextRecentAsleepFireDrill)) AS NextRecentAsleepFireDrill, " & _
        "IIf(IsNull(tblLocations.DAYStaffTrainedInPrivacyBefore),Null,DateValue(tblLocations.DAYStaffTrainedInPrivacyBefore)) AS DAYStaffTrainedInPrivacyBefore, " & _
        "IIf(IsNull(tblLocations.DAYAllPlansReviewedByStaffBefore),Null,DateValue(tblLocations.DAYAllPlansReviewedByStaffBefore)) AS DAYAllPlansReviewedByStaffBefore, " & _
        "IIf(IsNull(tblLocations.DAYQtrlySafetyChecklistDueBy),Null,DateValue(tblLocations.DAYQtrlySafetyChecklistDueBy)) AS DAYQtrlySafetyChecklistDueBy, " & _
        "IIf(IsNull(tblLocations.HouseSafetyPlanExpires),Null,DateValue(tblLocations.HouseSafetyPlanExpires)) AS HouseSafetyPlanExpires, " & _
        "IIf(IsNull(tblLocations.HousePlansReviewedByStaffBefore),Null,DateValue(tblLocations.HousePlansReviewedByStaffBefore)) AS HousePlansReviewedByStaffBefore, " & _
        "IIf(IsNull(tblLocations.MAPChecklistCompleted),Null,DateValue(tblLocations.MAPChecklistCompleted)) AS MAPChecklistCompleted, tblLocations.HumanRightsOfficer, " & _
        "IIf(IsNull(tblLocations.HROTrainsStaffBefore),Null,DateValue(tblLocations.HROTrainsStaffBefore)) AS HROTrainsStaffBefore, " & _
        "IIf(IsNull(tblLocations.HROTrainsIndividualsBefore),Null,DateValue(tblLocations.HROTrainsIndividualsBefore)) AS HROTrainsIndividualsBefore, tblLocations.FireSafetyOfficer, " & _
        "IIf(IsNull(tblLocations.FSOTrainsStaffBefore),Null,DateValue(tblLocations.FSOTrainsStaffBefore)) AS FSOTrainsStaffBefore, " & _
        "IIf(IsNull(tblLocations.FSOTrainsIndividualsBefore),Null,DateValue(tblLocations.FSOTrainsIndividualsBefore)) AS FSOTrainsIndividualsBefore INTO temptbl FROM tblLocations " & _
        "WHERE tblLocations.GPName Is Not Null AND (tblLocations.Department = 'Residential Services' OR tblLocations.Department = 'Individualized Support Options' OR tblLocations.Department = 'Day Services') " & _
        "ORDER BY tblLocations.Department, tblLocations.CityTown, tblLocations.LocationName;", "LocationExpirations")
    SysCmdResult = SysCmd(5)
End Sub

Private Sub GoDonors_Click()
    Dim ExportFileName As String
    ' Export the Donors table.
    SysCmdResult = SysCmd(4, "Building export table.")
    ExportFileName = Application.CurrentProject.Path & "\" & "TILLDB-Export-Donations-" & Format(Date, "yyyymmdd") & ".xls"
    If IsFileOpen(ExportFileName) Then
        If MsgBox(ExportFileName & " is already open.  Please close it and click OK to continue or Cancel to abort.", vbOKCancel, "ERROR!") = vbCancel Then
            MsgBox "Export aborted.", vbOKOnly, "Aborted"
            Exit Sub
        End If
    End If
    If Dir(ExportFileName) <> "" Then Kill ExportFileName
    DoCmd.TransferSpreadsheet acExport, acSpreadsheetTypeExcel9, "qryDonationsForExport", ExportFileName
'   CommandLine = LocateExecutable(ExportFileName) & " """ & ExportFileName & """"
    MsgBox "The requested Donor information has been exported to " & ExportFileName & "." & vbCrLf & vbCrLf & "This export may contain information that is protected under HIPAA and other privacy laws.  This export must be securely stored at all times and must be deleted when no longer being used.", _
        vbOKOnly, "Export Complete"
'   RetValue = Shell(CommandLine, 1)
End Sub

Private Sub GovAccounts_Click()
On Error GoTo ShowMeError
    Dim ExportFileName As String
    
    ExportFileName = Application.CurrentProject.Path & "\" & "TILLDB-Export-ResGovernmentAccounts-" & Format(Date, "yyyymmdd") & ".xls"
    SysCmdResult = SysCmd(4, "Building export table.")
    If IsFileOpen(ExportFileName) Then
        If MsgBox(ExportFileName & " is already open.  Please close it and click OK to continue or Cancel to abort.", vbOKCancel, "ERROR!") = vbCancel Then
            MsgBox "Export aborted.", vbOKOnly, "Aborted"
            Exit Sub
        End If
    End If
    If Dir(ExportFileName) <> "" Then Kill ExportFileName
    DoCmd.TransferSpreadsheet acExport, acSpreadsheetTypeExcel9, "qryClientGovernmentAccounts", ExportFileName
'   CommandLine = LocateExecutable(ExportFileName) & " """ & ExportFileName & """"
    MsgBox "The requested information has been exported to " & ExportFileName & ".  Launching Excel." & vbCrLf & vbCrLf & "This export may contain information that is protected under HIPAA and other privacy laws.  This export must be securely stored at all times and must be deleted when no longer being used.", _
        vbOKOnly, "Export Complete"
'   RetValue = Shell(CommandLine, 1)
    SysCmdResult = SysCmd(5)
    Exit Sub
ShowMeError:
    MsgBox "Error # " & Str(Err.Number) & " was generated by " & Me.Name & Chr(13) & Err.Description, vbOKOnly, "Error", Err.HelpFile, Err.HelpContext
End Sub

Private Sub Inactive_Deceased_Click()
On Error GoTo ShowMeError
    Call DropTempTables
    ' Deceased.
    SysCmdResult = SysCmd(4, "Identify deceased people.")
    TILLDataBase.Execute "SELECT tblPeople.IndexedName AS IndexedName, tblPeople.LastName AS LastName, tblPeople.FirstName AS FirstName, tblPeople.MiddleInitial AS MiddleInitial, 'Deceased' AS Status, IIf(IsNull([tblPeople]![DeceasedDate]),Null,DateValue([tblPeople]![DeceasedDate])) AS StatusDate, " & _
        "IIf([tblPeople]![IsClient],'Client',IIf([tblPeople]![IsFamilyGuardian],'Family',IIf([tblPeople]![IsDonor],'Donor',IIf([tblPeople]![IsInterestedParty],'Interested Party',IIf([tblPeople]![IsStaff],'Staff',IIf([tblPeople]![IsConsultant],'Consultant','Unspecified')))))) AS Category " & vbCrLf & _
        "INTO temptbl0 " & vbCrLf & "FROM tblPeople " & vbCrLf & "WHERE tblPeople.IsDeceased=True;", dbSeeChanges: Call BriefDelay
    ' All Inactive Clients.
    SysCmdResult = SysCmd(4, "Identify inactive clients.")
    TILLDataBase.Execute "INSERT INTO temptbl0 ( IndexedName, LastName, FirstName, MiddleInitial, Status, StatusDate, Category ) " & vbCrLf & _
        "SELECT tblPeople.IndexedName, tblPeople.LastName, tblPeople.FirstName, tblPeople.MiddleInitial, 'All Inactive' AS Status, IIf(IsNull([tblPeople]![ClientCompletelyInactiveDate]),Null,DateValue([tblPeople]![ClientCompletelyInactiveDate])) AS StatusDate, 'Client' AS Category " & vbCrLf & _
        "FROM tblPeople " & vbCrLf & "WHERE tblPeople.ClientCompletelyInactive=True;", dbSeeChanges: Call BriefDelay
    ' Inactive from Programs.
    SysCmdResult = SysCmd(4, "Identify inactive Autism clients.")
    TILLDataBase.Execute "INSERT INTO temptbl0 ( IndexedName, LastName, FirstName, MiddleInitial, Status, StatusDate, Category ) " & vbCrLf & _
        "SELECT tblPeople.IndexedName, tblPeople.LastName, tblPeople.FirstName, tblPeople.MiddleInitial, 'Inactive' AS Status, IIf(IsNull([tblPeopleClientsAutismServices]![DateInactive]),Null,DateValue([tblPeopleClientsAutismServices]![DateInactive])) AS StatusDate, 'Autism' AS Category " & vbCrLf & _
        "FROM tblPeople INNER JOIN tblPeopleClientsAutismServices ON tblPeople.IndexedName = tblPeopleClientsAutismServices.IndexedName " & vbCrLf & _
        "WHERE tblPeople.ClientCompletelyInactive=False AND tblPeople.IsClientAutism=True AND tblPeopleClientsAutismServices.Inactive=True;", dbSeeChanges: Call BriefDelay
    SysCmdResult = SysCmd(4, "Identify inactive CLO clients.")
    TILLDataBase.Execute "INSERT INTO temptbl0 ( IndexedName, LastName, FirstName, MiddleInitial, Status, StatusDate, Category ) " & vbCrLf & _
        "SELECT tblPeople.IndexedName, tblPeople.LastName, tblPeople.FirstName, tblPeople.MiddleInitial, 'Inactive' AS Status, IIf(IsNull([tblPeopleClientsCLOServices]![DateInactive]),Null,DateValue([tblPeopleClientsCLOServices]![DateInactive])) AS StatusDate, 'CLO' AS Category " & vbCrLf & _
        "FROM tblPeople INNER JOIN tblPeopleClientsCLOServices ON tblPeople.IndexedName = tblPeopleClientsCLOServices.IndexedName " & vbCrLf & _
        "WHERE tblPeople.ClientCompletelyInactive=False AND tblPeople.IsCilentCLO=True AND tblPeopleClientsCLOServices.Inactive=True;", dbSeeChanges: Call BriefDelay
    SysCmdResult = SysCmd(4, "Identify inactive Community Connections clients.")
    TILLDataBase.Execute "INSERT INTO temptbl0 ( IndexedName, LastName, FirstName, MiddleInitial, Status, StatusDate, Category ) " & vbCrLf & _
        "SELECT tblPeople.IndexedName, tblPeople.LastName, tblPeople.FirstName, tblPeople.MiddleInitial, 'Inactive' AS Status, IIf(IsNull([tblPeopleClientsCommunityConnectionsServices]![DateInactive]),Null,DateValue([tblPeopleClientsCommunityConnectionsServices]![DateInactive])) AS StatusDate, 'ComCon' AS Category " & vbCrLf & _
        "FROM tblPeople INNER JOIN tblPeopleClientsCommunityConnectionsServices ON tblPeople.IndexedName = tblPeopleClientsCommunityConnectionsServices.IndexedName " & vbCrLf & _
        "WHERE tblPeople.ClientCompletelyInactive=False AND tblPeople.IsClientCommunityConnections=True AND tblPeopleClientsCommunityConnectionsServices.Inactive=True;", dbSeeChanges: Call BriefDelay
    SysCmdResult = SysCmd(4, "Identify inactive Day clients.")
    TILLDataBase.Execute "INSERT INTO temptbl0 ( IndexedName, LastName, FirstName, MiddleInitial, Status, StatusDate, Category ) " & vbCrLf & _
        "SELECT tblPeople.IndexedName, tblPeople.LastName, tblPeople.FirstName, tblPeople.MiddleInitial, 'Inactive' AS Status, IIf(IsNull([tblPeopleClientsDayServices]![DateInactive]),Null,DateValue([tblPeopleClientsDayServices]![DateInactive])) AS StatusDate, 'Day' AS Category " & vbCrLf & _
        "FROM tblPeople INNER JOIN tblPeopleClientsDayServices ON tblPeople.IndexedName = tblPeopleClientsDayServices.IndexedName " & vbCrLf & _
        "WHERE tblPeople.ClientCompletelyInactive=False AND tblPeople.IsClientDay=True AND tblPeopleClientsDayServices.Inactive=True;", dbSeeChanges: Call BriefDelay
    SysCmdResult = SysCmd(4, "Identify inactive Individual Support clients.")
    TILLDataBase.Execute "INSERT INTO temptbl0 ( IndexedName, LastName, FirstName, MiddleInitial, Status, StatusDate, Category ) " & vbCrLf & _
        "SELECT tblPeople.IndexedName, tblPeople.LastName, tblPeople.FirstName, tblPeople.MiddleInitial, 'Inactive' AS Status, IIf(IsNull([tblPeopleClientsIndividualSupportServices]![DateInactive]),Null,DateValue([tblPeopleClientsIndividualSupportServices]![DateInactive])) AS StatusDate, 'ISS' AS Category " & vbCrLf & _
        "FROM tblPeople INNER JOIN tblPeopleClientsIndividualSupportServices ON tblPeople.IndexedName = tblPeopleClientsIndividualSupportServices.IndexedName " & vbCrLf & _
        "WHERE tblPeople.ClientCompletelyInactive=False AND tblPeople.IsClientIndiv=True AND tblPeopleClientsIndividualSupportServices.Inactive=True;", dbSeeChanges: Call BriefDelay
    SysCmdResult = SysCmd(4, "Identify inactive NH Day clients.")
    TILLDataBase.Execute "INSERT INTO temptbl0 ( IndexedName, LastName, FirstName, MiddleInitial, Status, StatusDate, Category ) " & vbCrLf & _
        "SELECT tblPeople.IndexedName, tblPeople.LastName, tblPeople.FirstName, tblPeople.MiddleInitial, 'Inactive' AS Status, IIf(IsNull([tblPeopleClientsNHDay]![DateInactive]),Null,DateValue([tblPeopleClientsNHDay]![DateInactive])) AS StatusDate, 'NHDay' AS Category " & vbCrLf & _
        "FROM tblPeople INNER JOIN tblPeopleClientsNHDay ON tblPeople.IndexedName = tblPeopleClientsNHDay.IndexedName " & vbCrLf & _
        "WHERE tblPeople.ClientCompletelyInactive=False AND tblPeople.IsClientNHDay=True AND tblPeopleClientsNHDay.Inactive=True;", dbSeeChanges: Call BriefDelay
    SysCmdResult = SysCmd(4, "Identify inactive NH Residential clients.")
    TILLDataBase.Execute "INSERT INTO temptbl0 ( IndexedName, LastName, FirstName, MiddleInitial, Status, StatusDate, Category ) " & vbCrLf & _
        "SELECT tblPeople.IndexedName, tblPeople.LastName, tblPeople.FirstName, tblPeople.MiddleInitial, 'Inactive' AS Status, IIf(IsNull([tblPeopleClientsNHRes]![DateInactive]),Null,DateValue([tblPeopleClientsNHRes]![DateInactive])) AS StatusDate, 'NHRes' AS Category " & vbCrLf & _
        "FROM tblPeople INNER JOIN tblPeopleClientsNHRes ON tblPeople.IndexedName = tblPeopleClientsNHRes.IndexedName " & vbCrLf & _
        "WHERE tblPeople.ClientCompletelyInactive=False AND tblPeople.IsClientNHRes=True AND tblPeopleClientsNHRes.Inactive=True;", dbSeeChanges: Call BriefDelay
    SysCmdResult = SysCmd(4, "Identify inactive PCA clients.")
    TILLDataBase.Execute "INSERT INTO temptbl0 ( IndexedName, LastName, FirstName, MiddleInitial, Status, StatusDate, Category ) " & vbCrLf & _
        "SELECT tblPeople.IndexedName, tblPeople.LastName, tblPeople.FirstName, tblPeople.MiddleInitial, 'Inactive' AS Status, IIf(IsNull([tblPeopleClientsPCAServices]![DateInactive]),Null,DateValue([tblPeopleClientsPCAServices]![DateInactive])) AS StatusDate, 'PCAServices' AS Category " & vbCrLf & _
        "FROM tblPeople INNER JOIN tblPeopleClientsPCAServices ON tblPeople.IndexedName = tblPeopleClientsPCAServices.IndexedName " & vbCrLf & _
        "WHERE tblPeople.ClientCompletelyInactive=False AND tblPeople.IsClientPCA=True AND tblPeopleClientsPCAServices.Inactive=True;", dbSeeChanges: Call BriefDelay
    SysCmdResult = SysCmd(4, "Identify inactive Residential clients.")
    TILLDataBase.Execute "INSERT INTO temptbl0 ( IndexedName, LastName, FirstName, MiddleInitial, Status, StatusDate, Category ) " & vbCrLf & _
        "SELECT tblPeople.IndexedName, tblPeople.LastName, tblPeople.FirstName, tblPeople.MiddleInitial, 'Inactive' AS Status, IIf(IsNull([tblPeopleClientsResidentialServices]![DateInactive]),Null,DateValue([tblPeopleClientsResidentialServices]![DateInactive])) AS StatusDate, 'ResidentialServices' AS Category " & vbCrLf & _
        "FROM tblPeople INNER JOIN tblPeopleClientsResidentialServices ON tblPeople.IndexedName = tblPeopleClientsResidentialServices.IndexedName " & vbCrLf & _
        "WHERE tblPeople.ClientCompletelyInactive=False AND tblPeople.IsClientRes=True AND tblPeopleClientsResidentialServices.Inactive=True;", dbSeeChanges: Call BriefDelay
    SysCmdResult = SysCmd(4, "Identify inactive Shared Living clients.")
    TILLDataBase.Execute "INSERT INTO temptbl0 ( IndexedName, LastName, FirstName, MiddleInitial, Status, StatusDate, Category ) " & vbCrLf & _
        "SELECT tblPeople.IndexedName, tblPeople.LastName, tblPeople.FirstName, tblPeople.MiddleInitial, 'Inactive' AS Status, IIf(IsNull([tblPeopleClientsSharedLivingServices]![DateInactive]),Null,DateValue([tblPeopleClientsSharedLivingServices]![DateInactive])) AS StatusDate, 'SharedLivingServices' AS Category " & vbCrLf & _
        "FROM tblPeople INNER JOIN tblPeopleClientsSharedLivingServices ON tblPeople.IndexedName = tblPeopleClientsSharedLivingServices.IndexedName " & vbCrLf & _
        "WHERE tblPeople.ClientCompletelyInactive=False AND tblPeople.IsClientSharedLiving=True AND tblPeopleClientsSharedLivingServices.Inactive=True;", dbSeeChanges: Call BriefDelay
    SysCmdResult = SysCmd(4, "Identify inactive AFC clients.")
    TILLDataBase.Execute "INSERT INTO temptbl0 ( IndexedName, LastName, FirstName, MiddleInitial, Status, StatusDate, Category ) " & vbCrLf & _
        "SELECT tblPeople.IndexedName, tblPeople.LastName, tblPeople.FirstName, tblPeople.MiddleInitial, 'Inactive' AS Status, IIf(IsNull([tblPeopleClientsAFCServices]![DateInactive]),Null,DateValue([tblPeopleClientsAFCServices]![DateInactive])) AS StatusDate, 'AFCServices' AS Category " & vbCrLf & _
        "FROM tblPeople INNER JOIN tblPeopleClientsAFCServices ON tblPeople.IndexedName = tblPeopleClientsAFCServices.IndexedName " & vbCrLf & _
        "WHERE tblPeople.ClientCompletelyInactive=False AND tblPeople.IsClientAFC=True AND tblPeopleClientsAFCServices.Inactive=True;", dbSeeChanges: Call BriefDelay
    SysCmdResult = SysCmd(4, "Identify inactive Springboard clients.")
    TILLDataBase.Execute "INSERT INTO temptbl0 ( IndexedName, LastName, FirstName, MiddleInitial, Status, StatusDate, Category ) " & vbCrLf & _
        "SELECT tblPeople.IndexedName, tblPeople.LastName, tblPeople.FirstName, tblPeople.MiddleInitial, 'Inactive' AS Status, IIf(IsNull([tblPeopleClientsSpringboardServices]![DateInactive]),Null,DateValue([tblPeopleClientsSpringboardServices]![DateInactive])) AS StatusDate, 'SpringboardServices' AS Category " & vbCrLf & _
        "FROM tblPeople INNER JOIN tblPeopleClientsSpringboardServices ON tblPeople.IndexedName = tblPeopleClientsSpringboardServices.IndexedName " & vbCrLf & _
        "WHERE tblPeople.ClientCompletelyInactive=False AND tblPeople.IsClientSpring=True AND tblPeopleClientsSpringboardServices.Inactive=True;", dbSeeChanges: Call BriefDelay
    SysCmdResult = SysCmd(4, "Identify inactive TRASE clients.")
    TILLDataBase.Execute "INSERT INTO temptbl0 ( IndexedName, LastName, FirstName, MiddleInitial, Status, StatusDate, Category ) " & vbCrLf & _
        "SELECT tblPeople.IndexedName, tblPeople.LastName, tblPeople.FirstName, tblPeople.MiddleInitial, 'Inactive' AS Status, IIf(IsNull([tblPeopleClientsTRASEServices]![DateInactive]),Null,DateValue([tblPeopleClientsTRASEServices]![DateInactive])) AS StatusDate, 'TRASEServices' AS Category " & vbCrLf & _
        "FROM tblPeople INNER JOIN tblPeopleClientsTRASEServices ON tblPeople.IndexedName = tblPeopleClientsTRASEServices.IndexedName " & vbCrLf & _
        "WHERE tblPeople.ClientCompletelyInactive=False AND tblPeople.IsClientTRASE=True AND tblPeopleClientsTRASEServices.Inactive=True;", dbSeeChanges: Call BriefDelay
    SysCmdResult = SysCmd(4, "Identify inactive Transportation Services clients.")
    TILLDataBase.Execute "INSERT INTO temptbl0 ( IndexedName, LastName, FirstName, MiddleInitial, Status, StatusDate, Category ) " & vbCrLf & _
        "SELECT tblPeople.IndexedName, tblPeople.LastName, tblPeople.FirstName, tblPeople.MiddleInitial, 'Inactive' AS Status, IIf(IsNull([tblPeopleClientsTransportationServices]![DateInactive]),Null,DateValue([tblPeopleClientsTransportationServices]![DateInactive])) AS StatusDate, 'TransportationServices' AS Category " & vbCrLf & _
        "FROM tblPeople INNER JOIN tblPeopleClientsTransportationServices ON tblPeople.IndexedName = tblPeopleClientsTransportationServices.IndexedName " & vbCrLf & _
        "WHERE tblPeople.ClientCompletelyInactive=False AND tblPeople.IsClientTrans=True AND tblPeopleClientsTransportationServices.Inactive=True;", dbSeeChanges: Call BriefDelay
    SysCmdResult = SysCmd(4, "Identify inactive Vocational Services clients.")
    TILLDataBase.Execute "INSERT INTO temptbl0 ( IndexedName, LastName, FirstName, MiddleInitial, Status, StatusDate, Category ) " & vbCrLf & _
        "SELECT tblPeople.IndexedName, tblPeople.LastName, tblPeople.FirstName, tblPeople.MiddleInitial, 'Inactive' AS Status, IIf(IsNull([tblPeopleClientsVocationalServices]![DateInactive]),Null,DateValue([tblPeopleClientsVocationalServices]![DateInactive])) AS StatusDate, 'VocationalServices' AS Category " & vbCrLf & _
        "FROM tblPeople INNER JOIN tblPeopleClientsVocationalServices ON tblPeople.IndexedName = tblPeopleClientsVocationalServices.IndexedName " & vbCrLf & _
        "WHERE tblPeople.ClientCompletelyInactive=False AND tblPeople.IsClientVocat=True AND tblPeopleClientsVocationalServices.Inactive=True;", dbSeeChanges: Call BriefDelay
    ProgressMessages = ProgressMessages & "Creating export table." & vbCrLf: ProgressMessages.Requery
    Call ParseFamilyMembersAndThenExport(9, "temptbl0", "InactiveAndDeceased")
    SysCmdResult = SysCmd(5)
    Exit Sub
ShowMeError:
    MsgBox "Error # " & Str(Err.Number) & " was generated by " & Me.Name & Chr(13) & Err.Description, vbOKOnly, "Error", Err.HelpFile, Err.HelpContext
End Sub

Private Sub InternalEval_Click()
On Error GoTo ShowMeError
    Call DropTempTables
    SysCmdResult = SysCmd(4, "Collect relevant clients.")
    TILLDataBase.Execute "SELECT tblPeople.TILLEGram, tblPeople.LastName, tblPeople.FirstName, tblPeople.CompanyOrganization, tblPeople.MiddleInitial, tblPeople.FamiliarGreeting, tblPeople.MailingAddress, tblPeople.MailingCity, tblPeople.MailingState, tblPeople.MailingZIP, tblPeople.IsFamilyGuardian, tblPeople.IsDonor, tblPeople.IsInterestedParty, tblPeople.IsConsultant, tblPeople.InterestedPartyCategory, tblPeople.InterestedPartyInactive, tblPeople.GoGreen, tblPeople.EmailAddress, tblPeopleFamily.Relationship , tblPeopleFamily.ClientIndexedName, tblPeopleFamily.ClientLastName, tblPeopleFamily.ClientFirstName, tblPeopleFamily.ClientMiddleInitial, tblPeopleFamily.Guardian, tblPeopleFamily.PrimaryContact INTO temptbl0" & vbCrLf & _
        "FROM tblPeople LEFT JOIN tblPeopleFamily ON tblPeople.IndexedName = tblPeopleFamily.IndexedName " & vbCrLf & _
        "WHERE tblPeopleFamily.Inactive = False AND tblPeople.IsFamilyGuardian = True AND tblPeople.IsDeceased = False AND tblPeople.NoMailings = False " & vbCrLf & _
        "ORDER BY tblPeople.IsInterestedParty DESC , tblPeople.IsConsultant DESC , tblPeople.IndexedName;", dbSeeChanges: Call BriefDelay
    TILLDataBase.Execute "SELECT temptbl0.LastName AS LN, temptbl0.FirstName as FN, temptbl0.MiddleInitial AS MI, temptbl0.CompanyOrganization AS CO, temptbl0.FamiliarGreeting AS FamilyFamiliarGreeting, temptbl0.MailingAddress AS MailingAddress, temptbl0.MailingCity AS MailingCity, temptbl0.MailingState AS MailingState, temptbl0.MailingZIP AS MailingZIP, temptbl0.IsFamilyGuardian, temptbl0.IsDonor, temptbl0.IsInterestedParty, temptbl0.IsConsultant, temptbl0.InterestedPartyCategory, temptbl0.InterestedPartyInactive, temptbl0.TILLEGram, temptbl0.GoGreen, temptbl0.EmailAddress, temptbl0.Relationship, temptbl0.Guardian, temptbl0.PrimaryContact, temptbl0.ClientIndexedName, temptbl0.ClientLastName, temptbl0.ClientFirstName, temptbl0.ClientMiddleInitial, tblPeople.IsClientDay, tblPeople.IsClientVocat, tblPeopleClientsVendors.DayVendorLocation, " & _
        "tblPeople.IsClientRes, tblPeople.IsCilentCLO AS IsClientCLO, tblPeopleClientsVendors.ResVendorLocation, tblPeople.IsClientTrans, tblPeople.IsClientSharedLiving, tblPeople.IsClientAFC, " & _
        "tblPeople.IsClientNHDay, tblPeople.IsClientNHRes, tblPeople.IsClientIndiv, tblPeople.IsClientAutism, tblPeople.IsClientPCA, tblPeople.IsClientSpring, " & _
        "tblPeople.IsClientTRASE, tblPeople.IsClientCommunityConnections, tblPeopleClientsVendors.LivingWithParentOrGuardian, tblPeopleClientsVendors.LivingIndependently " & vbCrLf & _
        "INTO temptbl1 FROM (((((((((((((((temptbl0 INNER JOIN tblPeople ON temptbl0.ClientIndexedName = tblPeople.IndexedName) INNER JOIN tblPeopleClientsVendors ON tblPeople.IndexedName = tblPeopleClientsVendors.IndexedName) LEFT JOIN tblPeopleClientsAutismServices ON temptbl0.ClientIndexedName = tblPeopleClientsAutismServices.IndexedName) LEFT JOIN tblPeopleClientsCLOServices ON temptbl0.ClientIndexedName = tblPeopleClientsCLOServices.IndexedName) LEFT JOIN tblPeopleClientsCommunityConnectionsServices ON temptbl0.ClientIndexedName = tblPeopleClientsCommunityConnectionsServices.IndexedName) LEFT JOIN tblPeopleClientsDayServices ON temptbl0.ClientIndexedName = tblPeopleClientsDayServices.IndexedName)  " & _
        "LEFT JOIN tblPeopleClientsIndividualSupportServices ON temptbl0.ClientIndexedName = tblPeopleClientsIndividualSupportServices.IndexedName) LEFT JOIN tblPeopleClientsPCAServices ON temptbl0.ClientIndexedName = tblPeopleClientsPCAServices.IndexedName) LEFT JOIN tblPeopleClientsResidentialServices ON temptbl0.ClientIndexedName = tblPeopleClientsResidentialServices.IndexedName) LEFT JOIN tblPeopleClientsSharedLivingServices ON temptbl0.ClientIndexedName = tblPeopleClientsSharedLivingServices.IndexedName) LEFT JOIN tblPeopleClientsAFCServices ON temptbl0.ClientIndexedName = tblPeopleClientsAFCServices.IndexedName) LEFT JOIN tblPeopleClientsSpringboardServices ON temptbl0.ClientIndexedName = tblPeopleClientsSpringboardServices.IndexedName) " & _
        "LEFT JOIN tblPeopleClientsTRASEServices ON temptbl0.ClientIndexedName = tblPeopleClientsTRASEServices.IndexedName) LEFT JOIN tblPeopleClientsVocationalServices ON temptbl0.ClientIndexedName = tblPeopleClientsVocationalServices.IndexedName) LEFT JOIN tblPeopleClientsNHRes ON temptbl0.ClientIndexedName = tblPeopleClientsNHRes.IndexedName) LEFT JOIN tblPeopleClientsNHDay ON temptbl0.ClientIndexedName = tblPeopleClientsNHDay.IndexedName " & vbCrLf & _
        "WHERE tblPeople.IsDeceased=False AND ((tblPeople.IsClientDay=True AND tblPeopleClientsDayServices.Inactive=False) " & _
                                           "OR (tblPeople.IsClientVocat=True AND tblPeopleClientsVocationalServices.Inactive=False) " & _
                                           "OR (tblPeople.IsClientRes=True AND tblPeopleClientsResidentialServices.Inactive=False) " & _
                                           "OR (tblPeople.IsCilentCLO=True AND tblPeopleClientsCLOServices.Inactive=False) " & _
                                           "OR (tblPeople.IsClientSharedLiving=True AND tblPeopleClientsSharedLivingServices.Inactive=False) " & _
                                           "OR (tblPeople.IsClientAFC=True AND tblPeopleClientsAFCServices.Inactive=False) " & _
                                           "OR (tblPeople.IsClientNHDay=True AND tblPeopleClientsNHDay.Inactive=False)" & _
                                           "OR (tblPeople.IsClientNHRes=True AND tblPeopleClientsNHRes.Inactive=False)" & _
                                           "OR (tblPeople.IsClientIndiv=True AND tblPeopleClientsIndividualSupportServices.Inactive=False)) ", dbSeeChanges: Call BriefDelay
    SysCmdResult = SysCmd(4, "Building export table.")
    TILLDataBase.Execute "CREATE TABLE temptbl2 (MailingAddress CHAR(50), MailingCity CHAR(25), MailingState CHAR(2), MailingZIP CHAR(10), LN1 CHAR(25), " & _
                                       "FN1 CHAR(25), CO1 CHAR(50), EmailAddress1 CHAR(50), FamilyMemberIsGuardian1 BIT, LN2 CHAR(25), FN2 CHAR(25), " & _
                                       "CO2 CHAR(50), EmailAddress2 CHAR(50), FamilyMemberIsGuardian2 BIT, " & _
                                       "DH CHAR(3), Voc CHAR(3), Res CHAR(3), CLO CHAR(3), SL CHAR(3), AFC CHAR(3), NHD CHAR(3), NHR CHAR(3), ISS CHAR(3)) ", dbSeeChanges: Call BriefDelay
    TILLDataBase.Execute "ALTER TABLE temptbl2 ADD CONSTRAINT temptbl2constraint PRIMARY KEY (MailingAddress,MailingCity,MailingState,MailingZIP);", dbSeeChanges: Call BriefDelay
    TILLDataBase.Execute "INSERT INTO temptbl2 (MailingAddress, MailingCity, MailingState, MailingZIP, DH, Voc, Res, SL, AFC, NHD, NHR, CLO, ISS) " & vbCrLf & _
        "SELECT MailingAddress, MailingCity, MailingState, MailingZIP, " & _
               "IIf(IsClientDay,""Day"",Null) AS DH, " & _
               "IIf(IsClientVocat,""Voc"",Null) AS Voc, " & _
               "IIf(IsClientRes,""Res"",Null) AS Res, " & _
               "IIf(IsClientSharedLiving,""SL"",Null) AS SL, " & _
               "IIf(IsClientAFC,""AFC"",Null) AS AFC, " & _
               "IIf(IsClientNHDay,""NHD"",Null) AS NHD, " & _
               "IIf(IsClientNHRes,""NHR"",Null) AS NHR, " & _
               "IIf(IsClientCLO,""CLO"",Null) AS CLO, " & _
               "IIf(IsClientIndiv,""ISS"",Null) AS ISS " & _
               "FROM temptbl1" & vbCrLf & _
        "WHERE Left(MailingAddress,1) <> ' '", dbSeeChanges: Call BriefDelay
    Call ParseFamilyMembersAndThenExport(0, "temptbl2", "IntEvalsFamilyMailing")
    SysCmdResult = SysCmd(5)
    Exit Sub
ShowMeError:
    MsgBox "Error # " & Str(Err.Number) & " was generated by " & Me.Name & Chr(13) & Err.Description, vbOKOnly, "Error", Err.HelpFile, Err.HelpContext
End Sub

Private Sub ISPDeadlines_Click()
    Call DropTempTables
    DoCmd.SetWarnings False
    DoCmd.OpenQuery "qryISPInfo", acViewNormal
    DoCmd.SetWarnings True
    Call ParseFamilyMembersAndThenExport(9, "temptbl", "ISPDocumentationDeadlines")
    SysCmdResult = SysCmd(5)
End Sub

Private Sub ProgramExpirations_Click()
    Call DropTempTables
    DoCmd.SetWarnings False
    DoCmd.OpenQuery "qryProgramExpirations", acViewNormal
    DoCmd.SetWarnings True
    Call ParseFamilyMembersAndThenExport(9, "temptbl", "ProgramExpirations")
    SysCmdResult = SysCmd(5)
End Sub

Private Sub REPORTS_Click()
    If CurrentProject.AllForms("frmRpt").IsLoaded Then Form_frmRpt.SetFocus
End Sub

Private Sub SelectedCityTownRes_AfterUpdate()
    SelectedLocationRes.Requery
End Sub

Private Sub SURVEY_Click()
On Error GoTo ShowMeError
    Call DropTempTables
    SysCmdResult = SysCmd(4, "Gathering family members.")
    TILLDataBase.Execute "SELECT tblPeople.TILLEGram, tblPeople.LastName, tblPeople.FirstName, tblPeople.CompanyOrganization, tblPeople.MiddleInitial, tblPeople.FamiliarGreeting, tblPeople.MailingAddress, tblPeople.MailingCity, tblPeople.MailingState, tblPeople.MailingZIP, tblPeople.IsFamilyGuardian, tblPeople.IsDonor, tblPeople.IsInterestedParty, tblPeople.IsConsultant, tblPeople.InterestedPartyCategory, tblPeople.InterestedPartyInactive, tblPeople.GoGreen, tblPeople.EmailAddress, tblPeopleFamily.Relationship , tblPeopleFamily.ClientIndexedName, tblPeopleFamily.ClientLastName, tblPeopleFamily.ClientFirstName, tblPeopleFamily.ClientMiddleInitial, tblPeopleFamily.Guardian, tblPeopleFamily.PrimaryContact " & _
        "INTO temptbl0 " & _
        "FROM tblPeople LEFT JOIN tblPeopleFamily ON tblPeople.IndexedName = tblPeopleFamily.IndexedName " & _
        "WHERE tblPeopleFamily.Inactive = False AND tblPeople.IsFamilyGuardian = True AND tblPeople.IsDeceased = False AND tblPeople.NoMailings = False " & _
        "ORDER BY tblPeople.IsInterestedParty DESC , tblPeople.IsConsultant DESC , tblPeople.IndexedName;", dbSeeChanges: Call BriefDelay
    
    SysCmdResult = SysCmd(4, "Gathering clients.")
    DoCmd.SetWarnings False
    DoCmd.OpenQuery "qrySurvey", acViewNormal
    DoCmd.SetWarnings True
    
    SysCmdResult = SysCmd(4, "Creating export table.")
    TILLDataBase.Execute "CREATE TABLE temptbl2 (MailingAddress CHAR(50), MailingCity CHAR(25), MailingState CHAR(2), MailingZIP CHAR(10), " & _
                                       "ClientIndexedName CHAR(160), " & _
                                       "LN1 CHAR(25), FN1 CHAR(25), MI1 CHAR(1), CO1 CHAR(50), EmailAddress1 CHAR(50), " & _
                                       "LN2 CHAR(25), FN2 CHAR(25), MI2 CHAR(1), CO2 CHAR(50), EmailAddress2 CHAR(50), " & _
                                       "TILLEGram BIT, ClientLastName CHAR(25), ClientFirstName CHAR(25), ClientMiddleInitial CHAR(1), " & _
                                       "IsClientDay CHAR(8), IsClientVocat CHAR(8), DayProgram CHAR(60), IsClientRes CHAR(8), IsClientCLO CHAR(8), " & _
                                       "ResProgram CHAR(60), IsClientSharedLiving CHAR(8), IsClientNHDay CHAR(8), IsClientNHRes BIT, IsClientIndiv CHAR(8) );", dbSeeChanges: Call BriefDelay
    
    TILLDataBase.Execute "ALTER TABLE temptbl2 ADD CONSTRAINT temptbl2constraint PRIMARY KEY (MailingAddress,MailingCity,MailingState,MailingZIP,ClientIndexedName);", dbSeeChanges: Call BriefDelay
    SysCmdResult = SysCmd(4, "Pre-fill export table.")
    TILLDataBase.Execute "INSERT INTO temptbl2 (MailingAddress, MailingCity, MailingState, MailingZIP, ClientIndexedName, TILLEGram, " & _
        "ClientLastName, ClientFirstName, ClientMiddleInitial, IsClientDay, IsClientVocat, DayProgram, IsClientRes, ResProgram, IsClientSharedLiving, " & _
        "IsClientNHDay, IsClientNHRes, IsClientCLO, IsClientIndiv ) " & vbCrLf & _
        "SELECT MailingAddress, MailingCity, MailingState, MailingZIP, ClientIndexedName, TILLEGram, ClientLastName, ClientFirstName, " & _
               "ClientMiddleInitial , IsClientDay, IsClientVocat, DayVendorLocation, IsClientRes, ResVendorLocation, IsClientSharedLiving, " & _
               "IsClientNHDay, IsClientNHRes, IsClientCLO, IsClientIndiv " & vbCrLf & _
               "FROM temptbl1" & vbCrLf & _
        "WHERE Left(MailingAddress,1) <> ' '", dbSeeChanges: Call BriefDelay
    
    Call ParseFamilyMembersAndThenExport(1, "temptbl2", "Survey")
    SysCmdResult = SysCmd(5)
    Exit Sub
ShowMeError:
    MsgBox "Error # " & Str(Err.Number) & " was generated by " & Me.Name & Chr(13) & Err.Description, vbOKOnly, "Error", Err.HelpFile, Err.HelpContext
End Sub

Private Sub TILLEgram_Click()
On Error GoTo ShowMeError
    Call DropTempTables
    SysCmdResult = SysCmd(4, "Gathering active clients.")
    TILLDataBase.Execute "SELECT tblPeople.IndexedName, tblPeople.IsClientAutism, tblPeople.IsCilentCLO, tblPeople.IsClientDay, tblPeople.IsClientIndiv, tblPeople.IsClientNHDay, tblPeople.IsClientNHRes, tblPeople.IsClientRes, tblPeople.IsClientSharedLiving, tblPeople.IsClientSpring, tblPeople.IsClientTRASE, tblPeople.IsClientVocat " & _
        "INTO temptbl0 " & _
        "FROM (((((((((((tblPeople LEFT JOIN tblPeopleClientsDemographics ON tblPeople.IndexedName = tblPeopleClientsDemographics.IndexedName) LEFT JOIN tblPeopleClientsAutismServices ON tblPeopleClientsDemographics.IndexedName = tblPeopleClientsAutismServices.IndexedName) " & _
        "LEFT JOIN tblPeopleClientsCLOServices ON tblPeopleClientsDemographics.IndexedName = tblPeopleClientsCLOServices.IndexedName) LEFT JOIN tblPeopleClientsDayServices ON tblPeopleClientsDemographics.IndexedName = tblPeopleClientsDayServices.IndexedName) " & _
        "LEFT JOIN tblPeopleClientsIndividualSupportServices ON tblPeopleClientsDemographics.IndexedName = tblPeopleClientsIndividualSupportServices.IndexedName) LEFT JOIN tblPeopleClientsNHDay ON tblPeopleClientsDemographics.IndexedName = tblPeopleClientsNHDay.IndexedName) " & _
        "LEFT JOIN tblPeopleClientsNHRes ON tblPeopleClientsDemographics.IndexedName = tblPeopleClientsNHRes.IndexedName) LEFT JOIN tblPeopleClientsResidentialServices ON tblPeopleClientsDemographics.IndexedName = tblPeopleClientsResidentialServices.IndexedName) " & _
        "LEFT JOIN tblPeopleClientsSpringboardServices ON tblPeopleClientsDemographics.IndexedName = tblPeopleClientsSpringboardServices.IndexedName) LEFT JOIN tblPeopleClientsTRASEServices ON tblPeopleClientsDemographics.IndexedName = tblPeopleClientsTRASEServices.IndexedName) " & _
        "LEFT JOIN tblPeopleClientsVocationalServices ON tblPeopleClientsDemographics.IndexedName = tblPeopleClientsVocationalServices.IndexedName) LEFT JOIN tblPeopleClientsSharedLivingServices ON tblPeopleClientsDemographics.IndexedName = tblPeopleClientsSharedLivingServices.IndexedName " & _
        "WHERE tblPeople.IsDeceased = False AND " & _
             "((tblPeople.IsCilentCLO = True And tblPeopleClientsCLOServices.Inactive = False) Or " & _
              "(tblPeople.IsClientDay = True And tblPeopleClientsDayServices.Inactive = False) Or " & _
              "(tblPeople.IsClientIndiv = True And tblPeopleClientsIndividualSupportServices.Inactive = False) Or " & _
              "(tblPeople.IsClientNHDay = True And tblPeopleClientsNHDay.Inactive = False) Or " & _
              "(tblPeople.IsClientNHRes = True And tblPeopleClientsNHRes.Inactive = False) Or " & _
              "(tblPeople.IsClientRes = True And tblPeopleClientsResidentialServices.Inactive = False) Or " & _
              "(tblPeople.IsClientSharedLiving = True And tblPeopleClientsSharedLivingServices.Inactive = False) Or " & _
              "(tblPeople.IsClientSpring = True And tblPeopleClientsSpringboardServices.Inactive = False) Or " & _
              "(tblPeople.IsClientTRASE = True And tblPeopleClientsTRASEServices.Inactive = False) OR " & _
              "(tblPeople.IsClientVocat=True AND tblPeopleClientsVocationalServices.Inactive=False)) " & _
        "ORDER BY tblPeople.IndexedName;", dbSeeChanges: Call BriefDelay
    TILLDataBase.Execute "ALTER TABLE temptbl0 ADD CONSTRAINT temptbl1constraint PRIMARY KEY (IndexedName);", dbSeeChanges: Call BriefDelay
    SysCmdResult = SysCmd(4, "Gathering active family members from list of active clients.")
    TILLDataBase.Execute "SELECT DISTINCT tblPeople.MailingAddress, tblPeople.MailingCity, tblPeople.MailingState, tblPeople.MailingZIP, tblPeople.IndexedName, tblPeople.LastName AS LN, tblPeople.FirstName AS FN, tblPeople.MiddleInitial AS MI, tblPeople.Salutation AS SAL, tblPeople.CompanyOrganization AS CO, tblPeople.Title AS TITLE, tblPeople.EmailAddress, 'Family' AS RecordType " & _
        "INTO temptbl1 " & _
        "FROM (tblPeople INNER JOIN tblPeopleFamily ON tblPeople.IndexedName = tblPeopleFamily.IndexedName) INNER JOIN temptbl0 ON tblPeopleFamily.ClientIndexedName = temptbl0.IndexedName " & _
        "WHERE tblPeople.MailingAddress Is Not Null And Len([MailingAddress]) > 0 And tblPeople.TILLEGram = True And tblPeople.NoMailings = False And tblPeople.IsDeceased = False And tblPeople.IsFamilyGuardian = True And tblPeopleFamily.Inactive = False " & _
        "ORDER BY tblPeople.MailingAddress, tblPeople.MailingCity, tblPeople.MailingState, tblPeople.MailingZIP, tblPeople.IndexedName;", dbSeeChanges: Call BriefDelay
    ' Build new table to accommodate multiple names per address.
    SysCmdResult = SysCmd(4, "Building and seeding export table.")
    TILLDataBase.Execute "CREATE TABLE temptbl (MailingAddress CHAR(50), MailingCity CHAR(25), MailingState CHAR(2), MailingZIP CHAR(10), SAL1 CHAR(8), LN1 CHAR(30), " & _
        "FN1 CHAR(30), CO1 CHAR(50), Title CHAR(40), EmailAddress1 CHAR(60), SAL2 CHAR(8), LN2 CHAR(30), FN2 CHAR(30), EmailAddress2 CHAR(60), RecordType CHAR(10))", dbSeeChanges: Call BriefDelay
    TILLDataBase.Execute "ALTER TABLE temptbl ADD CONSTRAINT temptblconstraint PRIMARY KEY (MailingAddress,MailingCity,MailingState,MailingZIP);", dbSeeChanges: Call BriefDelay
    TILLDataBase.Execute "INSERT INTO temptbl (MailingAddress, MailingCity, MailingState, MailingZIP, RecordType) " & vbCrLf & _
        "SELECT MailingAddress, MailingCity, MailingState, MailingZIP, RecordType " & _
               "FROM temptbl1;", dbSeeChanges: Call BriefDelay
    SysCmdResult = SysCmd(4, "Updating export table with family members.")
    Call ParseFamilyMembersAndThenExport(2, "temptbl", "TILLEGramMailing", False)
    ' Append interested parties.
    SysCmdResult = SysCmd(4, "Appending interested parties.")
    TILLDataBase.Execute "ALTER TABLE temptbl DROP CONSTRAINT temptblconstraint;", dbSeeChanges: Call BriefDelay
    TILLDataBase.Execute "INSERT INTO temptbl ( MailingAddress, MailingCity, MailingState, MailingZIP, LN1, FN1, SAL1, CO1, Title, RecordType ) " & _
        "SELECT DISTINCT tblPeople.MailingAddress AS MailingAddress, tblPeople.MailingCity AS MailingCity, tblPeople.MailingState AS MailingState, tblPeople.MailingZIP AS MailingZIP, tblPeople.LastName AS LN1, tblPeople.FirstName AS FN1, tblPeople.Salutation AS SAL1, tblPeople.CompanyOrganization AS CO1, tblPeople.Title AS Title, 'IP' AS RecordType " & _
        "FROM tblPeople WHERE tblPeople.MailingAddress Is Not Null AND Len([tblPeople.MailingAddress])>0 AND tblPeople.TILLEGram=True AND tblPeople.NoMailings=False AND ((tblPeople.IsDeceased)=False) AND tblPeople.IsInterestedParty=True AND tblPeople.InterestedPartyInactive=False;", dbSeeChanges: Call BriefDelay
    Call ParseFamilyMembersAndThenExport(9, "temptbl", "TILLEgramMailing")
    Call DropTempTables
    SysCmdResult = SysCmd(5)
    Exit Sub
ShowMeError:
    MsgBox "Error # " & Str(Err.Number) & " was generated by " & Me.Name & Chr(13) & Err.Description, vbOKOnly, "Error", Err.HelpFile, Err.HelpContext
End Sub

Private Sub SelectedDayHab_Click()
On Error GoTo ShowMeError
    If IsNull(Form_frmRptExpMAILINGSANDSPREADSHEETS.SelectedCityTownDay) Or IsNull(Form_frmRptExpMAILINGSANDSPREADSHEETS.SelectedLocationDay) Then
        MsgBox "You must provide a city/town and a location from the drop down menus to the left.", vbOKOnly, "ERROR!"
        Exit Sub
    End If
    Call DropTempTables
    
    SysCmdResult = SysCmd(4, "Gathering family information.")
    TILLDataBase.Execute "SELECT tblPeople.IndexedName AS ClientIndexedName, tblPeople.LastName AS ClientLastName, tblPeople.FirstName AS ClientFirstName, tblPeople.MiddleInitial AS ClientMiddleInitial, tblPeopleFamily.IndexedName AS FamilyIndexedName, tblPeopleFamily.Guardian AS Guardian " & _
        "INTO temptbl0 " & _
        "FROM ((tblPeople LEFT JOIN tblPeopleFamily ON tblPeople.IndexedName = tblPeopleFamily.ClientIndexedName) LEFT JOIN tblPeopleClientsDayServices ON tblPeople.IndexedName = tblPeopleClientsDayServices.IndexedName) LEFT JOIN tblPeopleClientsVocationalServices ON tblPeople.IndexedName = tblPeopleClientsVocationalServices.IndexedName " & _
        "WHERE tblPeopleFamily.Inactive=False AND tblPeople.NoMailings=False AND tblPeople.IsDeceased=False AND " & _
             "((tblPeopleFamily.PrimaryContact=True AND tblPeople.IsClientDay=True   AND tblPeopleClientsDayServices.Inactive=False) OR " & _
              "(tblPeopleFamily.Guardian=True       AND tblPeople.IsClientDay=True   AND tblPeopleClientsDayServices.Inactive=False) OR " & _
              "(tblPeopleFamily.PrimaryContact=True AND tblPeople.IsClientVocat=True AND tblPeopleClientsVocationalServices.Inactive=False) OR " & _
              "(tblPeopleFamily.Guardian=True       AND tblPeople.IsClientVocat=True AND tblPeopleClientsVocationalServices.Inactive=False));", dbSeeChanges: Call BriefDelay
    Loc = Form_frmRptExpMAILINGSANDSPREADSHEETS.SelectedCityTownDay & "-" & Form_frmRptExpMAILINGSANDSPREADSHEETS.SelectedLocationDay
    
    SysCmdResult = SysCmd(4, "Gathering client information for selected day program.")
    TILLDataBase.Execute "CREATE TABLE temptbl1 (ClientIndexedName CHAR(160), Location CHAR(50), ClientLastName CHAR(25), ClientFirstName CHAR(25), ClientMiddleInitial CHAR(1), FamilyIndexedName CHAR(160), Relationship CHAR(25), PrimaryContact BIT, Guardian BIT, LN CHAR(25), FN CHAR(25), MI CHAR(1), CO CHAR(40), FamilyFamiliarGreeting CHAR(40), MailingAddress CHAR(50), MailingCity CHAR(25), MailingState CHAR(2), MailingZIP CHAR(10), GoGreen BIT, EmailAddress CHAR(50));", dbSeeChanges: Call BriefDelay
    TILLDataBase.Execute "INSERT INTO temptbl1 (ClientIndexedName, Location, ClientLastName, ClientFirstName, ClientMiddleInitial, FamilyIndexedName, Guardian, LN, FN, MI, CO, FamilyFamiliarGreeting, MailingAddress, MailingCity, MailingState, MailingZIP, GoGreen, EmailAddress) " & _
        "SELECT temptbl0.ClientIndexedName, tblPeopleClientsVendors.DayVendorLocation, temptbl0.ClientLastName, temptbl0.ClientFirstName, temptbl0.ClientMiddleInitial, temptbl0.FamilyIndexedName, temptbl0.Guardian, tblPeople.LastName AS LN, tblPeople.FirstName AS FN, tblPeople.MiddleInitial AS MI, tblPeople.CompanyOrganization AS CO, tblPeople.FamiliarGreeting AS FamilyFamiliarGreeting, tblPeople.MailingAddress AS MailingAddress, tblPeople.MailingCity AS MailingCity, tblPeople.MailingState AS MailingState, tblPeople.MailingZIP AS MailingZIP, tblPeople.GoGreen, tblPeople.EmailAddress " & _
        "FROM (temptbl0 INNER JOIN tblPeopleClientsVendors ON temptbl0.ClientIndexedName = tblPeopleClientsVendors.IndexedName) INNER JOIN tblPeople ON temptbl0.FamilyIndexedName = tblPeople.IndexedName " & _
        "WHERE tblPeopleClientsVendors.DayVendorLocation='" & Loc & "' AND tblPeople.IsDeceased=False;", dbSeeChanges: Call BriefDelay
    
    SysCmdResult = SysCmd(4, "Building export table.")
    TILLDataBase.Execute "CREATE TABLE temptbl2 (MailingAddress CHAR(50), MailingCity CHAR(25), MailingState CHAR(2), MailingZIP CHAR(10), ClientIndexedName CHAR(160), Location CHAR(50), LN1 CHAR(25), " & _
        "FN1 CHAR(25), MI1 CHAR(1), CO1 CHAR(50), FamilyFamiliarGreeting1 CHAR(30), FamilyMemberIsGuardian1 BIT, FamilyGoGreen1 BIT, EmailAddress1 CHAR(50), LN2 CHAR(25), FN2 CHAR(25), MI2 CHAR(1), " & _
        "CO2 CHAR(50), FamilyFamiliarGreeting2 CHAR(30), FamilyMemberIsGuardian2 BIT, FamilyGoGreen2 BIT, EmailAddress2 CHAR(50), LN3 CHAR(25), FN3 CHAR(25), MI3 CHAR(1), CO3 CHAR(50), FamilyFamiliarGreeting3 CHAR(30), FamilyMemberIsGuardian3 BIT, FamilyGoGreen3 BIT, EmailAddress3 CHAR(50), LN4 CHAR(25), FN4 CHAR(25), MI4 CHAR(1), CO4 CHAR(50), FamilyFamiliarGreeting4 CHAR(30), FamilyMemberIsGuardian4 BIT, FamilyGoGreen4 BIT, EmailAddress4 CHAR(50) );", dbSeeChanges: Call BriefDelay
    TILLDataBase.Execute "ALTER TABLE temptbl2 ADD CONSTRAINT temptbl2constraint PRIMARY KEY (MailingAddress, MailingCity, MailingState, MailingZIP, ClientIndexedName);", dbSeeChanges: Call BriefDelay
    TILLDataBase.Execute "INSERT INTO temptbl2 (MailingAddress, MailingCity, MailingState, MailingZIP, ClientIndexedName, Location) " & _
        "SELECT MailingAddress, MailingCity, MailingState, MailingZIP, ClientIndexedName, Location FROM temptbl1;", dbSeeChanges: Call BriefDelay
    Call ParseFamilyMembersAndThenExport(0, "temptbl2", "DayProgramNewsletters-" & SelectedCityTownDay & "-" & SelectedLocationDay)
    
    SysCmdResult = SysCmd(5)
    Exit Sub
ShowMeError:
    MsgBox "Error # " & Str(Err.Number) & " was generated by " & Me.Name & Chr(13) & Err.Description, vbOKOnly, "Error", Err.HelpFile, Err.HelpContext
End Sub

Private Sub SelectedCityTownDay_AfterUpdate()
On Error GoTo ShowMeError
    If SelectedCityTownDay = "Billerica" Then SelectedLocationDay.RowSource = "'Day Hab';'TILL Central';'TILL Central Autism Initiative'" Else SelectedLocationDay.RowSource = "'Day Hab';'TILL Central'"
    Exit Sub
ShowMeError:
    MsgBox "Error # " & Str(Err.Number) & " was generated by " & Me.Name & Chr(13) & Err.Description, vbOKOnly, "Error", Err.HelpFile, Err.HelpContext
End Sub

Private Sub SelectedResidence_Click()
On Error GoTo ShowMeError
    If IsNull(Form_frmRptExpMAILINGSANDSPREADSHEETS.SelectedCityTownRes) Or IsNull(Form_frmRptExpMAILINGSANDSPREADSHEETS.SelectedLocationRes) Then
        MsgBox "You must provide a city/town and a location from the drop down menus to the left.", vbOKOnly, "ERROR!"
        Exit Sub
    End If
    
    Call DropTempTables
    SysCmdResult = SysCmd(4, "Gathering family information.")
    TILLDataBase.Execute "SELECT tblPeople.IndexedName AS ClientIndexedName, tblPeople.LastName AS ClientLastName, tblPeople.FirstName AS ClientFirstName, tblPeople.MiddleInitial AS ClientMiddleInitial, tblPeopleFamily.IndexedName AS FamilyIndexedName, tblPeopleFamily.Guardian AS Guardian INTO temptbl0 " & _
        "FROM ((tblPeople LEFT JOIN tblPeopleFamily ON tblPeople.IndexedName = tblPeopleFamily.ClientIndexedName) LEFT JOIN tblPeopleClientsResidentialServices ON tblPeople.IndexedName = tblPeopleClientsResidentialServices.IndexedName) LEFT JOIN tblPeopleClientsCLOServices ON tblPeople.IndexedName = tblPeopleClientsCLOServices.IndexedName " & _
        "WHERE tblPeopleFamily.Inactive=False AND tblPeople.NoMailings=False AND tblPeople.IsDeceased=False AND " & _
             "((tblPeopleFamily.PrimaryContact=True AND tblPeople.IsClientRes=True AND tblPeopleClientsResidentialServices.Inactive=False) OR " & _
              "(tblPeopleFamily.Guardian=True       AND tblPeople.IsClientRes=True AND tblPeopleClientsResidentialServices.Inactive=False) OR " & _
              "(tblPeopleFamily.PrimaryContact=True AND tblPeople.IsCilentCLO=True AND tblPeopleClientsCLOServices.Inactive=False) OR " & _
              "(tblPeopleFamily.Guardian=True       AND tblPeople.IsCilentCLO=True AND tblPeopleClientsCLOServices.Inactive=False));", dbSeeChanges: Call BriefDelay
    SysCmdResult = SysCmd(4, "Gathering client information for selected residential program.")
    TILLDataBase.Execute "CREATE TABLE temptbl1 (ClientIndexedName CHAR(160), Location CHAR(50), ClientLastName CHAR(25), ClientFirstName CHAR(25), ClientMiddleInitial CHAR(1), FamilyIndexedName CHAR(160), Relationship CHAR(25), PrimaryContact BIT, Guardian BIT, LN CHAR(25), FN CHAR(25), MI CHAR(1), CO CHAR(40), FamilyFamiliarGreeting CHAR(50), MailingAddress CHAR(40), MailingCity CHAR(25), MailingState CHAR(2), MailingZIP CHAR(10), GoGreen BIT, EmailAddress CHAR(50));", dbSeeChanges: Call BriefDelay
    Loc = Form_frmRptExpMAILINGSANDSPREADSHEETS.SelectedCityTownRes & "-" & Form_frmRptExpMAILINGSANDSPREADSHEETS.SelectedLocationRes
    TILLDataBase.Execute "INSERT INTO temptbl1 (ClientIndexedName, Location, ClientLastName, ClientFirstName, ClientMiddleInitial, FamilyIndexedName, Guardian, LN, FN, MI, CO, FamilyFamiliarGreeting, MailingAddress, MailingCity, MailingState, MailingZIP, GoGreen, EmailAddress ) " & _
        "SELECT ClientIndexedName, tblPeopleClientsVendors.ResVendorLocation, ClientLastName, ClientFirstName, ClientMiddleInitial, FamilyIndexedName, Guardian, tblPeople.LastName AS LN, tblPeople.FirstName AS FN, " & _
        "tblPeople.MiddleInitial AS MI, tblPeople.CompanyOrganization AS CO, tblPeople.FamiliarGreeting AS FamilyFamiliarGreeting, tblPeople.MailingAddress AS MailingAddress, tblPeople.MailingCity AS MailingCity, tblPeople.MailingState AS MailingState, tblPeople.MailingZIP AS MailingZIP, tblPeople.GoGreen, tblPeople.EmailAddress " & _
        "FROM (temptbl0 INNER JOIN tblPeopleClientsVendors ON temptbl0.ClientIndexedName = tblPeopleClientsVendors.IndexedName) INNER JOIN tblPeople ON temptbl0.FamilyIndexedName = tblPeople.IndexedName " & _
        "WHERE tblPeopleClientsVendors.ResVendorLocation='" & Loc & "' AND tblPeople.IsDeceased=False;", dbSeeChanges: Call BriefDelay
    SysCmdResult = SysCmd(4, "Building export table.")
    TILLDataBase.Execute "CREATE TABLE temptbl2 (MailingAddress CHAR(50), MailingCity CHAR(25), MailingState CHAR(2), MailingZIP CHAR(10), ClientIndexedName CHAR(160), Location CHAR(50), LN1 CHAR(25), " & _
        "FN1 CHAR(25), MI1 CHAR(1), CO1 CHAR(50), FamilyFamiliarGreeting1 CHAR(30), FamilyMemberIsGuardian1 BIT, FamilyGoGreen1 BIT, EmailAddress1 CHAR(50), LN2 CHAR(25), FN2 CHAR(25), MI2 CHAR(1), " & _
        "CO2 CHAR(50), FamilyFamiliarGreeting2 CHAR(30), FamilyMemberIsGuardian2 BIT, FamilyGoGreen2 BIT, EmailAddress2 CHAR(50), LN3 CHAR(25), FN3 CHAR(25), MI3 CHAR(1), CO3 CHAR(50), FamilyFamiliarGreeting3 CHAR(30), FamilyMemberIsGuardian3 BIT, FamilyGoGreen3 BIT, EmailAddress3 CHAR(50), LN4 CHAR(25), FN4 CHAR(25), MI4 CHAR(1), CO4 CHAR(50), FamilyFamiliarGreeting4 CHAR(30), FamilyMemberIsGuardian4 BIT, FamilyGoGreen4 BIT, EmailAddress4 CHAR(50) );", dbSeeChanges: Call BriefDelay
    TILLDataBase.Execute "ALTER TABLE temptbl2 ADD CONSTRAINT temptbl2constraint PRIMARY KEY (MailingAddress, MailingCity, MailingState, MailingZIP, ClientIndexedName);", dbSeeChanges: Call BriefDelay
    TILLDataBase.Execute "INSERT INTO temptbl2 (MailingAddress, MailingCity, MailingState, MailingZIP, ClientIndexedName, Location) " & _
        "SELECT MailingAddress, MailingCity, MailingState, MailingZIP, ClientIndexedName, Location FROM temptbl1;", dbSeeChanges: Call BriefDelay
    Call ParseFamilyMembersAndThenExport(0, "temptbl2", "ResCLOProgramNewsletters-" & SelectedCityTownRes & "-" & SelectedLocationRes)
    SysCmdResult = SysCmd(5)
    SelectedCityTownRes = "": SelectedLocationRes = ""
    Exit Sub
ShowMeError:
    MsgBox "Error # " & Str(Err.Number) & " was generated by " & Me.Name & Chr(13) & Err.Description, vbOKOnly, "Error", Err.HelpFile, Err.HelpContext
End Sub

Private Sub SPRINGBOARD_Click()
On Error GoTo ShowMeError
    Call DropTempTables
    
    ' Here, we load up the Springboard group leaders.  A leader could be responsible for up to three different groups so three separate queries are needed.
    SysCmdResult = SysCmd(4, "Collect Springboard leaders.")
    TILLDataBase.Execute "SELECT tblPeopleConsultants.IndexedName, tblPeopleConsultants.SpringboardGroupCode1 AS SprGroup, tblPeople.LastName, tblPeople.FirstName " & _
        "INTO temptbl FROM tblPeopleConsultants INNER JOIN tblPeople ON tblPeopleConsultants.IndexedName = tblPeople.IndexedName " & _
        "WHERE tblPeopleConsultants.SpringboardGroupCode1 Is Not Null And tblPeopleConsultants.Department = 'Springboard' And tblPeopleConsultants.Inactive = False " & _
        "ORDER BY tblPeopleConsultants.SpringboardGroupCode1;", dbSeeChanges: Call BriefDelay
    TILLDataBase.Execute "INSERT INTO temptbl ( IndexedName, SprGroup, LastName, FirstName ) " & _
        "SELECT tblPeopleConsultants.IndexedName, tblPeopleConsultants.SpringboardGroupCode2 AS SprGroup, tblPeople.LastName, tblPeople.FirstName " & _
        "FROM tblPeopleConsultants INNER JOIN tblPeople ON tblPeopleConsultants.IndexedName = tblPeople.IndexedName " & _
        "WHERE tblPeopleConsultants.SpringboardGroupCode2 Is Not Null And tblPeopleConsultants.Department = 'Springboard' And tblPeopleConsultants.Inactive = False " & _
        "ORDER BY tblPeopleConsultants.SpringboardGroupCode2;", dbSeeChanges: Call BriefDelay
    TILLDataBase.Execute "INSERT INTO temptbl ( IndexedName, SprGroup, LastName, FirstName ) " & _
        "SELECT tblPeopleConsultants.IndexedName, tblPeopleConsultants.SpringboardGroupCode3 AS SprGroup, tblPeople.LastName, tblPeople.FirstName " & _
        "FROM tblPeopleConsultants INNER JOIN tblPeople ON tblPeopleConsultants.IndexedName = tblPeople.IndexedName " & _
        "WHERE tblPeopleConsultants.SpringboardGroupCode3 Is Not Null And tblPeopleConsultants.Department = 'Springboard' And tblPeopleConsultants.Inactive = False " & _
        "ORDER BY tblPeopleConsultants.SpringboardGroupCode3;", dbSeeChanges: Call BriefDelay

    ' Update the main Springboard client table, inserting the current leaders.
        SysCmdResult = SysCmd(4, "Update working table with leaders.")
    TILLDataBase.Execute "UPDATE qrytblPeopleClientsSpringboardServices " & _
        "INNER JOIN temptbl ON qrytblPeopleClientsSpringboardServices.GroupCode = temptbl.SprGroup " & _
        "SET qrytblPeopleClientsSpringboardServices.LeaderIndexedName = [temptbl].[IndexedName], " & _
        "qrytblPeopleClientsSpringboardServices.Leader = [temptbl].[FirstName] & ' ' & [temptbl].[LastName];", dbSeeChanges: Call BriefDelay
    
    ' Before adding the clients, we need to update their ages to ensure they are accurate.
    SysCmdResult = SysCmd(4, "Add clients to the export table.")
    TILLDataBase.Execute "UPDATE qrytblPeopleClientsDemographics " & _
        "RIGHT JOIN qrytblPeopleClientsSpringboardServices ON qrytblPeopleClientsDemographics.IndexedName = qrytblPeopleClientsSpringboardServices.IndexedName " & _
        "SET qrytblPeopleClientsSpringboardServices.Age = IIf(DateValue(Now())<DateSerial(Year(Now()),Month([qrytblPeopleClientsDemographics]![DateOfBirth]),Day([qrytblPeopleClientsDemographics]![DateOfBirth])),DateDiff(""yyyy"",DateValue([qrytblPeopleClientsDemographics]![DateOfBirth]),Now())-1,DateDiff(""yyyy"",DateValue([qrytblPeopleClientsDemographics]![DateOfBirth]),Now())) " & _
        "WHERE (((qrytblPeopleClientsDemographics.DateOfBirth) Is Not Null) AND ((qrytblPeopleClientsDemographics.ActiveSpringboard)=True));"
    
    ' Insert the clients.
    TILLDataBase.Execute "SELECT tblPeople.LastName, tblPeople.FirstName, tblPeople.MiddleInitial, tblPeople.MailingAddress, tblPeople.MailingCity, " & _
        "tblPeople.MailingState, tblPeople.MailingZIP, tblPeople.EmailAddress, tblPeople.HomePhone, tblPeople.MobilePhone, tblPeopleClientsDemographics.DateOfBirth, tblPeopleClientsSpringboardServices.Age, " & _
        "tblPeopleClientsSpringboardServices.GroupCode, tblPeopleClientsSpringboardServices.Leader, tblPeopleClientsSpringboardServices.BeginBillingDate, " & _
        "tblPeopleClientsSpringboardServices.DateTerminated" & vbCrLf & _
        "INTO temptbl1 FROM (tblPeople INNER JOIN tblPeopleClientsDemographics ON tblPeople.IndexedName = tblPeopleClientsDemographics.IndexedName) LEFT JOIN tblPeopleClientsSpringboardServices ON tblPeople.IndexedName = tblPeopleClientsSpringboardServices.IndexedName " & vbCrLf & _
        "WHERE tblPeopleClientsDemographics.ActiveSpringboard = True And tblPeople.IsDeceased = False And tblPeopleClientsSpringboardServices.Inactive = False ORDER BY tblPeopleClientsSpringboardServices.GroupCode, tblPeople.IndexedName;", dbSeeChanges: Call BriefDelay
    ' And insert the group leaders...three groups.
    SysCmdResult = SysCmd(4, "Assign group leader for each client.")
    TILLDataBase.Execute "INSERT INTO temptbl1 (LastName, FirstName, MiddleInitial, MailingAddress, MailingCity, MailingState, MailingZIP, HomePhone, MobilePhone, EmailAddress, GroupCode, Leader, BeginBillingDate, DateTerminated) " & _
        "SELECT tblPeople.LastName, tblPeople.FirstName, tblPeople.MiddleInitial, tblPeople.MailingAddress, tblPeople.MailingCity, tblPeople.MailingState, tblPeople.MailingZIP, tblPeople.HomePhone, tblPeople.MobilePhone, tblPeople.EmailAddress, tblPeopleConsultants.SpringboardGroupCode1, Null AS Leader, Null as BeginBillingDate, Null as DateTerminated " & _
        "FROM tblPeople INNER JOIN tblPeopleConsultants ON tblPeople.IndexedName = tblPeopleConsultants.IndexedName " & _
        "WHERE tblPeople.MailingAddress Is Not Null AND tblPeople.IsConsultant=True AND tblPeopleConsultants.SpringboardGroupCode1 Is Not Null AND tblPeopleConsultants.Inactive=False;", dbSeeChanges: Call BriefDelay
    TILLDataBase.Execute "INSERT INTO temptbl1 (LastName, FirstName, MiddleInitial, MailingAddress, MailingCity, MailingState, MailingZIP, HomePhone, MobilePhone, EmailAddress, GroupCode, Leader, BeginBillingDate, DateTerminated) " & _
        "SELECT tblPeople.LastName, tblPeople.FirstName, tblPeople.MiddleInitial, tblPeople.MailingAddress, tblPeople.MailingCity, tblPeople.MailingState, tblPeople.MailingZIP, tblPeople.HomePhone, tblPeople.MobilePhone, tblPeople.EmailAddress, tblPeopleConsultants.SpringboardGroupCode2, Null AS Leader, Null as BeginBillingDate, Null as DateTerminated  " & _
        "FROM tblPeople INNER JOIN tblPeopleConsultants ON tblPeople.IndexedName = tblPeopleConsultants.IndexedName " & _
        "WHERE tblPeople.MailingAddress Is Not Null AND tblPeople.IsConsultant=True AND tblPeopleConsultants.SpringboardGroupCode2 Is Not Null AND tblPeopleConsultants.Inactive=False;", dbSeeChanges: Call BriefDelay
    TILLDataBase.Execute "INSERT INTO temptbl1 (LastName, FirstName, MiddleInitial, MailingAddress, MailingCity, MailingState, MailingZIP, HomePhone, MobilePhone, EmailAddress, GroupCode, Leader,  BeginBillingDate, DateTerminated) " & _
        "SELECT tblPeople.LastName, tblPeople.FirstName, tblPeople.MiddleInitial, tblPeople.MailingAddress, tblPeople.MailingCity, tblPeople.MailingState, tblPeople.MailingZIP, tblPeople.HomePhone, tblPeople.MobilePhone, tblPeople.EmailAddress, tblPeopleConsultants.SpringboardGroupCode3, Null AS Leader, Null as BeginBillingDate, Null as DateTerminated  " & _
        "FROM tblPeople INNER JOIN tblPeopleConsultants ON tblPeople.IndexedName = tblPeopleConsultants.IndexedName " & _
        "WHERE tblPeople.MailingAddress Is Not Null AND tblPeople.IsConsultant=True AND tblPeopleConsultants.SpringboardGroupCode3 Is Not Null AND tblPeopleConsultants.Inactive=False;", dbSeeChanges: Call BriefDelay
    SysCmdResult = SysCmd(4, "Building export table.")
    Call ParseFamilyMembersAndThenExport(9, "temptbl1", "Springboard")
    SysCmdResult = SysCmd(5)
    Exit Sub
ShowMeError:
    MsgBox "Error # " & Str(Err.Number) & " was generated by " & Me.Name & Chr(13) & Err.Description, vbOKOnly, "Error", Err.HelpFile, Err.HelpContext
End Sub

Private Sub TILLGames_Click()
On Error GoTo ShowMeError
    Call DropTempTables
    SysCmdResult = SysCmd(4, "Identify family members.")
    TILLDataBase.Execute "SELECT tblPeopleFamily.* INTO temptbl " & _
        "FROM tblPeople RIGHT JOIN tblPeopleFamily ON tblPeople.IndexedName = tblPeopleFamily.IndexedName " & _
        "WHERE tblPeopleFamily.Inactive=False AND tblPeople.IsFamilyGuardian=True AND tblPeople.IsDeceased=False;", dbSeeChanges: Call BriefDelay
    SysCmdResult = SysCmd(4, "Identify clients.")
    TILLDataBase.Execute "SELECT DISTINCT tblPeople.IndexedName AS ClientIndexedName, tblPeople.LastName AS ClientLastName, tblPeople.FirstName AS ClientFirstName, tblPeople.MiddleInitial AS ClientMiddleInitial, tblPeople.PhysicalAddress AS ClientPhysicalAddress, tblPeople.PhysicalCity AS ClientPhysicalCity, tblPeople.PhysicalState AS ClientPhysicalState, tblPeople.PhysicalZIP AS ClientPhysicalZIP, tblPeople.MailingAddress AS ClientMailingAddress, tblPeople.MailingCity AS ClientMailingCity, tblPeople.MailingState AS ClientMailingState, tblPeople.MailingZIP AS ClientMailingZIP, tblPeopleFamily.IndexedName AS FamilyIndexedName, tblPeopleFamily.Relationship AS Expr1, tblPeopleFamily.Guardian AS Guardian, tblPeopleFamily.PrimaryContact AS PrimaryContact, tblPeople.IsCilentCLO AS IsClientCLO, tblPeople.IsClientRes AS IsClientRes INTO temptbl0 " & _
        "FROM (((tblPeople LEFT JOIN temptbl ON tblPeople.IndexedName = temptbl.ClientIndexedName) LEFT JOIN tblPeopleClientsCLOServices ON tblPeople.IndexedName = tblPeopleClientsCLOServices.IndexedName) LEFT JOIN tblPeopleClientsResidentialServices ON tblPeople.IndexedName = tblPeopleClientsResidentialServices.IndexedName) INNER JOIN tblPeopleFamily ON tblPeople.IndexedName = tblPeopleFamily.ClientIndexedName " & _
        "WHERE (((tblPeople.IsClientRes)=True) AND ((tblPeople.IsDeceased)=False) AND ((tblPeople.NoMailings)=False) AND ((tblPeopleClientsResidentialServices.Inactive)=False)) OR (((tblPeople.IsCilentCLO)=True) AND ((tblPeopleClientsCLOServices.Inactive)=False));", dbSeeChanges: Call BriefDelay
    SysCmdResult = SysCmd(4, "Merge clients and family members table.")
    TILLDataBase.Execute "SELECT temptbl0.ClientIndexedName, tblPeopleClientsVendors.ResVendorLocation AS Location, temptbl0.ClientLastName, temptbl0.ClientFirstName, temptbl0.ClientMiddleInitial, temptbl0.ClientPhysicalAddress, temptbl0.ClientPhysicalCity, temptbl0.ClientPhysicalState, temptbl0.ClientPhysicalZIP, temptbl0.ClientMailingAddress, temptbl0.ClientMailingCity, temptbl0.ClientMailingState, temptbl0.ClientMailingZIP, temptbl0.FamilyIndexedName, temptbl0.Expr1 AS Relationship, temptbl0.PrimaryContact, temptbl0.Guardian, tblPeople.LastName AS LN, tblPeople.FirstName AS FN, tblPeople.MiddleInitial AS MI, tblPeople.CompanyOrganization AS CO, tblPeople.FamiliarGreeting AS FamilyFamiliarGreeting, tblPeople.MailingAddress AS MailingAddress, tblPeople.MailingCity AS MailingCity, tblPeople.MailingState AS MailingState, tblPeople.MailingZIP AS MailingZIP, tblPeople.GoGreen, tblPeople.EmailAddress " & _
        "INTO temptbl1 " & _
        "FROM (temptbl0 INNER JOIN tblPeopleClientsVendors ON temptbl0.ClientIndexedName = tblPeopleClientsVendors.IndexedName) INNER JOIN tblPeople ON temptbl0.FamilyIndexedName = tblPeople.IndexedName " & _
        "WHERE tblPeople.IsDeceased=False AND tblPeopleClientsVendors.ResVendorLocation Not Like '*Auburn*' AND tblPeopleClientsVendors.ResVendorLocation Not Like '*Muzzey*';", dbSeeChanges: Call BriefDelay
    SysCmdResult = SysCmd(4, "Building export table.")
    TILLDataBase.Execute "CREATE TABLE temptbl2 (MailingAddress CHAR(50), MailingCity CHAR(25), MailingState CHAR(2), MailingZIP CHAR(10), ClientIndexedName CHAR(160), Location CHAR(50), LN1 CHAR(25)," & _
        "FN1 CHAR(25), MI1 CHAR(1), CO1 CHAR(50), FamilyFamiliarGreeting1 CHAR(30), FamilyGoGreen1 BIT, EmailAddress1 CHAR(50), LN2 CHAR(25), FN2 CHAR(25), MI2 CHAR(1), " & _
        "CO2 CHAR(50), FamilyFamiliarGreeting2 CHAR(30), FamilyGoGreen2 BIT, EmailAddress2 CHAR(50), LN3 CHAR(25), FN3 CHAR(25), MI3 CHAR(1), CO3 CHAR(50), FamilyFamiliarGreeting3 CHAR(30), FamilyGoGreen3 BIT, EmailAddress3 CHAR(50), LN4 CHAR(25), FN4 CHAR(25), MI4 CHAR(1), CO4 CHAR(50), FamilyFamiliarGreeting4 CHAR(30), FamilyGoGreen4 BIT, EmailAddress4 CHAR(50));", dbSeeChanges: Call BriefDelay
    TILLDataBase.Execute "ALTER TABLE temptbl2 ADD CONSTRAINT temptbl2constraint PRIMARY KEY (MailingAddress,MailingCity,MailingState,MailingZIP,ClientIndexedName);", dbSeeChanges: Call BriefDelay
    TILLDataBase.Execute "INSERT INTO temptbl2 (MailingAddress, MailingCity, MailingState, MailingZIP, ClientIndexedName, Location)" & _
        "SELECT MailingAddress, MailingCity, MailingState, MailingZIP, ClientIndexedName, Location FROM temptbl1;", dbSeeChanges: Call BriefDelay
    Call ParseFamilyMembersAndThenExport(0, "temptbl2", "CommunityConnections")
    SysCmdResult = SysCmd(5)
    Exit Sub
ShowMeError:
    MsgBox "Error # " & Str(Err.Number) & " was generated by " & Me.Name & Chr(13) & Err.Description, vbOKOnly, "Error", Err.HelpFile, Err.HelpContext
End Sub

Private Sub CloseForm_Click()
    If CurrentProject.AllForms("frmRptExpMAILINGSANDSPREADSHEETS").IsLoaded Then DoCmd.Close acForm, "frmRptExpMAILINGSANDSPREADSHEETS", acSaveYes
    If CurrentProject.AllForms("frmRpt").IsLoaded Then DoCmd.Close acForm, "frmRpt", acSaveYes
End Sub

Private Sub DNRTILLEGram_Click()
    DoCmd.OpenForm "frmRptExpDNRTILLEGRAM"
End Sub

Private Sub DNRAPPEAL_Click()
On Error GoTo ShowMeError
    Dim ExportFileName1 As String, ExportFileName2 As String, SQLCommand As Variant
    
    Call DropTempTables
    If IsTableQuery("tmpMostRecentDonations") Then TILLDataBase.Execute "DROP TABLE tmpMostRecentDonations;", dbSeeChanges: Call BriefDelay
    ' Seed Donor data and delete unneeded records.
    SysCmdResult = SysCmd(4, "Seed donor data and delete unneeded records..")
    
    TILLDataBase.Execute "SELECT tblPeopleDonors.*, tblPeople.EmailAddress, tblPeople.NoMailings, tblPeople.NoSolicitations INTO temptbl0 " & _
        "FROM tblPeople RIGHT JOIN tblPeopleDonors ON tblPeople.IndexedName = tblPeopleDonors.IndexedName " & _
        "WHERE tblPeopleDonors.Inactive=False;", dbSeeChanges: Call BriefDelay
    
    TILLDataBase.Execute "DELETE temptbl0.* FROM temptbl0 WHERE temptbl0.NoSolicitations = True;", dbSeeChanges: Call BriefDelay
    TILLDataBase.Execute "DELETE temptbl0.* FROM temptbl0 WHERE temptbl0.NoMailings = True;", dbSeeChanges: Call BriefDelay
    
    TILLDataBase.Execute "DELETE temptbl0.*, Year(DateValue([temptbl0]![DateOfDonation])) AS Expr1 FROM temptbl0 " & _
        "WHERE (((Year(DateValue([temptbl0]![DateOfDonation])))<=Year(Now())-5));", dbSeeChanges: Call BriefDelay
    TILLDataBase.Execute "DELETE temptbl0.*, Year(DateValue([temptbl0]![DateOfDonation])) AS Expr1, temptbl0.SolicitationType FROM temptbl0 " & _
        "WHERE (((Year(DateValue([temptbl0]![DateOfDonation])))<=Year(Now())-4) AND ((temptbl0.SolicitationType)='In Memory of'));", dbSeeChanges: Call BriefDelay
    
    ' Run queries to identify most recent donations.
    SysCmdResult = SysCmd(4, "Identify most recent donations.")
    DoCmd.SetWarnings False
    DoCmd.OpenQuery "qryDonorAppealCreateMostRecentDonations"
    
    TILLDataBase.Execute "SELECT tmpMostRecentDonations.IndexedName, tmpMostRecentDonations.LastDonationNumeric, tmpMostRecentDonations.LastDonationDate, tmpMostRecentDonations.FormattedDate, " & _
        "tmpMostRecentDonations.LastDonationAmount, tmpMostRecentDonations.CurrentOrLapsed, temptbl0.DonationType, temptbl0.DonationFrom1Salutation, temptbl0.DonationFrom1FirstName, " & _
        "temptbl0.DonationFrom1LastName, Null AS IndexedName1, temptbl0.EmailAddress, temptbl0.DonationFrom2Salutation, temptbl0.DonationFrom2FirstName, temptbl0.DonationFrom2LastName, " & _
        "Null AS IndexedName2, temptbl0.DonationFromCompany INTO temptbl2 " & _
        "FROM tmpMostRecentDonations RIGHT JOIN temptbl0 ON tmpMostRecentDonations.IndexedName = temptbl0.IndexedName " & _
        "WHERE (((tmpMostRecentDonations.FormattedDate) = [temptbl0].[DateOfDonation])) ORDER BY tmpMostRecentDonations.IndexedName;", dbSeeChanges: Call BriefDelay
    
    TILLDataBase.Execute "SELECT [tblPeople]![MailingState] & '/' & [tblPeople]![MailingCity] & '/' & [tblPeople]![MailingAddress] & '/' & [tblPeople].[IndexedName] AS AddressKey, " & _
        "IIf([temptbl2].[DonationType]=""Goods"" Or [temptbl2].[DonationType]=""Services"",""GS"",'') AS Segment1, " & _
        "IIf([Segment1]=""GS"",""GS"",IIf([temptbl2].[CurrentOrLapsed]=""Current"",""CD"",""LD"")) AS Segment2, " & _
        "IIf(Val([temptbl2].[LastDonationAmount])<=14,""LO"",IIf(Val([temptbl2].[LastDonationAmount])>14 And Val([temptbl2].[LastDonationAmount])<=999,""MD"",IIf(Val([temptbl2].[LastDonationAmount])>999,""HI"",''))) AS Segment3, " & _
        "IIf([Segment1]=""GS"",""GS"",[Segment2] & [Segment3]) AS Segment, " & _
        "temptbl2.DonationType, temptbl2.DonationFrom1Salutation, temptbl2.DonationFrom1FirstName, temptbl2.DonationFrom1LastName, temptbl2.IndexedName1, temptbl2.EmailAddress, " & _
        "False AS Deceased1, temptbl2.DonationFrom2Salutation, temptbl2.DonationFrom2FirstName, temptbl2.DonationFrom2LastName, temptbl2.IndexedName2, " & _
        "False AS Deceased2, tblPeople.Title, temptbl2.DonationFromCompany, tblPeople.MailingAddress, tblPeople.MailingCity, tblPeople.MailingState, " & _
        "tblPeople.MailingZIP, tblPeople.MailingAddressValidated, tblPeople.Comment, temptbl2.LastDonationDate, temptbl2.LastDonationAmount, " & _
        "tblPeople.TILLEGram, tblPeople.NoSolicitations INTO temptbl3 " & _
        "FROM (tblPeople LEFT JOIN tblPeopleFamily ON tblPeople.IndexedName = tblPeopleFamily.IndexedName) " & _
        "INNER JOIN temptbl2 ON tblPeople.IndexedName = temptbl2.IndexedName " & _
        "WHERE temptbl2.LastDonationNumeric > 0 And Len([MailingAddress]) > 0 And Len([MailingCity]) > 0 And Len([MailingState]) = 2 And tblPeople.IsDeceased = False " & _
        "ORDER BY [tblPeople]![MailingState] & '/' & [tblPeople]![MailingCity] & '/' & [tblPeople]![MailingAddress] & '/' & [tblPeople].[IndexedName], temptbl2.LastDonationDate DESC , temptbl2.DonationFrom1LastName, temptbl2.DonationFrom1FirstName;", dbSeeChanges: Call BriefDelay
    DoCmd.SetWarnings True
    
    ' Select fields from the last queries run and seed the table.
    SysCmdResult = SysCmd(4, "Select fields and insert into working table.")
    TILLDataBase.Execute "SELECT temptbl3.AddressKey, temptbl3.MailingAddress, temptbl3.MailingCity, temptbl3.MailingState, " & _
        "temptbl3.MailingZIP, temptbl3.DonationFrom1Salutation AS SAL1, temptbl3.DonationFrom1LastName AS LN1, temptbl3.DonationFrom1FirstName AS FN1, " & _
        "temptbl3.Title, temptbl3.DonationFromCompany AS CO1, temptbl3.EmailAddress, temptbl3.DonationFrom2Salutation AS SAL2, temptbl3.DonationFrom2LastName AS LN2, " & _
        "temptbl3.DonationFrom2FirstName AS FN2, Null as LetterSalutation, temptbl3.Comment, temptbl3.DonationType, temptbl3.LastDonationDate, " & _
        "temptbl3.LastDonationAmount, temptbl3.Segment, 'Donors' AS Category INTO temptbl4 " & _
        "FROM temptbl3 ORDER BY temptbl3.AddressKey;", dbSeeChanges: Call BriefDelay
    
    ' Empty the table and then set address contraint so that duplicates are removed.
    SysCmdResult = SysCmd(4, "Empty export table and set address constraint to remove duplicates.")
    TILLDataBase.Execute "DELETE * FROM temptbl4;", dbSeeChanges: Call BriefDelay
    TILLDataBase.Execute "ALTER TABLE temptbl4 ADD CONSTRAINT temptbl3constraint PRIMARY KEY (AddressKey);", dbSeeChanges: Call BriefDelay
    TILLDataBase.Execute "INSERT INTO temptbl4 ( AddressKey, MailingAddress, MailingCity, MailingState, MailingZIP, SAL1, LN1, FN1, Title, CO1, EmailAddress, SAL2, LN2, FN2, " & _
        "LetterSalutation, Comment, DonationType, LastDonationDate, LastDonationAmount, Segment, Category) " & _
        "SELECT temptbl3.AddressKey, temptbl3.MailingAddress, temptbl3.MailingCity, temptbl3.MailingState, temptbl3.MailingZIP, " & _
        "temptbl3.DonationFrom1Salutation AS SAL1, temptbl3.DonationFrom1LastName AS LN1, temptbl3.DonationFrom1FirstName AS FN1, temptbl3.Title, " & _
        "temptbl3.DonationFromCompany AS CO1, temptbl3.EmailAddress, temptbl3.DonationFrom2Salutation AS SAL2, temptbl3.DonationFrom2LastName AS LN2, " & _
        "temptbl3.DonationFrom2FirstName AS FN2, Null as LetterSalutation, temptbl3.Comment, temptbl3.DonationType, temptbl3.LastDonationDate, " & _
        "temptbl3.LastDonationAmount, temptbl3.Segment, 'Donors' AS Category " & _
        "FROM temptbl3 " & _
        "ORDER BY temptbl3.AddressKey;"
    
    ' Build letter salutations.  Multiple permutations.  Strip Jr, Sr, II, III, and IV from male last names.
    If Not DonorAppealFixSalutations("Export1") Then GoTo ShowMeError
    ' Export the Donors table.
    SysCmdResult = SysCmd(4, "Building export table.")
    ExportFileName1 = Application.CurrentProject.Path & "\" & "TILLDB-Export-DonorAppeal-Donors-" & Format(Date, "yyyymmdd") & ".xls"
    If IsFileOpen(ExportFileName1) Then
        If MsgBox(ExportFileName1 & " is already open.  Please close it and click OK to continue or Cancel to abort.", vbOKCancel, "ERROR!") = vbCancel Then
            MsgBox "Export aborted.", vbOKOnly, "Aborted"
            Exit Sub
        End If
    End If
    If Dir(ExportFileName1) <> "" Then Kill ExportFileName1
    DoCmd.TransferSpreadsheet acExport, acSpreadsheetTypeExcel9, "temptbl4", ExportFileName1
    Call DropTempTables
    
    ' =============================== Family Processing.
    SysCmdResult = SysCmd(4, "Processing family records.")
    
    ' Family records.
    TILLDataBase.Execute "SELECT tblPeople.IndexedName, tblPeople.TILLEGram, tblPeople.Salutation, tblPeople.LastName, tblPeople.FirstName, tblPeople.CompanyOrganization, tblPeople.MiddleInitial, tblPeople.Title, " & _
        "tblPeople.FamiliarGreeting, tblPeople.MailingAddress, tblPeople.MailingCity, tblPeople.MailingState, tblPeople.MailingZIP, tblPeople.IsFamilyGuardian, " & _
        "tblPeople.IsDonor, tblPeople.IsInterestedParty, tblPeople.IsConsultant, tblPeople.InterestedPartyCategory, tblPeople.InterestedPartyInactive, tblPeople.GoGreen, " & _
        "tblPeople.EmailAddress, tblPeopleFamily.Relationship , tblPeopleFamily.ClientIndexedName, tblPeopleFamily.ClientLastName, tblPeopleFamily.ClientFirstName, " & _
        "tblPeopleFamily.ClientMiddleInitial, tblPeopleFamily.Guardian, tblPeopleFamily.PrimaryContact, tblPeople.NoMailings, tblPeople.NoSolicitations " & _
        "INTO temptbl0" & vbCrLf & _
        "FROM tblPeople LEFT JOIN tblPeopleFamily ON tblPeople.IndexedName = tblPeopleFamily.IndexedName " & vbCrLf & _
        "WHERE tblPeopleFamily.Inactive = False AND tblPeople.IsFamilyGuardian = True AND tblPeople.IsDeceased = False " & vbCrLf & _
        "ORDER BY tblPeople.IsInterestedParty DESC , tblPeople.IsConsultant DESC , tblPeople.IndexedName;", dbSeeChanges: Call BriefDelay
    TILLDataBase.Execute "DELETE temptbl0.* FROM temptbl0 WHERE temptbl0.NoSolicitations = True;", dbSeeChanges: Call BriefDelay
    TILLDataBase.Execute "DELETE temptbl0.* FROM temptbl0 WHERE temptbl0.NoMailings = True;", dbSeeChanges: Call BriefDelay
    
    ' Add Client information.
    TILLDataBase.Execute "SELECT temptbl0.IndexedName, temptbl0.Salutation AS SAL, temptbl0.LastName AS LN, temptbl0.FirstName as FN, temptbl0.MiddleInitial AS MI, temptbl0.CompanyOrganization AS CO, temptbl0.Title, temptbl0.FamiliarGreeting AS FamilyFamiliarGreeting, temptbl0.MailingAddress AS MailingAddress, temptbl0.MailingCity AS MailingCity, temptbl0.MailingState AS MailingState, temptbl0.MailingZIP AS MailingZIP, temptbl0.IsFamilyGuardian, temptbl0.IsDonor, temptbl0.IsInterestedParty, temptbl0.IsConsultant, temptbl0.InterestedPartyCategory, temptbl0.InterestedPartyInactive, temptbl0.TILLEGram, temptbl0.GoGreen, temptbl0.EmailAddress, temptbl0.Relationship, temptbl0.Guardian, temptbl0.PrimaryContact, temptbl0.ClientIndexedName, temptbl0.ClientLastName, temptbl0.ClientFirstName, temptbl0.ClientMiddleInitial, tblPeople.IsClientDay, tblPeople.IsClientVocat, tblPeopleClientsVendors.DayVendorLocation, " & _
        "tblPeople.IsClientRes, tblPeople.IsCilentCLO AS IsClientCLO, tblPeopleClientsVendors.ResVendorLocation, tblPeople.IsClientTrans, tblPeople.IsClientSharedLiving, tblPeople.IsClientAFC, " & _
        "tblPeople.IsClientNHDay, tblPeople.IsClientNHRes, tblPeople.IsClientIndiv, tblPeople.IsClientAutism, tblPeople.IsClientPCA, tblPeople.IsClientSpring, " & _
        "tblPeople.IsClientTRASE, tblPeople.IsClientCommunityConnections, tblPeopleClientsVendors.LivingWithParentOrGuardian, tblPeopleClientsVendors.LivingIndependently " & vbCrLf & _
        "INTO temptbl1 FROM (((((((((((((((temptbl0 INNER JOIN tblPeople ON temptbl0.ClientIndexedName = tblPeople.IndexedName) INNER JOIN tblPeopleClientsVendors ON tblPeople.IndexedName = tblPeopleClientsVendors.IndexedName) LEFT JOIN tblPeopleClientsAutismServices ON temptbl0.ClientIndexedName = tblPeopleClientsAutismServices.IndexedName) LEFT JOIN tblPeopleClientsCLOServices ON temptbl0.ClientIndexedName = tblPeopleClientsCLOServices.IndexedName) LEFT JOIN tblPeopleClientsCommunityConnectionsServices ON temptbl0.ClientIndexedName = tblPeopleClientsCommunityConnectionsServices.IndexedName) LEFT JOIN tblPeopleClientsDayServices ON temptbl0.ClientIndexedName = tblPeopleClientsDayServices.IndexedName)  " & _
        "LEFT JOIN tblPeopleClientsIndividualSupportServices ON temptbl0.ClientIndexedName = tblPeopleClientsIndividualSupportServices.IndexedName) LEFT JOIN tblPeopleClientsPCAServices ON temptbl0.ClientIndexedName = tblPeopleClientsPCAServices.IndexedName) LEFT JOIN tblPeopleClientsResidentialServices ON temptbl0.ClientIndexedName = tblPeopleClientsResidentialServices.IndexedName) LEFT JOIN tblPeopleClientsSharedLivingServices ON temptbl0.ClientIndexedName = tblPeopleClientsSharedLivingServices.IndexedName) LEFT JOIN tblPeopleClientsAFCServices ON temptbl0.ClientIndexedName = tblPeopleClientsAFCServices.IndexedName) LEFT JOIN tblPeopleClientsSpringboardServices ON temptbl0.ClientIndexedName = tblPeopleClientsSpringboardServices.IndexedName) " & _
        "LEFT JOIN tblPeopleClientsTRASEServices ON temptbl0.ClientIndexedName = tblPeopleClientsTRASEServices.IndexedName) LEFT JOIN tblPeopleClientsVocationalServices ON temptbl0.ClientIndexedName = tblPeopleClientsVocationalServices.IndexedName) LEFT JOIN tblPeopleClientsNHRes ON temptbl0.ClientIndexedName = tblPeopleClientsNHRes.IndexedName) LEFT JOIN tblPeopleClientsNHDay ON temptbl0.ClientIndexedName = tblPeopleClientsNHDay.IndexedName " & vbCrLf & _
        "WHERE tblPeople.IsDeceased=False  AND ((tblPeople.IsClientDay=True AND tblPeopleClientsDayServices.Inactive=False) " & _
                                            "OR (tblPeople.IsClientVocat=True AND tblPeopleClientsVocationalServices.Inactive=False) " & _
                                            "OR (tblPeople.IsClientRes=True AND tblPeopleClientsResidentialServices.Inactive=False) " & _
                                            "OR (tblPeople.IsCilentCLO=True AND tblPeopleClientsCLOServices.Inactive=False) " & _
                                            "OR (tblPeople.IsClientSharedLiving=True AND tblPeopleClientsSharedLivingServices.Inactive=False) " & _
                                            "OR (tblPeople.IsClientAutism=True AND tblPeopleClientsAutismServices.Inactive=False) " & _
                                            "OR (tblPeople.IsClientAFC=True AND tblPeopleClientsAFCServices.Inactive=False) " & _
                                            "OR (tblPeople.IsClientNHDay=True AND tblPeopleClientsNHDay.Inactive=False)" & _
                                            "OR (tblPeople.IsClientNHRes=True AND tblPeopleClientsNHRes.Inactive=False)" & _
                                            "OR (tblPeople.IsClientIndiv=True AND tblPeopleClientsIndividualSupportServices.Inactive=False)) ", dbSeeChanges: Call BriefDelay
    
    ' Create export table.
    SysCmdResult = SysCmd(4, "Building export table.")
    TILLDataBase.Execute "CREATE TABLE temptbl (AddressKey CHAR (120), MailingAddress CHAR(50), MailingCity CHAR(25), MailingState CHAR(2), MailingZIP CHAR(10), IndexedName CHAR(160), SAL1 CHAR(8), LN1 CHAR(30), " & _
                                       "FN1 CHAR(30), Title CHAR(40), CO1 CHAR(50), EmailAddress1 CHAR(60), SAL2 CHAR(8), LN2 CHAR(30), FN2 CHAR(30), EmailAddress2 CHAR(60), LetterSalutation CHAR(50), Comment CHAR(1), DonationType CHAR(1), LastDonationDate CHAR(1), LastDonationAmount CHAR(1), Segment CHAR(1), " & _
                                       "Category CHAR(10), ISDonor BIT, DonorStatus CHAR(10))", dbSeeChanges: Call BriefDelay
    
    ' Set primary key to get rid of duplicate addresses.
    TILLDataBase.Execute "ALTER TABLE temptbl ADD CONSTRAINT temptblconstraint PRIMARY KEY (AddressKey);", dbSeeChanges: Call BriefDelay
    
    ' Insert family records into new table.
    TILLDataBase.Execute "INSERT INTO temptbl (AddressKey, MailingAddress, MailingCity, MailingState, MailingZIP, IndexedName, Category, Comment, DonationType, LastDonationDate, LastDonationAmount, Segment, IsDonor) " & vbCrLf & _
        "SELECT Trim(temptbl1.MailingState) & '/' & Trim(temptbl1.MailingCity) & '/' & Trim(temptbl1.MailingAddress) AS AddressKey, MailingAddress, MailingCity, MailingState, MailingZIP, IndexedName, 'Family' AS Category, Null as Comment, Null AS DonationType, Null AS LastDonationDate, Null AS LastDonationAmount, Null as Segment, IsDonor " & _
               "FROM temptbl1" & vbCrLf & "WHERE Left(MailingAddress,1) <> ' ' ORDER BY Trim(temptbl1.MailingState) & '/' & Trim(temptbl1.MailingCity) & '/' & Trim(temptbl1.MailingAddress)", dbSeeChanges: Call BriefDelay
    
    ' Add family names.
    Call ParseFamilyMembersAndThenExport(2, "temptbl", "", False)
    
    ' Remove primary key constraint.
    TILLDataBase.Execute "ALTER TABLE temptbl DROP CONSTRAINT temptblconstraint", dbSeeChanges: Call BriefDelay
    
    ' IP Processing.
    SysCmdResult = SysCmd(4, "Processing interested parties.")
    TILLDataBase.Execute "INSERT INTO temptbl (AddressKey, MailingAddress, MailingCity, MailingState, MailingZIP, IndexedName, SAL1, LN1, FN1, Title, CO1, EmailAddress1, SAL2, LN2, FN2, EmailAddress2, LetterSalutation, Category, Comment, DonationType, LastDonationDate, LastDonationAmount, Segment, IsDonor) " & _
        "SELECT Trim(tblPeople.MailingState) & '/' & Trim(tblPeople.MailingCity) & '/' & Trim(tblPeople.MailingAddress) AS AddressKey, tblPeople.MailingAddress, tblPeople.MailingCity, tblPeople.MailingState, tblPeople.MailingZIP, tblPeople.IndexedName, tblPeople.Salutation AS SAL1, tblPeople.LastName AS LN1, tblPeople.FirstName AS FN1, tblPeople.Title, tblPeople.CompanyOrganization AS CO1, tblPeople.EmailAddress AS EmailAddress1, Null AS SAL2, Null AS LN2, Null AS FN2, Null AS EmailAddress2, Null AS LetterSalutation, 'IP' As Category, Null as Comment, Null AS DonationType, Null AS LastDonationDate, Null AS LastDonationAmount, Null as Segment, IsDonor FROM tblPeople " & _
        "WHERE tblPeople.IsInterestedParty=True AND tblPeople.InterestedPartyInactive=False AND tblPeople.IsDeceased=False AND tblPeople.NoSolicitations=False ORDER BY Trim(tblPeople.MailingState) & '/' & Trim(tblPeople.MailingCity) & '/' & Trim(tblPeople.MailingAddress);", dbSeeChanges: Call BriefDelay
    
    ' Here, determine if donor is current or lapsed.
    TILLDataBase.Execute "UPDATE temptbl INNER JOIN tmpMostRecentDonations ON temptbl.IndexedName = tmpMostRecentDonations.IndexedName SET temptbl.DonorStatus = [tmpMostRecentDonations]![CurrentOrLapsed]" & _
        "WHERE (((temptbl.ISDonor)=True));", dbSeeChanges: Call BriefDelay
    
    ' Build letter salutations.  Multiple permutations.  Strip Jr, Sr, II, III, and IV from male last names.
    If Not DonorAppealFixSalutations("Export2") Then GoTo ShowMeError
    ' Export the Family/IP table.
    SysCmdResult = SysCmd(4, "Exporting family/IP table.")
    ExportFileName2 = Application.CurrentProject.Path & "\" & "TILLDB-Export-DonorAppeal-Families-IP-" & Format(Date, "yyyymmdd") & ".xls"
    If IsFileOpen(ExportFileName2) Then
        If MsgBox(ExportFileName2 & " is already open.  Please close it and click OK to continue or Cancel to abort.", vbOKCancel, "ERROR!") = vbCancel Then
            MsgBox "Export aborted.", vbOKOnly, "Aborted"
            Exit Sub
        End If
    End If
    If Dir(ExportFileName2) <> "" Then Kill ExportFileName2
    DoCmd.TransferSpreadsheet acExport, acSpreadsheetTypeExcel9, "temptbl", ExportFileName2
    ' Open Excel for each table.
    SysCmdResult = SysCmd(4, "Exporting donors table.")
'   CommandLine = LocateExecutable(ExportFileName1) & " """ & ExportFileName1 & """"
    MsgBox "The requested Donor information has been exported to " & ExportFileName1 & "." & vbCrLf & vbCrLf & "This export may contain information that is protected under HIPAA and other privacy laws.  This export must be securely stored at all times and must be deleted when no longer being used.", _
        vbOKOnly, "Export Complete"
'   RetValue = Shell(CommandLine, 1)
'   CommandLine = LocateExecutable(ExportFileName2) & " """ & ExportFileName2 & """"
    MsgBox "The requested Family/IP information has been exported to " & ExportFileName2 & "." & vbCrLf & vbCrLf & "This export may contain information that is protected under HIPAA and other privacy laws.  This export must be securely stored at all times and must be deleted when no longer being used.", _
        vbOKOnly, "Export Complete"
'   RetValue = Shell(CommandLine, 1)
    If IsTableQuery("tmpMostRecentDonations") Then TILLDataBase.Execute "DROP TABLE tmpMostRecentDonations;", dbSeeChanges: Call BriefDelay
    Call DropTempTables
    SysCmdResult = SysCmd(5)
    Exit Sub
ShowMeError:
    If IsTableQuery("tmpMostRecentDonations") Then TILLDataBase.Execute "DROP TABLE tmpMostRecentDonations;", dbSeeChanges: Call BriefDelay
    Call DropTempTables
    SysCmdResult = SysCmd(5)
    If Err.Number = 70 Then
        MsgBox "Please close all open exported spreadsheets then try again.", vbOKOnly, "Error!"
    Else
        MsgBox "Error # " & Str(Err.Number) & " was generated by " & Me.Name & Chr(13) & Err.Description, vbOKOnly, "Error", Err.HelpFile, Err.HelpContext
    End If
End Sub

Private Function DonorAppealFixSalutations(WhichExport As Variant) As Boolean
On Error GoTo ShowMeError
    SysCmdResult = SysCmd(4, "Fixing salutations.")
    Select Case WhichExport
        Case "Export1"
            TILLDataBase.Execute "UPDATE temptbl4 SET temptbl4.LN1 = Left(Trim([temptbl4].[LN1]),Len(Trim([temptbl4].[LN1]))-3) " & _
                "WHERE Right(Trim([temptbl4].[LN1]),3) = ' Jr' OR " & _
                      "Right(Trim([temptbl4].[LN1]),3) = ' Sr' OR " & _
                      "Right(Trim([temptbl4].[LN1]),3) = ' II' OR " & _
                      "Right(Trim([temptbl4].[LN1]),3) = ' IV';", dbSeeChanges: Call BriefDelay
            TILLDataBase.Execute "UPDATE temptbl4 SET temptbl4.LN1 = Left(Trim([temptbl4].[LN1]),Len(Trim([temptbl4].[LN1]))-4) WHERE Right(Trim([temptbl4].[LN1]),4) = ' III';", dbSeeChanges: Call BriefDelay
            TILLDataBase.Execute "UPDATE temptbl4 SET temptbl4.LN2 = Left(Trim([temptbl4].[LN2]),Len(Trim([temptbl4].[LN2]))-3) " & _
                "WHERE Right(Trim([temptbl4].[LN2]),3) = ' Jr' OR " & _
                      "Right(Trim([temptbl4].[LN2]),3) = ' Sr' OR " & _
                      "Right(Trim([temptbl4].[LN1]),3) = ' II' OR " & _
                      "Right(Trim([temptbl4].[LN1]),3) = ' IV';", dbSeeChanges: Call BriefDelay
            TILLDataBase.Execute "UPDATE temptbl4 SET temptbl4.LN2 = Left(Trim([temptbl4].[LN2]),Len(Trim([temptbl4].[LN2]))-4) WHERE Right(Trim([temptbl4].[LN2]),4) = ' III';", dbSeeChanges: Call BriefDelay
            ' Company only.
            TILLDataBase.Execute "UPDATE temptbl4 SET [temptbl4].[LetterSalutation] = 'Friend' " & _
                "WHERE (Len(Trim([temptbl4].[LetterSalutation])) <= 0 OR Trim([temptbl4].[LetterSalutation]) Is Null) AND " & _
                "(Len(Trim([temptbl4].[LN1])) <= 0 OR Trim([temptbl4].[LN1]) Is Null) AND " & _
                "(Len(Trim([temptbl4].[LN2])) <= 0 OR Trim([temptbl4].[LN2]) Is Null) AND " & _
                "(Len(Trim([temptbl4].[CO1])) >0 OR Trim([temptbl4].[CO1]) Is Not Null);", dbSeeChanges: Call BriefDelay
            ' Single name only.
            TILLDataBase.Execute "UPDATE temptbl4 SET temptbl4.LetterSalutation = Trim([temptbl4]![SAL1]) & ' ' & Trim([temptbl4]![LN1]) " & _
                "WHERE (Len(Trim([temptbl4].[LetterSalutation])) <= 0 OR Trim([temptbl4].[LetterSalutation]) Is Null) AND " & _
                "(Len(Trim([temptbl4].[LN2])) <= 0 OR Trim([temptbl4].[LN2]) Is Null) AND " & _
                "(Len(Trim([temptbl4].[LN1])) >  0 OR Trim([temptbl4].[LN1]) Is Not Null);", dbSeeChanges: Call BriefDelay
            ' Different last names.
            TILLDataBase.Execute "UPDATE temptbl4 SET [temptbl4].[LetterSalutation] = Trim([temptbl4].[SAL1]) & ' ' & Trim([temptbl4].[LN1]) & ' and ' & Trim([temptbl4].[SAL2]) & ' ' & Trim(Trim([temptbl4].[LN2])) " & _
                "WHERE (Len(Trim([temptbl4].[LetterSalutation])) <= 0 OR Trim([temptbl4].[LetterSalutation]) Is Null) AND " & _
                "(Len(Trim([temptbl4].[LN1])) > 0 OR Trim([temptbl4].[LN1]) Is Not Null) AND " & _
                "(Len(Trim([temptbl4].[LN2])) > 0 OR Trim([temptbl4].[LN2]) Is Not Null) AND " & _
                "Trim([temptbl4].[LN1]) <> Trim([temptbl4].[LN2]);", dbSeeChanges: Call BriefDelay
            ' Mr and Ms/Mrs. Same last name.
            TILLDataBase.Execute "UPDATE temptbl4 SET temptbl4.LetterSalutation = Trim([temptbl4].[SAL1]) & ' and ' & Trim([temptbl4].[SAL2]) & ' ' & Trim([temptbl4].[LN1]) " & _
                "WHERE (Len(Trim([temptbl4].[LetterSalutation])) <= 0 OR Trim([temptbl4].[LetterSalutation]) Is Null) AND " & _
                "Trim([temptbl4].[LN1])  = Trim([temptbl4].[LN2]) AND " & _
                "(Trim([temptbl4].[SAL1]) = 'Mr' OR Trim([temptbl4].[SAL1]) = 'Dr') AND (Trim([temptbl4].[SAL2]) ='Ms' OR Trim([temptbl4].[SAL2]) ='Mrs');", dbSeeChanges: Call BriefDelay
            ' Ms/Mrs and Mr.  Same last name.
            TILLDataBase.Execute "UPDATE temptbl4 SET temptbl4.LetterSalutation = Trim([temptbl4]![SAL2]) & ' and ' & Trim([temptbl4]![SAL1]) & ' ' & Trim([temptbl4]![LN1]) " & _
                "WHERE (Len(Trim([temptbl4].[LetterSalutation])) <= 0 OR Trim([temptbl4].[LetterSalutation]) Is Null) AND " & _
                "Trim([temptbl4].[LN1])  = Trim([temptbl4].[LN2]) AND " & _
                "(Trim([temptbl4].[SAL2]) = 'Mr' OR Trim([temptbl4].[SAL2]) = 'Dr') AND (Trim([temptbl4].[SAL1]) ='Ms' OR Trim([temptbl4].[SAL1]) ='Mrs');", dbSeeChanges: Call BriefDelay
            ' Two Mr's or Two Ms's.
            TILLDataBase.Execute "UPDATE temptbl4 SET temptbl4.LetterSalutation = Trim([temptbl4].[FN1]) & ' and ' & Trim([temptbl4].[FN2]) " & _
                "WHERE (Len(Trim([temptbl4].[LetterSalutation])) <= 0 OR Trim([temptbl4].[LetterSalutation]) Is Null) AND " & _
                "Trim([temptbl4].[LN1]) = Trim([temptbl4].[LN2]) AND " & _
                "((Trim([temptbl4].[SAL1]) = 'Mr' AND Trim([temptbl4].[SAL2]) = 'Mr') OR (Trim([temptbl4].[SAL1]) = 'Ms' AND Trim([temptbl4].[SAL2]) = 'Ms'));", dbSeeChanges: Call BriefDelay
            TILLDataBase.Execute "UPDATE temptbl4 SET temptbl4.LetterSalutation = Trim([temptbl4].[SAL1]) & ' ' & Trim([temptbl4].[LN1]) & ' and ' & Trim([temptbl4].[SAL2]) & ' ' & Trim([temptbl4].[LN2]) " & _
                "WHERE (Len(Trim([temptbl4].[LetterSalutation])) <= 0 OR Trim([temptbl4].[LetterSalutation]) Is Null) AND " & _
                "Trim([temptbl4].[LN1]) <> Trim([temptbl4].[LN2]) AND " & _
                "((Trim([temptbl4].[SAL1]) = 'Mr' AND Trim([temptbl4].[SAL2]) = 'Mr') OR (Trim([temptbl4].[SAL1]) = 'Ms' AND Trim([temptbl4].[SAL2]) = 'Ms'));", dbSeeChanges: Call BriefDelay
        Case "Export2"
            TILLDataBase.Execute "UPDATE temptbl SET temptbl.LN1 = Left(Trim([temptbl].[LN1]),Len(Trim([temptbl].[LN1]))-3) " & _
                "WHERE Right(Trim([temptbl].[LN1]),3) = ' Jr' OR " & _
                      "Right(Trim([temptbl].[LN1]),3) = ' Sr' OR " & _
                      "Right(Trim([temptbl].[LN1]),3) = ' II' OR " & _
                      "Right(Trim([temptbl].[LN1]),3) = ' IV';", dbSeeChanges: Call BriefDelay
            TILLDataBase.Execute "UPDATE temptbl SET temptbl.LN1 = Left(Trim([temptbl].[LN1]),Len(Trim([temptbl].[LN1]))-4) WHERE Right(Trim([temptbl].[LN1]),4) = ' III';", dbSeeChanges: Call BriefDelay
            TILLDataBase.Execute "UPDATE temptbl SET temptbl.LN2 = Left(Trim([temptbl].[LN2]),Len(Trim([temptbl].[LN2]))-3) " & _
                "WHERE Right(Trim([temptbl].[LN2]),3) = ' Jr' OR " & _
                      "Right(Trim([temptbl].[LN2]),3) = ' Sr' OR " & _
                      "Right(Trim([temptbl].[LN1]),3) = ' II' OR " & _
                      "Right(Trim([temptbl].[LN1]),3) = ' IV';", dbSeeChanges: Call BriefDelay
            TILLDataBase.Execute "UPDATE temptbl SET temptbl.LN2 = Left(Trim([temptbl].[LN2]),Len(Trim([temptbl].[LN2]))-4) WHERE Right(Trim([temptbl].[LN2]),4) = ' III';", dbSeeChanges: Call BriefDelay
            ' Company only.
            TILLDataBase.Execute "UPDATE temptbl SET [temptbl].[LetterSalutation] = 'Friend' " & _
                "WHERE (Len(Trim([temptbl].[LetterSalutation])) <= 0 OR Trim([temptbl].[LetterSalutation]) Is Null) AND " & _
                "(Len(Trim([temptbl].[LN1])) <= 0 OR Trim([temptbl].[LN1]) Is Null) AND " & _
                "(Len(Trim([temptbl].[LN2])) <= 0 OR Trim([temptbl].[LN2]) Is Null) AND " & _
                "(Len(Trim([temptbl].[CO1])) >0 OR Trim([temptbl].[CO1]) Is Not Null);", dbSeeChanges: Call BriefDelay
            ' Single name only.
            TILLDataBase.Execute "UPDATE temptbl SET temptbl.LetterSalutation = Trim([temptbl]![SAL1]) & ' ' & Trim([temptbl]![LN1]) " & _
                "WHERE (Len(Trim([temptbl].[LetterSalutation])) <= 0 OR Trim([temptbl].[LetterSalutation]) Is Null) AND " & _
                "(Len(Trim([temptbl].[LN2])) <= 0 OR Trim([temptbl].[LN2]) Is Null) AND " & _
                "(Len(Trim([temptbl].[LN1])) >  0 OR Trim([temptbl].[LN1]) Is Not Null);", dbSeeChanges: Call BriefDelay
            ' Different last names.
            TILLDataBase.Execute "UPDATE temptbl SET [temptbl].[LetterSalutation] = Trim([temptbl].[SAL1]) & ' ' & Trim([temptbl].[LN1]) & ' and ' & Trim([temptbl].[SAL2]) & ' ' & Trim(Trim([temptbl].[LN2])) " & _
                "WHERE (Len(Trim([temptbl].[LetterSalutation])) <= 0 OR Trim([temptbl].[LetterSalutation]) Is Null) AND " & _
                "(Len(Trim([temptbl].[LN1])) > 0 OR Trim([temptbl].[LN1]) Is Not Null) AND " & _
                "(Len(Trim([temptbl].[LN2])) > 0 OR Trim([temptbl].[LN2]) Is Not Null) AND " & _
                "Trim([temptbl].[LN1]) <> Trim([temptbl].[LN2]);", dbSeeChanges: Call BriefDelay
            ' Mr and Ms/Mrs. Same last name.
            TILLDataBase.Execute "UPDATE temptbl SET temptbl.LetterSalutation = Trim([temptbl].[SAL1]) & ' and ' & Trim([temptbl].[SAL2]) & ' ' & Trim([temptbl].[LN1]) " & _
                "WHERE (Len(Trim([temptbl].[LetterSalutation])) <= 0 OR Trim([temptbl].[LetterSalutation]) Is Null) AND " & _
                "Trim([temptbl].[LN1])  = Trim([temptbl].[LN2]) AND " & _
                "Trim([temptbl].[SAL1]) = 'Mr' AND (Trim([temptbl].[SAL2]) ='Ms' OR Trim([temptbl].[SAL2]) ='Mrs');", dbSeeChanges: Call BriefDelay
            ' Ms/Mrs and Mr.  Same last name.
            TILLDataBase.Execute "UPDATE temptbl SET temptbl.LetterSalutation = Trim([temptbl]![SAL2]) & ' and ' & Trim([temptbl]![SAL1]) & ' ' & Trim([temptbl]![LN1]) " & _
                "WHERE (Len(Trim([temptbl].[LetterSalutation])) <= 0 OR Trim([temptbl].[LetterSalutation]) Is Null) AND " & _
                "Trim([temptbl].[LN1])  = Trim([temptbl].[LN2]) AND " & _
                "Trim([temptbl].[SAL2]) = 'Mr' AND (Trim([temptbl].[SAL1]) ='Ms' OR Trim([temptbl].[SAL1]) ='Mrs');", dbSeeChanges: Call BriefDelay
            ' Two Mr's or Two Ms's.
            TILLDataBase.Execute "UPDATE temptbl SET temptbl.LetterSalutation = Trim([temptbl].[FN1]) & ' and ' & Trim([temptbl].[FN2]) " & _
                "WHERE (Len(Trim([temptbl].[LetterSalutation])) <= 0 OR Trim([temptbl].[LetterSalutation]) Is Null) AND " & _
                "Trim([temptbl].[LN1]) = Trim([temptbl].[LN2]) AND " & _
                "((Trim([temptbl].[SAL1]) = 'Mr' AND Trim([temptbl].[SAL2]) = 'Mr') OR (Trim([temptbl].[SAL1]) = 'Ms' AND Trim([temptbl].[SAL2]) = 'Ms'));", dbSeeChanges: Call BriefDelay
            TILLDataBase.Execute "UPDATE temptbl SET temptbl.LetterSalutation = Trim([temptbl].[SAL1]) & ' ' & Trim([temptbl].[LN1]) & ' and ' & Trim([temptbl].[SAL2]) & ' ' & Trim([temptbl].[LN2]) " & _
                "WHERE (Len(Trim([temptbl].[LetterSalutation])) <= 0 OR Trim([temptbl].[LetterSalutation]) Is Null) AND " & _
                "Trim([temptbl].[LN1]) <> Trim([temptbl].[LN2]) AND " & _
                "((Trim([temptbl].[SAL1]) = 'Mr' AND Trim([temptbl].[SAL2]) = 'Mr') OR (Trim([temptbl].[SAL1]) = 'Ms' AND Trim([temptbl].[SAL2]) = 'Ms'));", dbSeeChanges: Call BriefDelay
    End Select
    DonorAppealFixSalutations = True
    Exit Function
ShowMeError:
    MsgBox "Error # " & Str(Err.Number) & " was generated by " & Me.Name & Chr(13) & Err.Description, vbOKOnly, "Error", Err.HelpFile, Err.HelpContext
    DonorAppealFixSalutations = False
End Function
