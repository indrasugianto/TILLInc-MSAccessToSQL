' Module Name: AddressValidation
' Module Type: Standard Module
' Lines of Code: 100
' Extracted: 1/29/2026 4:12:28 PM

Option Compare Database
Option Explicit

Public Function ValidateAddress(XA As Object, XC As Object, XS As Object, XZ As Object, XV As Object, HF As Form, Optional XD As Variant) As Boolean
On Error GoTo ShowMeError
    Dim strUrl As String    ' Our URL which will include the authentication info
    Dim strReq As String    ' The body of the POST request
    Dim xmlHttp As New MSXML2.XMLHTTP60, xmlDoc As MSXML2.DOMDocument60
    Dim candidates As MSXML2.IXMLDOMNode, candidate As MSXML2.IXMLDOMNode, components As MSXML2.IXMLDOMNode, metadata As MSXML2.IXMLDOMNode, analysis As MSXML2.IXMLDOMNode
    Dim AddressToCheck As Variant, CityToCheck As Variant, StateToCheck As Variant, ZIPToCheck As Variant, Validated As Boolean, MatchCode As Variant, Footnotes As Variant
    Dim candidate_count As Long, SQLCommand As String, Start, Finish

' This URL will execute the search request and return the resulting matches to the search in XML.
    strUrl = "https://api.smartystreets.com/street-address?auth-id=fb88500b-8f44-0321-a36c-723fea139ad7&auth-token=mp0xh7qlSywY362eewOK"
    AddressToCheck = XA.Value: CityToCheck = XC.Value: StateToCheck = XS.Value
    If Len(XZ.Value) = 6 Then ZIPToCheck = Left(XZ.Value, 5) Else ZIPToCheck = XZ.Value
' Body of the POST request
    strReq = "<?xml version=""1.0"" encoding=""utf-8""?>" & "<request>" & "<address>" & _
                "<street>" & AddressToCheck & "</street>" & "<city>" & CityToCheck & "</city>" & _
                "<state>" & StateToCheck & "</state>" & "<zipcode>" & ZIPToCheck & "</zipcode>" & _
                "<candidates>5</candidates>" & "</address>" & "</request>"
    With xmlHttp
        .Open "POST", strUrl, False                     ' Prepare POST request
        .setRequestHeader "Content-Type", "text/xml"    ' Sending XML ...
        .setRequestHeader "Accept", "text/xml"          ' ... expect XML in return.
        .send strReq                                    ' Send request body
    End With
    Call BriefDelay(2)
' The request has been saved into xmlHttp.responseText and is now ready to be parsed. Remember that fields in our XML response may change or be added to later, so make sure your method of parsing accepts that.
    Set xmlDoc = New MSXML2.DOMDocument60
    If Not xmlDoc.loadXML(xmlHttp.responseText) Then
        Err.Raise xmlDoc.parseError.errorCode, , xmlDoc.parseError.reason
        MsgBox "XML Error processing address.  Error #" & xmlDoc.parseError.errorCode & ".  Reason = " & xmlDoc.parseError.reason, vbOKOnly, "Error!"
        Exit Function
    End If
' According to the schema (http://smartystreets.com/kb/liveaddress-api/parsing-the-response#xml), <candidates> is a top-level node with each <candidate> below it. Let's obtain each one.
    Set candidates = xmlDoc.documentElement
    candidate_count = 0
' Get a count of all the search results.
    For Each candidate In candidates.childNodes
        candidate_count = candidate_count + 1
    Next
    Set candidates = xmlDoc.documentElement
    Select Case candidate_count
        Case 0 ' Bad address cannot be corrected.  Try again.
            HF.SetFocus
            MsgBox "The address supplied does not match a valid address in the USPS database.  Please correct this.", vbOKOnly, "Warning"
            XA.BackColor = RGB(255, 0, 0): XC.BackColor = RGB(255, 0, 0): XS.BackColor = RGB(255, 0, 0): XZ.BackColor = RGB(255, 0, 0)
            ValidateAddress = False
            Exit Function
        Case 1 ' Only one candidate address...use it and return.
            For Each candidate In candidates.childNodes
                Set analysis = candidate.selectSingleNode("analysis")
                XA.Value = candidate.selectSingleNode("delivery_line_1").nodeTypedValue
                Set components = candidate.selectSingleNode("components")
                If Not components Is Nothing Then
                    XC.Value = components.selectSingleNode("city_name").nodeTypedValue
                    XS.Value = components.selectSingleNode("state_abbreviation").nodeTypedValue
                    XZ.Value = components.selectSingleNode("zipcode").nodeTypedValue & "-" & components.selectSingleNode("plus4_code").nodeTypedValue
                End If
                Set metadata = candidate.selectSingleNode("metadata")
                If Not metadata Is Nothing Then If Not IsMissing(XD) Then XD.Value = 0
                XA.BackColor = RGB(0, 255, 0): XC.BackColor = RGB(0, 255, 0): XS.BackColor = RGB(0, 255, 0): XZ.BackColor = RGB(0, 255, 0)
                XV.Value = True
            Next
            ValidateAddress = False
            Exit Function
        Case Else ' Multiple candidate addresses...post them and allow the user to select.
            If IsTableQuery("temptbl") Then TILLDataBase.Execute "DROP TABLE temptbl", dbSeeChanges: Call BriefDelay
            TILLDataBase.Execute "CREATE TABLE temptbl (Selected BIT, CandidateAddress CHAR(50), CandidateCity CHAR(25), CandidateState CHAR(2), CandidateZIP CHAR(10), CandidateCongressionalDistrict INTEGER, MatchCode CHAR(1), Footnotes CHAR(30));", dbSeeChanges: Call BriefDelay
            Start = Timer
            For Each candidate In candidates.childNodes
                AddressToCheck = candidate.selectSingleNode("delivery_line_1").nodeTypedValue
                Set components = candidate.selectSingleNode("components")
                If Not components Is Nothing Then
                    CityToCheck = components.selectSingleNode("city_name").nodeTypedValue
                    StateToCheck = components.selectSingleNode("state_abbreviation").nodeTypedValue
                    ZIPToCheck = components.selectSingleNode("zipcode").nodeTypedValue & "-" & components.selectSingleNode("plus4_code").nodeTypedValue
                End If
                Set metadata = candidate.selectSingleNode("metadata")
                If Not metadata Is Nothing Then If Not IsMissing(XD) Then XD.Value = 0
                Set analysis = candidate.selectSingleNode("analysis")
                If Not analysis Is Nothing Then
                    MatchCode = analysis.selectSingleNode("dpv_match_code").nodeTypedValue
                    Footnotes = analysis.selectSingleNode("dpv_footnotes").nodeTypedValue
                End If
                TILLDataBase.Execute "INSERT INTO temptbl ( CandidateAddress, CandidateAddress, CandidateState, CandidateZIP, CandidateCongressionalDistrict, MatchCode, Footnotes )" & _
                    vbCrLf & "SELECT """ & AddressToCheck & """ AS CandidateAddress, """ & CityToCheck & """ AS CandidateAddress, """ & StateToCheck & _
                    """ AS CandidateState, """ & ZIPToCheck & """ AS CandidateZIP, " & XD & " AS CandidateCongressionalDistrict, """ & MatchCode & """ AS MatchCode, """ & _
                    Footnotes & """ AS Footnotes;", dbSeeChanges: Call BriefDelay
            Next
            ValidateAddress = True
            HF.SetFocus
    End Select
    Exit Function
ShowMeError:
    Err.Source = "PublicSubroutines" & "(Line #" & Str(Err.Erl) & ")": TILLDBErrorMessage = "Error # " & Str(Err.Number) & " was generated by " & Err.Source & Chr(13) & Err.Description
    MsgBox TILLDBErrorMessage, vbOKOnly, "Error", Err.HelpFile, Err.HelpContext
End Function
