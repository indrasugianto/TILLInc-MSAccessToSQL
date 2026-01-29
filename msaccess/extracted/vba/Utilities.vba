' Module Name: Utilities
' Module Type: Standard Module
' Lines of Code: 232
' Extracted: 1/29/2026 4:12:28 PM

Option Compare Database
Option Explicit

Public Sub BriefDelay(Optional Secs As Integer = 1)
    Dim S As Variant
    
    S = Timer
    Do While Timer < S + Secs: DoEvents: Loop
End Sub

Public Function CalcAge(B As Variant) As Integer
    Dim A As Variant
    
    If IsNull(B) Then
        CalcAge = 0
        Exit Function
    End If
    A = DateDiff("yyyy", B, Now)
    If Date < DateSerial(Year(Now), Month(B), Day(B)) Then A = A - 1
    CalcAge = CInt(A)
End Function

Public Function Highlight(Field As Object, HLOn As Boolean)
    If HLOn Then
        Highlight = True
        RememberPreviousBackColor = Field.BackColor
        Field.BackColor = RGB(164, 213, 226)
    Else
        Highlight = False
        Field.BackColor = RememberPreviousBackColor
    End If
End Function

Public Function IsFileOpen(FileName As String)
On Error Resume Next
    Dim iFilenum As Long, iErr As Long
     
    iFilenum = FreeFile()
    Open FileName For Input Lock Read As #iFilenum
    Close iFilenum
    iErr = Err
    
On Error GoTo 0
    Select Case iErr
        Case 0, 53: IsFileOpen = False
        Case 70:    IsFileOpen = True
        Case Else:  Error iErr
    End Select
End Function

Public Function IsObjectOpen(strName As String, Optional intObjectType As Integer = acForm) As Boolean
' intObjectType can be: acTable (value 0), acQuery (value 1), acForm (value 2) Default, acReport (value 3), acMacro (value 4), acModule (value 5)
' Returns True if strName is open, False otherwise.
On Error Resume Next
    IsObjectOpen = (SysCmd(acSysCmdGetObjectState, intObjectType, strName) <> 0)
    If Err <> 0 Then IsObjectOpen = False
End Function

Public Function IsTableQuery(TName As String) As Boolean
On Error Resume Next
    Dim D As Database, T As String

    IsTableQuery = False
    Set D = CurrentDb()
' See if the name is in the Tables collection.
    T = D.TableDefs(TName).Name
    If Err <> NAME_NOT_IN_COLLECTION Then
        IsTableQuery = True
        D.Close
        Exit Function
    End If
' Reset the error variable.
    Err = 0
    ' See if the name is in the Queries collection.
    T = D.QueryDefs(TName$).Name
    If Err <> NAME_NOT_IN_COLLECTION Then
        IsTableQuery = True
        D.Close
        Exit Function
    End If
End Function

Public Sub LoopUntilClosed(strName As String, ObjectType As Integer)
    Do While IsObjectOpen(strName, ObjectType)
        DoEvents
    Loop
End Sub

Public Function ValidDate(DateStr As String)
    If Not IsDate(DateStr) Then
        MsgBox "Not a valid date.  Dates must be entered as 'MM/DD/YYYY' including all leading zeroes.", vbOKOnly, "Error!"
        DateStr = Null: ValidDate = False
    Else
        DateStr = Format(DateStr, "mm/dd/yyyy")
        ValidDate = True
    End If
End Function

Public Function ValidateDateString(DateString As Variant)
    Dim EMonth As String, EDay As String, EYear As String, Pointer As Integer

    ValidateDateString = True
    
    If IsDate(DateString) Or IsNull(DateString) Then
        ValidateDateString = True
    Else
        MsgBox "The date string you entered is not recognized as a legitimate date.", vbOKOnly, "ERROR!"
        DateString = Null
        ValidateDateString = False
    End If
    
    Exit Function
    
    EMonth = Left(DateString, 2): EDay = Mid(DateString, 4, 2): EYear = Right(DateString, 2)

    Select Case EMonth
        Case "01", "02", "03", "04", "05", "06", "07", "08", "09", "10", "11", "12"
            Select Case EMonth
                Case "01"
                    Select Case EDay
                        Case "01", "02", "03", "04", "05", "06", "07", "08", "09", "10", "11", "12", "13", "14", "15", "16", "17", "18", "19", "20", "21", "22", "23", "24", "25", "26", "27", "28", "29", "30", "31"
                            Exit Function
                        Case Else
                            MsgBox "Invalid day number.", vbOKOnly, "Error!": DateString = Null:  ValidateDateString = False: Exit Function
                    End Select
                Case "02"
                    Select Case EDay
                        Case "01", "02", "03", "04", "05", "06", "07", "08", "09", "10", "11", "12", "13", "14", "15", "16", "17", "18", "19", "20", "21", "22", "23", "24", "25", "26", "27", "28", "29"
                            Exit Function
                        Case Else
                            MsgBox "Invalid day number.", vbOKOnly, "Error!": DateString = Null:  ValidateDateString = False: Exit Function
                    End Select
                Case "03"
                    Select Case EDay
                        Case "01", "02", "03", "04", "05", "06", "07", "08", "09", "10", "11", "12", "13", "14", "15", "16", "17", "18", "19", "20", "21", "22", "23", "24", "25", "26", "27", "28", "29", "30", "31"
                            Exit Function
                        Case Else
                            MsgBox "Invalid day number.", vbOKOnly, "Error!": DateString = Null:  ValidateDateString = False: Exit Function
                    End Select
                Case "04"
                    Select Case EDay
                        Case "01", "02", "03", "04", "05", "06", "07", "08", "09", "10", "11", "12", "13", "14", "15", "16", "17", "18", "19", "20", "21", "22", "23", "24", "25", "26", "27", "28", "29", "30"
                            Exit Function
                        Case Else
                            MsgBox "Invalid day number.", vbOKOnly, "Error!": DateString = Null:  ValidateDateString = False: Exit Function
                    End Select
                Case "05"
                    Select Case EDay
                        Case "01", "02", "03", "04", "05", "06", "07", "08", "09", "10", "11", "12", "13", "14", "15", "16", "17", "18", "19", "20", "21", "22", "23", "24", "25", "26", "27", "28", "29", "30", "31"
                            Exit Function
                        Case Else
                            MsgBox "Invalid day number.", vbOKOnly, "Error!": DateString = Null:  ValidateDateString = False: Exit Function
                    End Select
                Case "06"
                    Select Case EDay
                        Case "01", "02", "03", "04", "05", "06", "07", "08", "09", "10", "11", "12", "13", "14", "15", "16", "17", "18", "19", "20", "21", "22", "23", "24", "25", "26", "27", "28", "29", "30"
                            Exit Function
                        Case Else
                            MsgBox "Invalid day number.", vbOKOnly, "Error!": DateString = Null:  ValidateDateString = False: Exit Function
                    End Select
                Case "07"
                    Select Case EDay
                        Case "01", "02", "03", "04", "05", "06", "07", "08", "09", "10", "11", "12", "13", "14", "15", "16", "17", "18", "19", "20", "21", "22", "23", "24", "25", "26", "27", "28", "29", "30", "31"
                            Exit Function
                        Case Else
                            MsgBox "Invalid day number.", vbOKOnly, "Error!": DateString = Null:  ValidateDateString = False: Exit Function
                    End Select
                Case "08"
                    Select Case EDay
                        Case "01", "02", "03", "04", "05", "06", "07", "08", "09", "10", "11", "12", "13", "14", "15", "16", "17", "18", "19", "20", "21", "22", "23", "24", "25", "26", "27", "28", "29", "30", "31"
                            Exit Function
                        Case Else
                            MsgBox "Invalid day number.", vbOKOnly, "Error!": DateString = Null:  ValidateDateString = False: Exit Function
                    End Select
                Case "09"
                    Select Case EDay
                        Case "01", "02", "03", "04", "05", "06", "07", "08", "09", "10", "11", "12", "13", "14", "15", "16", "17", "18", "19", "20", "21", "22", "23", "24", "25", "26", "27", "28", "29", "30"
                            Exit Function
                        Case Else
                            MsgBox "Invalid day number.", vbOKOnly, "Error!": DateString = Null:  ValidateDateString = False: Exit Function
                    End Select
                Case "10"
                    Select Case EDay
                        Case "01", "02", "03", "04", "05", "06", "07", "08", "09", "10", "11", "12", "13", "14", "15", "16", "17", "18", "19", "20", "21", "22", "23", "24", "25", "26", "27", "28", "29", "30", "31"
                            Exit Function
                        Case Else
                            MsgBox "Invalid day number.", vbOKOnly, "Error!": DateString = Null:  ValidateDateString = False: Exit Function
                    End Select
                Case "11"
                    Select Case EDay
                        Case "01", "02", "03", "04", "05", "06", "07", "08", "09", "10", "11", "12", "13", "14", "15", "16", "17", "18", "19", "20", "21", "22", "23", "24", "25", "26", "27", "28", "29", "30"
                            Exit Function
                        Case Else
                            MsgBox "Invalid day number.", vbOKOnly, "Error!": DateString = Null:  ValidateDateString = False: Exit Function
                    End Select
                Case "12"
                    Select Case EDay
                        Case "01", "02", "03", "04", "05", "06", "07", "08", "09", "10", "11", "12", "13", "14", "15", "16", "17", "18", "19", "20", "21", "22", "23", "24", "25", "26", "27", "28", "29", "30", "31"
                            Exit Function
                        Case Else
                            MsgBox "Invalid day number.", vbOKOnly, "Error!": DateString = Null:  ValidateDateString = False: Exit Function
                    End Select
            End Select
        Case Else
            MsgBox "Invalid month number.", vbOKOnly, "Error!": DateString = Null:  ValidateDateString = False: Exit Function
    End Select
End Function

Public Function FieldIsEmpty(Field As Variant) As Boolean
    If IsNull(Field) Then
        FieldIsEmpty = True
    ElseIf Len(Field) <= 0 Then
        FieldIsEmpty = True
    Else
        FieldIsEmpty = False
    End If
End Function

Public Function ProgressMessages(Action As String, MessageToPost As String) As Boolean
    ProgressMessages = True
    
    Select Case Action
        Case "Open"
            DoCmd.OpenForm "frmProgressMessages"
            Form_frmProgressMessages.MEssages = MessageToPost & vbCrLf
        Case "Append"
            Form_frmProgressMessages.MEssages = Form_frmProgressMessages.MEssages & MessageToPost & vbCrLf
            Form_frmProgressMessages.MEssages.Requery
        Case "Close"
            DoCmd.Close acForm, "frmProgressMessages"
    End Select
End Function