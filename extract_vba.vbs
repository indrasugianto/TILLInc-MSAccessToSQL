' VBScript to extract VBA code from MS Access database
' This script must be run with MS Access installed
' Usage: cscript extract_vba.vbs

Option Explicit

Dim objAccess, objFSO, objDB
Dim strDBPath, strOutputDir, strVBADir, strReportPath
Dim intModuleCount, strReport

' Configuration
strDBPath = "c:\GitHub\TILLInc-MSAccessToSQL\msaccess\TILLDB_V9.14_20260128 - WEB.accdb"
strOutputDir = "c:\GitHub\TILLInc-MSAccessToSQL\extracted"
strVBADir = strOutputDir & "\vba"
strReportPath = strOutputDir & "\reports\vba_extraction_report.txt"

' Create FileSystemObject
Set objFSO = CreateObject("Scripting.FileSystemObject")

' Create output directory
If Not objFSO.FolderExists(strVBADir) Then
    objFSO.CreateFolder(strVBADir)
End If

WScript.Echo "==============================================="
WScript.Echo "VBA Code Extraction Script"
WScript.Echo "==============================================="
WScript.Echo "Database: " & strDBPath
WScript.Echo "Output: " & strVBADir
WScript.Echo ""

On Error Resume Next

' Create Access Application
Set objAccess = CreateObject("Access.Application")

If Err.Number <> 0 Then
    WScript.Echo "ERROR: Could not create Access.Application object"
    WScript.Echo "Error: " & Err.Description
    WScript.Echo ""
    WScript.Echo "MS Access must be installed to extract VBA code."
    WScript.Quit 1
End If

Err.Clear

' Open the database
WScript.Echo "Opening database..."
objAccess.OpenCurrentDatabase strDBPath, False

If Err.Number <> 0 Then
    WScript.Echo "ERROR: Could not open database"
    WScript.Echo "Error: " & Err.Description
    objAccess.Quit
    WScript.Quit 1
End If

WScript.Echo "Database opened successfully"
WScript.Echo ""

' Initialize counter
intModuleCount = 0
strReport = "VBA Code Extraction Report" & vbCrLf
strReport = strReport & "=============================" & vbCrLf & vbCrLf
strReport = strReport & "Database: " & strDBPath & vbCrLf
strReport = strReport & "Date: " & Now() & vbCrLf & vbCrLf

Err.Clear

' Extract VBA Code
WScript.Echo "Extracting VBA modules..."
WScript.Echo "--------------------------------------------"

On Error Resume Next

Dim vbProj, vbComp, codeModule
Dim i, moduleName, moduleType, lineCount, code
Dim objFile, safeFileName

Set vbProj = objAccess.VBE.VBProjects(1)

If Err.Number <> 0 Then
    WScript.Echo "ERROR: Could not access VBA project"
    WScript.Echo "Error: " & Err.Description
    WScript.Echo ""
    WScript.Echo "The database may be:"
    WScript.Echo "  - Password protected"
    WScript.Echo "  - Have VBA project protection enabled"
    WScript.Echo "  - Not have macro security configured properly"
    WScript.Echo ""
    strReport = strReport & "ERROR: Could not access VBA project" & vbCrLf
    strReport = strReport & "Error: " & Err.Description & vbCrLf
Else
    WScript.Echo "VBA Project: " & vbProj.Name
    WScript.Echo "Total Components: " & vbProj.VBComponents.Count
    WScript.Echo ""
    
    For i = 1 To vbProj.VBComponents.Count
        Err.Clear
        Set vbComp = vbProj.VBComponents(i)
        
        If Err.Number = 0 Then
            moduleName = vbComp.Name
            moduleType = vbComp.Type
            
            ' Get module type name
            Dim moduleTypeName
            Select Case moduleType
                Case 1: moduleTypeName = "Standard Module"
                Case 2: moduleTypeName = "Class Module"
                Case 3: moduleTypeName = "Form Module"
                Case 100: moduleTypeName = "Document Module"
                Case Else: moduleTypeName = "Unknown (" & moduleType & ")"
            End Select
            
            Set codeModule = vbComp.CodeModule
            lineCount = codeModule.CountOfLines
            
            If lineCount > 0 Then
                intModuleCount = intModuleCount + 1
                
                ' Get code
                Err.Clear
                code = codeModule.Lines(1, lineCount)
                
                If Err.Number = 0 Then
                    ' Sanitize filename
                    safeFileName = moduleName
                    safeFileName = Replace(safeFileName, "\", "_")
                    safeFileName = Replace(safeFileName, "/", "_")
                    safeFileName = Replace(safeFileName, ":", "_")
                    safeFileName = Replace(safeFileName, "*", "_")
                    safeFileName = Replace(safeFileName, "?", "_")
                    safeFileName = Replace(safeFileName, """", "_")
                    safeFileName = Replace(safeFileName, "<", "_")
                    safeFileName = Replace(safeFileName, ">", "_")
                    safeFileName = Replace(safeFileName, "|", "_")
                    
                    ' Write to file
                    Set objFile = objFSO.CreateTextFile(strVBADir & "\" & safeFileName & ".vba", True)
                    objFile.WriteLine "' Module Name: " & moduleName
                    objFile.WriteLine "' Module Type: " & moduleTypeName
                    objFile.WriteLine "' Lines of Code: " & lineCount
                    objFile.WriteLine "' Extracted: " & Now()
                    objFile.WriteLine ""
                    objFile.Write code
                    objFile.Close
                    
                    WScript.Echo "  [+] " & moduleName & " (" & moduleTypeName & ") - " & lineCount & " lines"
                    
                    strReport = strReport & moduleName & vbTab & moduleTypeName & vbTab & lineCount & " lines" & vbCrLf
                Else
                    WScript.Echo "  [-] " & moduleName & " - Error reading code: " & Err.Description
                    strReport = strReport & moduleName & vbTab & "ERROR" & vbTab & Err.Description & vbCrLf
                End If
            End If
        End If
    Next
End If

WScript.Echo ""
WScript.Echo "--------------------------------------------"
WScript.Echo "Total VBA Modules Extracted: " & intModuleCount
WScript.Echo "==============================================="

strReport = strReport & vbCrLf & "Total Modules Extracted: " & intModuleCount & vbCrLf

' Write report
Set objFile = objFSO.CreateTextFile(strReportPath, True)
objFile.Write strReport
objFile.Close

WScript.Echo ""
WScript.Echo "Report saved to: " & strReportPath
WScript.Echo "VBA files saved to: " & strVBADir

' Cleanup
objAccess.CloseCurrentDatabase
objAccess.Quit

Set vbComp = Nothing
Set vbProj = Nothing
Set objAccess = Nothing
Set objFSO = Nothing

WScript.Echo ""
WScript.Echo "Done!"
WScript.Quit 0
