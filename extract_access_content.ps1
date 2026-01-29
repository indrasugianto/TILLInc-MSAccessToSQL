# Script to extract queries and VBA code from MS Access database
param(
    [string]$DatabasePath = "c:\GitHub\TILLInc-MSAccessToSQL\msaccess\TILLDB_V9.14_20260128 - WEB.accdb",
    [string]$OutputDir = "c:\GitHub\TILLInc-MSAccessToSQL\extracted"
)

# Create output directories
$QueriesDir = Join-Path $OutputDir "queries"
$VBADir = Join-Path $OutputDir "vba"
$ReportsDir = Join-Path $OutputDir "reports"

New-Item -ItemType Directory -Force -Path $QueriesDir | Out-Null
New-Item -ItemType Directory -Force -Path $VBADir | Out-Null
New-Item -ItemType Directory -Force -Path $ReportsDir | Out-Null

Write-Host "Starting extraction from: $DatabasePath" -ForegroundColor Green

try {
    # Create Access Application object
    $Access = New-Object -ComObject Access.Application
    $Access.Visible = $false
    
    # Open the database
    Write-Host "Opening database..." -ForegroundColor Yellow
    $Access.OpenCurrentDatabase($DatabasePath, $false)
    
    $db = $Access.CurrentDb()
    
    # Extract Queries
    Write-Host "`nExtracting Queries..." -ForegroundColor Cyan
    $queryCount = 0
    $queryList = @()
    
    foreach ($qry in $db.QueryDefs) {
        $queryName = $qry.Name
        
        # Skip system queries (those starting with ~)
        if ($queryName -notlike "~*" -and $queryName -notlike "MSys*") {
            $queryCount++
            $sql = $qry.SQL
            $queryType = $qry.Type
            
            # Sanitize filename
            $safeFileName = $queryName -replace '[\\/:*?"<>|]', '_'
            $filePath = Join-Path $QueriesDir "$safeFileName.sql"
            
            # Create query info
            $queryInfo = @"
-- Query Name: $queryName
-- Query Type: $queryType
-- Extracted: $(Get-Date -Format "yyyy-MM-dd HH:mm:ss")

$sql
"@
            
            Set-Content -Path $filePath -Value $queryInfo -Encoding UTF8
            
            $queryList += [PSCustomObject]@{
                Name = $queryName
                Type = $queryType
                File = "$safeFileName.sql"
            }
            
            Write-Host "  Extracted: $queryName" -ForegroundColor Gray
        }
    }
    
    Write-Host "Total Queries Extracted: $queryCount" -ForegroundColor Green
    
    # Extract VBA Code
    Write-Host "`nExtracting VBA Code..." -ForegroundColor Cyan
    $vbaCount = 0
    $moduleList = @()
    
    try {
        $vbaProject = $Access.VBE.VBProjects(1)
        
        foreach ($component in $vbaProject.VBComponents) {
            $moduleName = $component.Name
            $moduleType = $component.Type
            
            # Get module type name
            $moduleTypeName = switch ($moduleType) {
                1 { "Standard Module" }
                2 { "Class Module" }
                3 { "Form Module" }
                100 { "Document Module" }
                default { "Unknown ($moduleType)" }
            }
            
            $vbaCount++
            
            # Get code
            if ($component.CodeModule.CountOfLines -gt 0) {
                $code = $component.CodeModule.Lines(1, $component.CodeModule.CountOfLines)
                
                # Sanitize filename
                $safeFileName = $moduleName -replace '[\\/:*?"<>|]', '_'
                $filePath = Join-Path $VBADir "$safeFileName.vba"
                
                # Create module info
                $moduleInfo = @"
' Module Name: $moduleName
' Module Type: $moduleTypeName
' Lines of Code: $($component.CodeModule.CountOfLines)
' Extracted: $(Get-Date -Format "yyyy-MM-dd HH:mm:ss")

$code
"@
                
                Set-Content -Path $filePath -Value $moduleInfo -Encoding UTF8
                
                $moduleList += [PSCustomObject]@{
                    Name = $moduleName
                    Type = $moduleTypeName
                    Lines = $component.CodeModule.CountOfLines
                    File = "$safeFileName.vba"
                }
                
                Write-Host "  Extracted: $moduleName ($moduleTypeName) - $($component.CodeModule.CountOfLines) lines" -ForegroundColor Gray
            }
        }
    } catch {
        Write-Host "  Warning: Could not access VBA project. The database may be password protected or VBA may not be accessible." -ForegroundColor Yellow
        Write-Host "  Error: $($_.Exception.Message)" -ForegroundColor Yellow
    }
    
    Write-Host "Total VBA Modules Extracted: $vbaCount" -ForegroundColor Green
    
    # Create summary report
    Write-Host "`nCreating summary report..." -ForegroundColor Cyan
    
    $summaryReport = @"
# MS Access Database Extraction Report
Database: $DatabasePath
Extraction Date: $(Get-Date -Format "yyyy-MM-dd HH:mm:ss")

## Connection Information
Server: tillsqlserver.database.windows.net
User: tillsqladmin
Database Type: Azure SQL Database

## Summary
- Total Queries Extracted: $queryCount
- Total VBA Modules Extracted: $vbaCount

## Queries Extracted
"@
    
    if ($queryList.Count -gt 0) {
        $summaryReport += "`n| Query Name | Type | Output File |`n"
        $summaryReport += "|------------|------|-------------|`n"
        foreach ($q in $queryList) {
            $summaryReport += "| $($q.Name) | $($q.Type) | $($q.File) |`n"
        }
    } else {
        $summaryReport += "`nNo queries found.`n"
    }
    
    $summaryReport += "`n## VBA Modules Extracted`n"
    
    if ($moduleList.Count -gt 0) {
        $summaryReport += "`n| Module Name | Type | Lines | Output File |`n"
        $summaryReport += "|-------------|------|-------|-------------|`n"
        foreach ($m in $moduleList) {
            $summaryReport += "| $($m.Name) | $($m.Type) | $($m.Lines) | $($m.File) |`n"
        }
    } else {
        $summaryReport += "`nNo VBA modules found or VBA not accessible.`n"
    }
    
    $reportPath = Join-Path $ReportsDir "extraction_summary.md"
    Set-Content -Path $reportPath -Value $summaryReport -Encoding UTF8
    
    Write-Host "`nSummary report saved to: $reportPath" -ForegroundColor Green
    
    # Close database
    $Access.CloseCurrentDatabase()
    $Access.Quit()
    
    # Release COM objects
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($db) | Out-Null
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($Access) | Out-Null
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
    
    Write-Host "`nExtraction completed successfully!" -ForegroundColor Green
    Write-Host "Output directory: $OutputDir" -ForegroundColor Green
    
} catch {
    Write-Host "`nError during extraction: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host $_.Exception.StackTrace -ForegroundColor Red
    
    # Cleanup
    if ($Access) {
        try {
            $Access.Quit()
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($Access) | Out-Null
        } catch {}
    }
    
    exit 1
}
