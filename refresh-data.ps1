# ============================================================
# QLI Security Dashboard - Data Refresh Script
# Paths are RELATIVE to this script's location.
# Move the entire folder anywhere and it will still work.
# ============================================================

# Always resolve paths relative to THIS script's folder
$scriptDir   = if ($PSScriptRoot) { $PSScriptRoot } else { Split-Path -Parent $MyInvocation.MyCommand.Path }
$outputPath  = Join-Path $scriptDir "data.js"
$configFile  = Join-Path $scriptDir ".excel-path.txt"
$excelName   = "QLI_Project_Tracking_Template_AutoRAG_Override_Backlog_withBacklogSummary.xlsx"

Write-Host ""
Write-Host "QLI Security Dashboard - Data Refresh" -ForegroundColor Cyan
Write-Host "Script location: $scriptDir" -ForegroundColor Gray
Write-Host ""

# ============================================================
# AUTO-DETECT EXCEL FILE LOCATION
# ============================================================
function Find-ExcelFile {
    param($scriptDir, $excelName, $configFile)

    # 1. Check saved path from last run
    if (Test-Path $configFile) {
        $savedPath = (Get-Content $configFile -Raw).Trim()
        if (Test-Path $savedPath) {
            return $savedPath
        }
    }

    # 2. Check same folder as script
    $p = Join-Path $scriptDir $excelName
    if (Test-Path $p) { return $p }

    # 3. Check sibling "Excel folder" (default project structure)
    $p = Join-Path (Split-Path $scriptDir -Parent) "Excel folder\$excelName"
    if (Test-Path $p) { return $p }

    # 4. Check parent folder directly
    $p = Join-Path (Split-Path $scriptDir -Parent) $excelName
    if (Test-Path $p) { return $p }

    # 5. Search one level up recursively (handles OneDrive sync variations)
    $found = Get-ChildItem -Path (Split-Path $scriptDir -Parent) -Filter $excelName -Recurse -ErrorAction SilentlyContinue | Select-Object -First 1
    if ($found) { return $found.FullName }

    return $null
}

$excelPath = Find-ExcelFile -scriptDir $scriptDir -excelName $excelName -configFile $configFile

if (-not $excelPath) {
    Write-Host "[!] Excel file not found automatically." -ForegroundColor Yellow
    Write-Host "    Expected file: $excelName" -ForegroundColor Gray
    Write-Host ""
    Write-Host "    Please enter the full path to your Excel file:" -ForegroundColor Cyan
    $excelPath = Read-Host "    Path"
    $excelPath = $excelPath.Trim('"').Trim("'")

    if (-not (Test-Path $excelPath)) {
        Write-Host "[ERROR] File not found at: $excelPath" -ForegroundColor Red
        Read-Host "Press Enter to exit"
        exit 1
    }
}

# Save the resolved path for next run
$excelPath | Out-File -FilePath $configFile -Encoding UTF8 -NoNewline
Write-Host "[OK] Excel file: $excelPath" -ForegroundColor Green
Write-Host "[OK] Output:     $outputPath" -ForegroundColor Green
Write-Host ""

# ============================================================
# READ EXCEL AND GENERATE data.js
# ============================================================
Write-Host "Opening Excel..." -ForegroundColor Cyan

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

try {
    $workbook = $excel.Workbooks.Open($excelPath)

    # ---- Task Tracker ----
    $sheet     = $workbook.Sheets["Task Tracker"]
    $usedRange = $sheet.UsedRange
    $rowCount  = $usedRange.Rows.Count
    $tasks     = @()

    for ($r = 2; $r -le $rowCount; $r++) {
        $id      = $usedRange.Cells($r, 1).Text.Trim()
        $name    = $usedRange.Cells($r, 2).Text.Trim()
        $cat     = $usedRange.Cells($r, 3).Text.Trim()
        $owner   = $usedRange.Cells($r, 4).Text.Trim()
        $rel     = $usedRange.Cells($r, 5).Text.Trim()
        $eta     = $usedRange.Cells($r, 6).Text.Trim()
        $revEta  = $usedRange.Cells($r, 7).Text.Trim()
        $status  = $usedRange.Cells($r, 8).Text.Trim()
        $ragOvr  = $usedRange.Cells($r, 9).Text.Trim()
        $ragMan  = $usedRange.Cells($r, 10).Text.Trim()
        $rag     = $usedRange.Cells($r, 11).Text.Trim()
        $pctRaw  = $usedRange.Cells($r, 12).Text.Trim()
        $dep     = $usedRange.Cells($r, 13).Text.Trim()
        $risk    = $usedRange.Cells($r, 14).Text.Trim()
        $next    = $usedRange.Cells($r, 15).Text.Trim()
        $updated = $usedRange.Cells($r, 16).Text.Trim()
        $notes   = $usedRange.Cells($r, 17).Text.Trim()

        if ($id -eq "") { continue }

        $pct = 0
        if ($pctRaw -match "(\d+)") { $pct = [int]$matches[1] }
        if ($rag -eq "") { $rag = "Green" }

        $name  = $name  -replace "'", "\'"
        $owner = $owner -replace "'", "\'"
        $risk  = $risk  -replace "'", "\'"
        $notes = $notes -replace "'", "\'"
        $next  = $next  -replace "'", "\'"

        $tasks += "  {id:'$id',name:'$name',cat:'$cat',owner:'$owner',rel:'$rel',eta:'$eta',revEta:'$revEta',status:'$status',rag:'$rag',pct:$pct,risk:'$risk',nextStep:'$next',updated:'$updated'}"
    }

    # ---- Backlog ----
    $sheet2     = $workbook.Sheets["Backlog"]
    $usedRange2 = $sheet2.UsedRange
    $rowCount2  = $usedRange2.Rows.Count
    $backlogItems = @()

    for ($r = 2; $r -le $rowCount2; $r++) {
        $id      = $usedRange2.Cells($r, 1).Text.Trim()
        $name    = $usedRange2.Cells($r, 2).Text.Trim()
        $cat     = $usedRange2.Cells($r, 3).Text.Trim()
        $owner   = $usedRange2.Cells($r, 4).Text.Trim()
        $scope   = $usedRange2.Cells($r, 5).Text.Trim()
        $risk    = $usedRange2.Cells($r, 14).Text.Trim()
        $updated = $usedRange2.Cells($r, 16).Text.Trim()

        if ($id -eq "") { continue }

        $name  = $name  -replace "'", "\'"
        $owner = $owner -replace "'", "\'"
        $risk  = $risk  -replace "'", "\'"

        $backlogItems += "  {id:'$id',name:'$name',cat:'$cat',owner:'$owner',scope:'$scope',risk:'$risk',updated:'$updated'}"
    }

    # ---- Release Summary ----
    $sheet3     = $workbook.Sheets["Release Summary"]
    $usedRange3 = $sheet3.UsedRange
    $rowCount3  = $usedRange3.Rows.Count
    $relSummary = @()

    for ($r = 2; $r -le $rowCount3; $r++) {
        $relName   = $usedRange3.Cells($r, 1).Text.Trim()
        $total     = $usedRange3.Cells($r, 2).Text.Trim()
        $done      = $usedRange3.Cells($r, 3).Text.Trim()
        $inprog    = $usedRange3.Cells($r, 4).Text.Trim()
        $blocked   = $usedRange3.Cells($r, 5).Text.Trim()
        $notstart  = $usedRange3.Cells($r, 6).Text.Trim()
        $inscope   = $usedRange3.Cells($r, 7).Text.Trim()
        $onhold    = $usedRange3.Cells($r, 8).Text.Trim()
        $avgPctRaw = $usedRange3.Cells($r, 9).Text.Trim()

        if ($relName -eq "") { continue }

        $avgPct = 0
        if ($avgPctRaw -match "(\d+)") { $avgPct = [int]$matches[1] }

        $relSummary += "  {rel:'$relName',total:$total,completed:$done,inProgress:$inprog,blocked:$blocked,notStarted:$notstart,inScoping:$inscope,onHold:$onhold,avgPct:$avgPct}"
    }

    $lastUpdated = Get-Date -Format "dd-MMM-yyyy HH:mm"

    $js = @"
// Auto-generated by refresh-data.ps1
// Last refreshed: $lastUpdated

var tasks = [
$($tasks -join ",`n")
];

var backlog = [
$($backlogItems -join ",`n")
];

var releaseSummary = [
$($relSummary -join ",`n")
];

var lastRefreshed = '$lastUpdated';
"@

    $js | Out-File -FilePath $outputPath -Encoding UTF8

    Write-Host "[OK] data.js written to: $outputPath" -ForegroundColor Green
    Write-Host "     Tasks:         $($tasks.Count)" -ForegroundColor White
    Write-Host "     Backlog items: $($backlogItems.Count)" -ForegroundColor White
    Write-Host "     Last updated:  $lastUpdated" -ForegroundColor White

    # ============================================================
    # GENERATE STANDALONE dashboard-share.html
    # A single self-contained file — share via email, Teams,
    # OneDrive link, SharePoint, etc. No data.js needed.
    # ============================================================
    $shareOutputPath = Join-Path $scriptDir "dashboard-share.html"
    $indexPath       = Join-Path $scriptDir "index.html"

    if (Test-Path $indexPath) {
        $htmlContent = Get-Content $indexPath -Raw -Encoding UTF8

        # Replace the external data.js script tag with inline data
        $inlineData = "<script>`n$js`n</script>"
        $htmlContent = $htmlContent -replace '<script src="data\.js"[^>]*></script>', $inlineData

        $htmlContent | Out-File -FilePath $shareOutputPath -Encoding UTF8
        Write-Host "[OK] Standalone file: $shareOutputPath" -ForegroundColor Green
        Write-Host "     Share this single file with your team — no other files needed." -ForegroundColor Cyan
    }

    Write-Host ""
    Write-Host "Open index.html in your browser to view the dashboard." -ForegroundColor Cyan
    Write-Host "Share dashboard-share.html as a single file (email, Teams, OneDrive)." -ForegroundColor Cyan

} catch {
    Write-Host "[ERROR] $_" -ForegroundColor Red
} finally {
    $workbook.Close($false)
    $excel.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
}

Write-Host ""
Read-Host "Press Enter to exit"
