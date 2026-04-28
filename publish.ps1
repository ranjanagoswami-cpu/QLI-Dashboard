# ============================================================
# QLI Security Dashboard - Publish to GitHub Pages
# Paths are RELATIVE to this script's location.
# Move the entire folder anywhere and it will still work.
# Run this every time you update the Excel file.
# ============================================================

$scriptDir  = if ($PSScriptRoot) { $PSScriptRoot } else { Split-Path -Parent $MyInvocation.MyCommand.Path }
$outputPath = Join-Path $scriptDir "data.js"
$configFile = Join-Path $scriptDir ".excel-path.txt"
$ghConfig   = Join-Path $scriptDir ".github-config.ps1"
$excelName  = "QLI_Project_Tracking_Template_AutoRAG_Override_Backlog_withBacklogSummary.xlsx"

Set-Location $scriptDir

Write-Host ""
Write-Host "============================================" -ForegroundColor Cyan
Write-Host "  QLI Security Dashboard - Publishing..." -ForegroundColor Cyan
Write-Host "============================================" -ForegroundColor Cyan
Write-Host "  Script location: $scriptDir" -ForegroundColor Gray
Write-Host ""

# ---- Load GitHub Pages config ----
$pagesUrl = $null
if (Test-Path $ghConfig) { . $ghConfig }

# ---- Check git ----
try { git --version 2>&1 | Out-Null } catch {
    Write-Host "[ERROR] Git not installed. Run setup-github-pages.ps1 first." -ForegroundColor Red
    Read-Host "Press Enter to exit"; exit 1
}
if (-not (Test-Path (Join-Path $scriptDir ".git"))) {
    Write-Host "[ERROR] Git repository not found." -ForegroundColor Red
    Write-Host "  Run setup-github-pages.ps1 first." -ForegroundColor Yellow
    Read-Host "Press Enter to exit"; exit 1
}

# ============================================================
# AUTO-DETECT EXCEL FILE
# ============================================================
function Find-ExcelFile {
    param($scriptDir, $excelName, $configFile)

    if (Test-Path $configFile) {
        $savedPath = (Get-Content $configFile -Raw).Trim()
        if (Test-Path $savedPath) { return $savedPath }
    }

    $p = Join-Path $scriptDir $excelName
    if (Test-Path $p) { return $p }

    $p = Join-Path (Split-Path $scriptDir -Parent) "Excel folder\$excelName"
    if (Test-Path $p) { return $p }

    $p = Join-Path (Split-Path $scriptDir -Parent) $excelName
    if (Test-Path $p) { return $p }

    $found = Get-ChildItem -Path (Split-Path $scriptDir -Parent) -Filter $excelName -Recurse -ErrorAction SilentlyContinue | Select-Object -First 1
    if ($found) { return $found.FullName }

    return $null
}

# ============================================================
# STEP 1: Read Excel -> Generate data.js
# ============================================================
Write-Host "[1/3] Reading Excel and generating data.js..." -ForegroundColor Yellow
Write-Host ""

$excelPath = Find-ExcelFile -scriptDir $scriptDir -excelName $excelName -configFile $configFile

if (-not $excelPath) {
    Write-Host "[!] Excel file not found automatically." -ForegroundColor Yellow
    Write-Host "    Expected: $excelName" -ForegroundColor Gray
    Write-Host "    Please enter the full path to your Excel file:" -ForegroundColor Cyan
    $excelPath = (Read-Host "    Path").Trim('"').Trim("'")
    if (-not (Test-Path $excelPath)) {
        Write-Host "[ERROR] File not found: $excelPath" -ForegroundColor Red
        Read-Host "Press Enter to exit"; exit 1
    }
}

$excelPath | Out-File -FilePath $configFile -Encoding UTF8 -NoNewline
Write-Host "    Excel:  $excelPath" -ForegroundColor Gray
Write-Host "    Output: $outputPath" -ForegroundColor Gray
Write-Host ""

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
        $rag     = $usedRange.Cells($r, 11).Text.Trim()
        $pctRaw  = $usedRange.Cells($r, 12).Text.Trim()
        $risk    = $usedRange.Cells($r, 14).Text.Trim()
        $next    = $usedRange.Cells($r, 15).Text.Trim()
        $updated = $usedRange.Cells($r, 16).Text.Trim()

        if ($id -eq "") { continue }

        $pct = 0
        if ($pctRaw -match "(\d+)") { $pct = [int]$matches[1] }
        if ($rag -eq "") { $rag = "Green" }

        $name  = $name  -replace "'", "\'"
        $owner = $owner -replace "'", "\'"
        $risk  = $risk  -replace "'", "\'"
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
// Auto-generated by publish.ps1
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
    Write-Host "    Tasks:         $($tasks.Count)" -ForegroundColor White
    Write-Host "    Backlog items: $($backlogItems.Count)" -ForegroundColor White
    Write-Host "    Last updated:  $lastUpdated" -ForegroundColor White
    Write-Host "    [OK] data.js generated." -ForegroundColor Green

    # ---- Also generate standalone dashboard-share.html ----
    $shareOutputPath = Join-Path $scriptDir "dashboard-share.html"
    $indexPath       = Join-Path $scriptDir "index.html"
    if (Test-Path $indexPath) {
        $htmlContent = Get-Content $indexPath -Raw -Encoding UTF8
        $inlineData  = "<script>`n$js`n</script>"
        $htmlContent = $htmlContent -replace '<script src="data\.js"[^>]*></script>', $inlineData
        $htmlContent | Out-File -FilePath $shareOutputPath -Encoding UTF8
        Write-Host "    [OK] dashboard-share.html generated (single file for direct sharing)." -ForegroundColor Green
    }

} catch {
    Write-Host "[ERROR] Failed to read Excel: $_" -ForegroundColor Red
    try { $workbook.Close($false) } catch {}
    $excel.Quit()
    Read-Host "Press Enter to exit"; exit 1
} finally {
    try { $workbook.Close($false) } catch {}
    $excel.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
}

# ============================================================
# STEP 2: Commit to Git
# ============================================================
Write-Host ""
Write-Host "[2/3] Committing changes to Git..." -ForegroundColor Yellow

$timestamp = Get-Date -Format "dd-MMM-yyyy HH:mm"
git add data.js 2>&1 | Out-Null

$gitStatus = git status --porcelain 2>&1
if ($gitStatus -eq "") {
    Write-Host "    No changes -- data.js is already up to date." -ForegroundColor Gray
} else {
    git commit -m "Dashboard update: $timestamp" 2>&1 | Out-Null
    Write-Host "    [OK] Committed: Dashboard update: $timestamp" -ForegroundColor Green
}

# ============================================================
# STEP 3: Push to GitHub Pages
# ============================================================
Write-Host ""
Write-Host "[3/3] Pushing to GitHub Pages..." -ForegroundColor Yellow

$pushResult = git push origin main 2>&1
if ($LASTEXITCODE -eq 0) {
    Write-Host "    [OK] Pushed successfully." -ForegroundColor Green
    Write-Host ""
    Write-Host "============================================" -ForegroundColor Green
    Write-Host "  PUBLISHED SUCCESSFULLY!" -ForegroundColor Green
    Write-Host "============================================" -ForegroundColor Green
    Write-Host ""
    if ($pagesUrl) {
        Write-Host "  Dashboard URL:" -ForegroundColor White
        Write-Host "  $pagesUrl" -ForegroundColor Cyan
        Write-Host ""
        Write-Host "  Share this link with your team." -ForegroundColor White
        Write-Host "  Updates are live in ~1-2 minutes." -ForegroundColor Gray
    } else {
        Write-Host "  Find your URL: GitHub repo Settings Pages" -ForegroundColor Yellow
    }
} else {
    Write-Host ""
    Write-Host "[ERROR] Push failed. Possible reasons:" -ForegroundColor Red
    Write-Host "  - Not authenticated with GitHub" -ForegroundColor Yellow
    Write-Host "  - Remote repo does not exist (create at github.com/new)" -ForegroundColor Yellow
    Write-Host "  - Run setup-github-pages.ps1 if not done yet" -ForegroundColor Yellow
    Write-Host ""
    Write-Host "  Push output: $pushResult" -ForegroundColor Gray
}

Write-Host ""
Read-Host "Press Enter to exit"
