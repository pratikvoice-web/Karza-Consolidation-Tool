<# :
@echo off
setlocal
color 0B
title Karza Master Consolidation Dashboard
echo ======================================================
echo   INITIALIZING KARZA DEEP-CONSOLIDATION ENGINE
echo ======================================================
set "currentDir=%~dp0"
powershell -noprofile -executionpolicy bypass -command "iex ([System.IO.File]::ReadAllText('%~f0'))"
echo.
echo Process Finished.
pause
exit /b
#>

# --- Master Logic Starts Here ---

$currentFolder = $env:currentDir
if (-not $currentFolder) { $currentFolder = Get-Location }

# 1. Full 2026 GST State Map
$stateMap = @{ "01"="J&K"; "02"="HP"; "03"="Punjab"; "04"="Chandigarh"; "05"="Uttarakhand"; "06"="Haryana"; "07"="Delhi"; "08"="Rajasthan"; "09"="UP"; "10"="Bihar"; "11"="Sikkim"; "12"="Arunachal"; "13"="Nagaland"; "14"="Manipur"; "15"="Mizoram"; "16"="Tripura"; "17"="Meghalaya"; "18"="Assam"; "19"="WB"; "20"="Jharkhand"; "21"="Odisha"; "22"="Chhattisgarh"; "23"="MP"; "24"="Gujarat"; "26"="DNHDD"; "27"="Maharashtra"; "29"="Karnataka"; "30"="Goa"; "31"="Lakshadweep"; "32"="Kerala"; "33"="TN"; "34"="Puducherry"; "35"="A&N Islands"; "36"="Telangana"; "37"="Andhra Pradesh"; "38"="Ladakh"; "97"="UN Bodies"; "99"="Foreign Entities" }

# --- Robust Dashboard UI Function ---
$script:firstBarInPhase = $true

function Show-Dashboard-Bar (${current}, ${total}, ${taskLabel}, ${phaseTitle}) {
    $percent = [int]((${current} / ${total}) * 100)
    $width = 40
    $done = [int]($percent * $width / 100)
    $left = $width - $done
    $bar = "█" * $done + "░" * $left
    
    # Overwrite previous two lines if this isn't the first call in the phase
    if (-not $script:firstBarInPhase) {
        $pos = $host.ui.RawUI.CursorPosition
        $pos.Y -= 2
        $host.ui.RawUI.CursorPosition = $pos
    }
    $script:firstBarInPhase = $false

    # Truncate label to prevent wrapping which breaks cursor logic
    if (${taskLabel}.Length -gt 60) { ${taskLabel} = ${taskLabel}.Substring(0, 57) + "..." }
    
    # Line 1: Progress Bar (Wrapped in braces to fix parser error)
    Write-Host "      ${phaseTitle}: [${bar}] ${percent}% " -ForegroundColor Cyan
    # Line 2: Status Label
    Write-Host "      STATUS: ${taskLabel}".PadRight(85) -ForegroundColor Gray
    
    if (${current} -eq ${total}) { 
        Write-Host "" 
        $script:firstBarInPhase = $true
    }
}

function Is-ValidMonth($str) {
    if ($null -eq $str -or $str.Length -lt 5) { return $false }
    return $str -match "^[A-Za-z]{3}-\d{2}$"
}

function Get-ColLetter($n) {
    $s = ""; while ($n -gt 0) { $m = ($n - 1) % 26; $s = [char](65 + $m) + $s; $n = [int](($n - $m) / 26) }; return $s
}

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

try {
    Write-Host "`n[1/3] SCANNING DIRECTORY..." -ForegroundColor Yellow
    $files = Get-ChildItem -Path $currentFolder -Filter *.xlsx | Where-Object { $_.Name -notlike "CONSOLIDATED_*" }
    if ($files.Count -eq 0) { Write-Host "!! No reports found in: $currentFolder" -ForegroundColor Red; return }

    $entityGroups = @{}
    foreach ($f in $files) {
        $wb = $excel.Workbooks.Open($f.FullName); $ws = $wb.Sheets.Item("Entity Profile")
        $p = $ws.Range("B6").Text.Trim().Substring(2, 10); $n = $ws.Range("B4").Text.Trim() -replace '[\\/:*?""<>|]', '_'
        $k = "$p|$n"; if (-not $entityGroups.ContainsKey($k)) { $entityGroups[$k] = New-Object System.Collections.Generic.List[PSObject] }
        $entityGroups[$k].Add($f); $wb.Close($false)
    }

    foreach ($entry in $entityGroups.GetEnumerator()) {
        $key = [string]$entry.Key; $entityFiles = $entry.Value; $parts = $key.Split('|')
        $myPan = $parts[0]; $nameDisp = $parts[1]
        Write-Host "`n======================================================" -ForegroundColor White
        Write-Host " ACTIVE ENTITY: $nameDisp ($myPan)" -ForegroundColor Green
        Write-Host "======================================================" -ForegroundColor White

        $summaryData = @(); $matrixCust = @(); $matrixSupp = @(); $relatedPANs = @()

        # --- PHASE 2: EXTRACTION ---
        Write-Host "[2/3] EXTRACTING DATA FROM SOURCE REPORTS..." -ForegroundColor Yellow
        $fIdx = 0
        foreach ($file in $entityFiles) {
            $fIdx++
            $wb = $excel.Workbooks.Open($file.FullName); $wsP = $wb.Sheets.Item("Entity Profile")
            $code = $wsP.Range("B6").Text.Trim().Substring(0, 2)
            $sN = $stateMap[$code]; if ($null -eq $sN) { $sN = "Unknown" }
            $stHead = "$code-$sN"
            
            Show-Dashboard-Bar $fIdx $entityFiles.Count "Reading $stHead ($($file.Name))" "EXTRACTION"

            foreach ($rpS in @("Related Party Sales - Monthly", "Related Party Purchases-Monthly")) {
                try { $wsRP = $wb.Sheets.Item($rpS)
                    for ($c=2; $c -le 100; $c+=8) { for ($r=4; $r -le 50; $r++) {
                        $rpP = $wsRP.Cells.Item($r, $c).Text.Trim()
                        if ($rpP.Length -eq 10) { $relatedPANs += $rpP }
                        if (-not $rpP -and $r -gt 10) { break }
                    } }
                } catch {}
            }

            $wsG = $wb.Sheets.Item("GSTR1 vs 3B"); $fileMonths = @{}
            for ($i = 4; $i -le $wsG.UsedRange.Rows.Count; $i++) {
                $m = $wsG.Cells.Item($i, 1).Text.Trim()
                if (Is-ValidMonth $m) { $fileMonths[$m] = @{ GT=[double]$wsG.Cells.Item($i, 5).Value2; GI=[double]$wsG.Cells.Item($i, 4).Value2; IT_C=0; II_C=0; IT_S=0; II_S=0 } }
            }

            foreach ($type in @("Customer", "Supplier")) {
                $ws = $wb.Sheets.Item("$type Wise - Monthly Data")
                for ($c = 1; $c -le 250; $c += 9) {
                    $m = $ws.Cells.Item(2, $c).Text.Trim()
                    if (-not $fileMonths.ContainsKey($m)) { continue }
                    for ($r = 4; $r -le 150; $r++) {
                        $cP = $ws.Cells.Item($r, $c + 1).Text.Trim(); $cN = $ws.Cells.Item($r, $c + 2).Text.Trim()
                        if ($cN -like "*Total*" -or (-not $cP -and $r -gt 10)) { break }
                        if ($cP) {
                            $vT = [double]$ws.Cells.Item($r, $c + 3).Value2; $vI = [double]$ws.Cells.Item($r, $c + 5).Value2
                            if ($cP -eq $myPan) {
                                if ($type -eq "Customer") { $fileMonths[$m].IT_C += $vT; $fileMonths[$m].II_C += $vI }
                                else { $fileMonths[$m].IT_S += $vT; $fileMonths[$m].II_S += $vI }
                            } else {
                                $matrixCust += [PSCustomObject]@{ Name=$cN; PAN=$cP; State=$stHead; Month=$m; T=$vT; I=$vI; IsRP=($relatedPANs -contains $cP); Type=$type }
                            }
                        }
                    }
                }
            }
            foreach ($mK in $fileMonths.Keys) {
                $d = $fileMonths[$mK]; $summaryData += [PSCustomObject]@{ Month=$mK; State=$stHead; Type="Customer"; GT=$d.GT; GI=$d.GI; IT=$d.IT_C; II=$d.II_C }
                $summaryData += [PSCustomObject]@{ Month=$mK; State=$stHead; Type="Supplier"; GT=$d.GT; GI=$d.GI; IT=$d.IT_S; II=$d.II_S }
            }
            $wb.Close($false)
        }

        # Audit Display
        $months = ($summaryData.Month + $matrixCust.Month) | Select-Object -Unique | Where-Object { $_ } | Sort-Object { [DateTime]::ParseExact($_, "MMM-yy", $null) }
        $states = $summaryData.State | Select-Object -Unique | Sort-Object
        $auditGT = ($summaryData | Where-Object { $_.Type -eq "Customer" } | Measure-Object GT -Sum).Sum
        $auditIT = ($summaryData | Where-Object { $_.Type -eq "Customer" } | Measure-Object IT -Sum).Sum
        
        Write-Host "  AUDIT SUMMARY (Taxable Revenue):" -ForegroundColor White
        Write-Host "  > Gross: INR $($auditGT.ToString('#,##0.00'))" -ForegroundColor White
        Write-Host "  > Net:   INR $(($auditGT-$auditIT).ToString('#,##0.00'))" -ForegroundColor Green

        # --- PHASE 3: GENERATION ---
        Write-Host "`n[3/3] BUILDING CONSOLIDATED WORKBOOK..." -ForegroundColor Yellow
        $outName = "CONSOLIDATED_$($key -replace '\|', '_').xlsx"
        $outPath = Join-Path $currentFolder $outName
        if (Test-Path $outPath) { try { Remove-Item $outPath -Force -ErrorAction Stop } catch { Write-Host "!! Close $outName and press key..." -ForegroundColor Yellow; $null = $Host.UI.RawUI.ReadKey(); Remove-Item $outPath -Force } }
        
        $outWb = $excel.Workbooks.Add(); $step = 0; $totalSteps = 7
        
        # Updated Sheet Names (Strictly < 31 chars)
        $netCfgs = @(
            @{N="Tax. Value - Internal Sales"; G="GT"; I="IT"; T="Customer Taxable"; Tp="Customer"}, 
            @{N="Inv. Value - Internal Sales"; G="GI"; I="II"; T="Customer Invoice"; Tp="Customer"}, 
            @{N="Tax. Value - Internal Purchases"; G="GT"; I="IT"; T="Supplier Taxable"; Tp="Supplier"}, 
            @{N="Inv. Value - Internal Purchases"; G="GI"; I="II"; T="Supplier Invoice"; Tp="Supplier"}
        )
        
        foreach ($cfg in $netCfgs) {
            $step++; Show-Dashboard-Bar $step $totalSteps "Writing Sheet: $($cfg.N)" "GENERATION"
            $ws = $outWb.Sheets.Add(); $ws.Name = $cfg.N; $r = 1
            foreach ($mode in @("Gross", "Internal", "Net")) {
                $ws.Cells.Item($r, 1).NumberFormat = "@"; $ws.Cells.Item($r, 1) = "$mode - $($cfg.T)"; $ws.Cells.Item($r, 1).Font.Bold = $true; $ws.Cells.Item($r+1, 1) = "Month"; $cc = 2
                foreach ($s in $states) { $ws.Cells.Item($r+1, $cc++) = $s }; $ws.Cells.Item($r+1, $cc) = "Total"
                $dr = $r + 2; foreach ($m in $months) {
                    $ws.Cells.Item($dr, 1).NumberFormat = "@"; $ws.Cells.Item($dr, 1) = [string]$m; $cc = 2
                    foreach ($s in $states) {
                        $recs = $summaryData | Where-Object { $_.Month -eq $m -and $_.State -eq $s -and $_.Type -eq $cfg.Tp }
                        $v = 0; if ($recs) { 
                            if ($mode -eq "Gross") { $v = ($recs | Measure-Object $($cfg.G) -Sum).Sum } 
                            elseif ($mode -eq "Internal") { $v = ($recs | Measure-Object $($cfg.I) -Sum).Sum } 
                            else { $v = ($recs | Measure-Object $($cfg.G) -Sum).Sum - ($recs | Measure-Object $($cfg.I) -Sum).Sum }
                        }
                        $ws.Cells.Item($dr, $cc++) = $v
                    }
                    $ws.Cells.Item($dr, $cc).Formula = "=SUM(B${dr}:$(Get-ColLetter($cc-1))${dr})"; $dr++
                }
                $r = $dr + 2
            }
            $ws.UsedRange.EntireColumn.AutoFit() | Out-Null; $ws.UsedRange.NumberFormat = "#,##0.00" | Out-Null
        }

        foreach ($mCfg in @(@{N="Detailed_Customer";D=$matrixCust | Where-Object { $_.Type -eq "Customer" }}, @{N="Detailed_Supplier";D=$matrixCust | Where-Object { $_.Type -eq "Supplier" }})) {
            $step++; Show-Dashboard-Bar $step $totalSteps "Matrixing: $($mCfg.N)" "GENERATION"
            $ws = $outWb.Sheets.Add(); $ws.Name = $mCfg.N; $ws.Outline.SummaryRow = 0 | Out-Null
            $ws.Cells.Item(1, 1).NumberFormat = "@"; $ws.Cells.Item(1, 1) = "Party / State"; $ws.Cells.Item(1, 2).NumberFormat = "@"; $ws.Cells.Item(1, 2) = "PAN"; $cc = 3
            foreach ($m in $months) { $ws.Cells.Item(1, $cc).NumberFormat = "@"; $ws.Cells.Item(1, $cc++) = [string]$m }
            $ws.Cells.Item(1, $cc) = "Total"; $ws.Range("1:1").Font.Bold = $true; $dr = 2
            foreach ($pg in ($mCfg.D | Group-Object PAN)) {
                $ws.Cells.Item($dr, 1) = $pg.Group[0].Name; $ws.Cells.Item($dr, 2) = $pg.Name; $ws.Rows.Item($dr).Font.Bold = $true; $ws.Rows.Item($dr).Interior.ColorIndex = if ($pg.Group[0].IsRP) { 36 } else { 15 }
                $cc = 3; foreach ($m in $months) { $v = ($pg.Group | Where-Object { $_.Month -eq $m } | Measure-Object T -Sum).Sum; if ($v -gt 0) { $ws.Cells.Item($dr, $cc) = $v }; $cc++ }
                $ws.Cells.Item($dr, $cc).Formula = "=SUM(C${dr}:$(Get-ColLetter($cc-1))${dr})"; $pRow = $dr; $dr++
                foreach ($sg in ($pg.Group | Group-Object State)) {
                    $ws.Cells.Item($dr, 1) = "   >> " + $sg.Name; $cc = 3
                    foreach ($m in $months) { $v = ($sg.Group | Where-Object { $_.Month -eq $m } | Measure-Object T -Sum).Sum; if ($v -gt 0) { $ws.Cells.Item($dr, $cc) = $v }; $cc++ }
                    $ws.Cells.Item($dr, $cc).Formula = "=SUM(C${dr}:$(Get-ColLetter($cc-1))${dr})"; $dr++
                }
                if ($dr - 1 -gt $pRow) { $ws.Rows.Item("$(${pRow}+1):$(${dr}-1)").Group() | Out-Null }
            }
            $ws.UsedRange.EntireColumn.AutoFit() | Out-Null; $ws.UsedRange.NumberFormat = "#,##0.00" | Out-Null
        }

        $step++; Show-Dashboard-Bar $step $totalSteps "Finalizing glossary..." "GENERATION"
        $wsG = $outWb.Sheets.Add(); $wsG.Name = "Audit_Glossary"
        $wsG.Cells.Item(1, 1) = "Dashboard Color Key"; $wsG.Cells.Item(1, 1).Font.Bold = $true
        $wsG.Cells.Item(3, 1).Interior.ColorIndex = 36; $wsG.Cells.Item(3, 2) = "Related Party"
        $wsG.Cells.Item(4, 1).Interior.ColorIndex = 15; $wsG.Cells.Item(4, 2) = "Third Party"
        $wsG.UsedRange.EntireColumn.AutoFit() | Out-Null
        $outWb.SaveAs($outPath, 51) | Out-Null; $outWb.Close($false) | Out-Null
        Write-Host "`n[SUCCESS] CONSOLIDATION COMPLETE" -ForegroundColor Green
        Write-Host "  FILE: $outPath" -ForegroundColor Yellow
        explorer.exe /select,`"$outPath`"
    }
} catch { Write-Host "`nERROR: $($_.Exception.Message)" -ForegroundColor Red } finally {
    if ($null -ne $excel) { $excel.Quit(); [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null }
}
