# Vérification headless de Integrateur-de-Balance.xlsm
$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent $PSScriptRoot
$xlsmPath = Join-Path $root 'dist\Integrateur-de-Balance.xlsm'
$samplePath = Join-Path $root '..\ExpertComptable\public\Balance comparative 2023-2024.xlsx'
. (Join-Path $PSScriptRoot 'Test-Common.ps1')

if (-not (Test-Path $xlsmPath)) { throw "Classeur introuvable : $xlsmPath" }
if (-not (Test-Path $samplePath)) { throw "Fichier balance exemple introuvable : $samplePath" }

$excel = $null
$wb = $null
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.AutomationSecurity = 1

    $wb = $excel.Workbooks.Open($xlsmPath, $false, $false)
    $excel.Run('modUtil.SetTestMode', $true) | Out-Null
    foreach ($n in @('Balance_Import','Balance_Analyse')) { $wb.Sheets.Item($n).Unprotect($null) }

    Write-Host "`n=== Test 1 : structure des feuilles ==="
    $expectedSheets = @('Balance_Import','Balance_Data','Balance_Analyse','Référentiel')
    $actualSheets = @($wb.Sheets | ForEach-Object { $_.Name })
    foreach ($s in $expectedSheets) { Assert-Equal "Feuille présente : $s" $true ($actualSheets -contains $s) }

    Write-Host "`n=== Test 2 : import de la balance réelle (fixture) ==="
    $wb.Names.Item('Bal_ClientName').RefersToRange.Value2 = 'Client Balance Test'
    $wb.Names.Item('Bal_Year').RefersToRange.Value2 = 2024
    $excel.Run('modBalance.ImporterBalance', $samplePath) | Out-Null

    $dataWs = $wb.Sheets.Item('Balance_Data')
    $lastDataRow = $dataWs.Cells.Item($dataWs.Rows.Count, 1).End(-4162).Row
    $importedCount = $lastDataRow - 1
    Write-Host "Lignes importées : $importedCount"
    Assert-Equal 'Import balance - lignes > 0' $true ($importedCount -gt 0)

    $withReading = 0
    $withCycle = @{}
    for ($r = 2; $r -le $lastDataRow; $r++) {
        if ([string]$dataWs.Cells.Item($r, 8).Value2 -ne '') { $withReading++ }
        $cyc = [string]$dataWs.Cells.Item($r, 7).Value2
        if (-not $withCycle.ContainsKey($cyc)) { $withCycle[$cyc] = 0 }
        $withCycle[$cyc] = $withCycle[$cyc] + 1
    }
    Write-Host "Lignes avec lecture écologique : $withReading"
    Assert-Equal 'Import balance - au moins une lecture écologique détectée' $true ($withReading -gt 0)
    $nonAutresCycles = @($withCycle.Keys | Where-Object { $_ -ne 'autres' })
    Assert-Equal 'Import balance - au moins un cycle autre que autres' $true ($nonAutresCycles.Count -gt 0)

    Write-Host "`n=== Test 3 : lancement de l'analyse (comptage des cartes de cycle) ==="
    $excel.Run('modBalance.LancerAnalyse') | Out-Null
    $balWs = $wb.Sheets.Item('Balance_Analyse')
    $cardTextSample = $null
    foreach ($shp in $balWs.Shapes) {
        if ($shp.Name -eq 'CardCycle_immobilisations') { $cardTextSample = $shp.TextFrame2.TextRange.Text }
    }
    Write-Host "Texte carte immobilisations : $cardTextSample"
    Assert-Equal 'Carte cycle immobilisations mise à jour' $true ($null -ne $cardTextSample -and $cardTextSample -match 'compte')

    Write-Host "`n=== Test 4 : export de la synthèse (pont vers Trame-de-Rapport) ==="
    $excel.Run('modBalance.ExporterSynthese') | Out-Null
    $exportPath = Join-Path $root 'dist\Synthese_Balance_Client Balance Test.txt'
    Assert-Equal 'Fichier de synthèse balance créé' $true (Test-Path $exportPath)
    if (Test-Path $exportPath) {
        $line = (Get-Content -Path $exportPath -Raw -Encoding Default).TrimEnd()
        $parts = $line -split '\|', 3
        Assert-Equal 'Synthèse balance - client' 'Client Balance Test' $parts[0]
        Assert-Equal 'Synthèse balance - année' '2024' $parts[1]
        Assert-Equal 'Synthèse balance - résumé non vide' $true ($parts.Count -eq 3 -and $parts[2].Length -gt 0)
    }

    Write-Host "`n=== Test 5 : intégrité des contrôles (OnAction) ==="
    $knownMacros = @('modNav.GoToBalanceImport','modNav.GoToBalanceAnalyse',
                      'modBalance.ImporterBalance','modBalance.LancerAnalyse','modBalance.FiltrerCycle','modBalance.ExporterAnalysePDF','modBalance.ExporterSynthese')
    Test-OnActionIntegrity -Workbook $wb -SheetNames @('Balance_Import','Balance_Analyse') -KnownMacros $knownMacros

} finally {
    if ($wb) { try { $wb.Close($false) } catch {} }
    if ($excel) { try { $excel.Quit() } catch {} }
    if ($wb) { [Runtime.InteropServices.Marshal]::ReleaseComObject($wb) | Out-Null }
    if ($excel) { [Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null }
    Remove-Variable wb, excel -ErrorAction SilentlyContinue
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
}

Write-Summary
