# Vérification headless de Trame-de-Rapport.xlsm
$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent $PSScriptRoot
$xlsmPath = Join-Path $root 'dist\Trame-de-Rapport.xlsm'
. (Join-Path $PSScriptRoot 'Test-Common.ps1')

if (-not (Test-Path $xlsmPath)) { throw "Classeur introuvable : $xlsmPath" }

$excel = $null
$wb = $null
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.AutomationSecurity = 1

    $wb = $excel.Workbooks.Open($xlsmPath, $false, $false)
    $excel.Run('modUtil.SetTestMode', $true) | Out-Null
    $wb.Sheets.Item('Rapport_Trame').Unprotect($null)

    Write-Host "`n=== Test 1 : structure des feuilles ==="
    $expectedSheets = @('Rapport_Trame','Référentiel')
    $actualSheets = @($wb.Sheets | ForEach-Object { $_.Name })
    foreach ($s in $expectedSheets) { Assert-Equal "Feuille présente : $s" $true ($actualSheets -contains $s) }

    Write-Host "`n=== Test 2 : saisie manuelle des champs de synthèse ==="
    $wb.Names.Item('Rap_ProfilClient').RefersToRange.Cells.Item(1,1).Value2 = 'Client Rapport Test'
    $wb.Names.Item('Rap_ProfilLabel').RefersToRange.Cells.Item(1,1).Value2 = 'Ouvert'
    $wb.Names.Item('Rap_ProfilScore').RefersToRange.Value2 = '19/30'
    $wb.Names.Item('Rap_BalanceClient').RefersToRange.Cells.Item(1,1).Value2 = 'Client Rapport Test (2024)'
    $wb.Names.Item('Rap_BalanceResume').RefersToRange.Cells.Item(1,1).Value2 = '12 lignes avec lecture ecologique.'
    Assert-Equal 'Champ client (profilage) saisi' 'Client Rapport Test' ([string](Get-NamedValue $wb 'Rap_ProfilClient'))
    Assert-Equal 'Champ profil saisi' 'Ouvert' ([string](Get-NamedValue $wb 'Rap_ProfilLabel'))

    Write-Host "`n=== Test 3 : import du pont export/import (Profilage + Balance) ==="
    $profExport = Join-Path $root 'dist\Synthese_Profilage_Client Test SARL.txt'
    $balExport = Join-Path $root 'dist\Synthese_Balance_Client Balance Test.txt'
    if ((Test-Path $profExport) -and (Test-Path $balExport)) {
        $excel.Run('modRapport.ImporterProfilage', $profExport) | Out-Null
        Assert-Equal 'Import profilage - client' 'Client Test SARL' ([string](Get-NamedValue $wb 'Rap_ProfilClient'))
        Assert-Equal 'Import profilage - profil' 'Réticent' ([string](Get-NamedValue $wb 'Rap_ProfilLabel'))
        Assert-Equal 'Import profilage - score' '1/30' ([string](Get-NamedValue $wb 'Rap_ProfilScore'))

        $excel.Run('modRapport.ImporterBalanceSynthese', $balExport) | Out-Null
        Assert-Equal 'Import balance - client' 'Client Balance Test (2024)' ([string](Get-NamedValue $wb 'Rap_BalanceClient'))
        $resume = [string](Get-NamedValue $wb 'Rap_BalanceResume')
        Assert-Equal 'Import balance - résumé non vide' $true ($resume.Length -gt 0)
    } else {
        Write-Host "(ignoré - exécutez Verify-Profilage.ps1 et Verify-Balance.ps1 avant pour produire les fichiers de synthèse)"
    }

    Write-Host "`n=== Test 4 : budget des préconisations ==="
    $wb.Names.Item('Rap_Mission1').RefersToRange.Value2 = $true  # ENERGY 2500, 4 services -> 2500*(1+0.8)=4500
    $wb.Names.Item('Rap_Mission2').RefersToRange.Value2 = $false
    $wb.Names.Item('Rap_Mission3').RefersToRange.Value2 = $false
    $wb.Names.Item('Rap_Mission4').RefersToRange.Value2 = $false
    $wb.Names.Item('Rap_Mission5').RefersToRange.Value2 = $false
    $excel.Run('modRapport.RecalcBudget') | Out-Null
    $budgetText = [string](Get-NamedValue $wb 'Rap_BudgetTotal')
    Write-Host "Budget calculé : $budgetText"
    Assert-Equal 'Budget = 4 500 EUR pour ENERGY seul' $true ($budgetText -match '4.?500')

    Write-Host "`n=== Test 5 : intégrité des contrôles (OnAction) ==="
    $knownMacros = @('modRapport.ExporterRapportPDF','modRapport.RecalcBudget','modRapport.ImporterProfilage','modRapport.ImporterBalanceSynthese')
    Test-OnActionIntegrity -Workbook $wb -SheetNames @('Rapport_Trame') -KnownMacros $knownMacros

} finally {
    if ($wb) { try { $wb.Close($false) } catch {} }
    if ($excel) { try { $excel.Quit() } catch {} }
    if ($wb) { [Runtime.InteropServices.Marshal]::ReleaseComObject($wb) | Out-Null }
    if ($excel) { [Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null }
    Remove-Variable wb, excel -ErrorAction SilentlyContinue
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
}

Write-Summary
