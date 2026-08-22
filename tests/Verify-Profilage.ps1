# Vérification headless de Outil-de-Profilage.xlsm
$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent $PSScriptRoot
$xlsmPath = Join-Path $root 'dist\Outil-de-Profilage.xlsm'
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
    foreach ($n in @('Profilage_Questionnaire','Profilage_Resultats')) { $wb.Sheets.Item($n).Unprotect($null) }

    Write-Host "`n=== Test 1 : structure des feuilles ==="
    $expectedSheets = @('Profilage_Questionnaire','Profilage_Resultats','Profils_Historique','Référentiel')
    $actualSheets = @($wb.Sheets | ForEach-Object { $_.Name })
    foreach ($s in $expectedSheets) { Assert-Equal "Feuille présente : $s" $true ($actualSheets -contains $s) }

    Write-Host "`n=== Test 2 : scoring - cas Proactif ==="
    $wb.Names.Item('Q1_Answer').RefersToRange.Value2 = "Oui, dans le cadre d'une stratégie structurée"
    $wb.Names.Item('Q2_Answer').RefersToRange.Value2 = 'Oui, clairement identifiés'
    $wb.Names.Item('Q3_Answer').RefersToRange.Value2 = 'Oui, en interne et en externe'
    for ($i = 1; $i -le 5; $i++) { $wb.Names.Item("Q4_Opt$i").RefersToRange.Value2 = $false }
    for ($i = 1; $i -le 5; $i++) { $wb.Names.Item("Q5_Opt$i").RefersToRange.Value2 = $true }
    $wb.Names.Item('Q6_Answer').RefersToRange.Value2 = "Oui, je pense que c'est le bon interlocuteur"
    for ($i = 1; $i -le 5; $i++) { $wb.Names.Item("Q7_S$i").RefersToRange.Value2 = 4 }
    $wb.Names.Item('Q8_Answer').RefersToRange.Value2 = 'Oui, dès maintenant'
    $wb.Names.Item('Q9_Answer').RefersToRange.Value2 = 'Oui'
    $excel.Run('modProfilage.ValiderQuestionnaire') | Out-Null
    Assert-Equal 'Score brut (Proactif)' 25 ([int](Get-NamedValue $wb 'Res_Brut'))
    Assert-Equal 'Score final (Proactif)' 25 ([int](Get-NamedValue $wb 'Res_ScoreFinal'))
    Assert-Equal 'Profil (Proactif)' 'Proactif' ([string](Get-NamedValue $wb 'Res_ProfilKey'))

    Write-Host "`n=== Test 3 : scoring - cas Réticent ==="
    $wb.Names.Item('Q1_Answer').RefersToRange.Value2 = "Non, ce sujet n'est pas encore à l'ordre du jour"
    $wb.Names.Item('Q2_Answer').RefersToRange.Value2 = 'Je ne sais pas'
    $wb.Names.Item('Q3_Answer').RefersToRange.Value2 = 'Non, aucun besoin identifié'
    for ($i = 1; $i -le 5; $i++) { $wb.Names.Item("Q4_Opt$i").RefersToRange.Value2 = $true }
    for ($i = 1; $i -le 5; $i++) { $wb.Names.Item("Q5_Opt$i").RefersToRange.Value2 = $false }
    $wb.Names.Item('Q6_Answer').RefersToRange.Value2 = "Non, je préfère d'autres profils de conseil"
    for ($i = 1; $i -le 5; $i++) { $wb.Names.Item("Q7_S$i").RefersToRange.Value2 = 1 }
    $wb.Names.Item('Q8_Answer').RefersToRange.Value2 = 'Non'
    $wb.Names.Item('Q9_Answer').RefersToRange.Value2 = 'Non'
    $excel.Run('modProfilage.ValiderQuestionnaire') | Out-Null
    Assert-Equal 'Score brut (Réticent)' -5 ([int](Get-NamedValue $wb 'Res_Brut'))
    Assert-Equal 'Score final (Réticent)' 1 ([int](Get-NamedValue $wb 'Res_ScoreFinal'))
    Assert-Equal 'Profil (Réticent)' 'Réticent' ([string](Get-NamedValue $wb 'Res_ProfilKey'))

    Write-Host "`n=== Test 4 : enregistrement du profil (historique) ==="
    $wb.Names.Item('Res_ClientName').RefersToRange.Value2 = 'Client Test SARL'
    $excel.Run('modProfilage.EnregistrerProfil') | Out-Null
    $hist = $wb.Sheets.Item('Profils_Historique')
    $lastRow = $hist.Cells.Item($hist.Rows.Count, 1).End(-4162).Row
    Assert-Equal 'Historique - client enregistré' 'Client Test SARL' ([string]$hist.Cells.Item($lastRow, 2).Value2)
    Assert-Equal 'Historique - profil enregistré' 'Réticent' ([string]$hist.Cells.Item($lastRow, 4).Value2)

    Write-Host "`n=== Test 5 : export de la synthèse (pont vers Trame-de-Rapport) ==="
    $excel.Run('modProfilage.ExporterSynthese') | Out-Null
    $exportPath = Join-Path $root 'dist\Synthese_Profilage_Client Test SARL.txt'
    Assert-Equal 'Fichier de synthèse profilage créé' $true (Test-Path $exportPath)
    if (Test-Path $exportPath) {
        $line = Get-Content -Path $exportPath -Raw -Encoding Default
        Assert-Equal 'Contenu synthèse profilage' $true ($line.TrimEnd() -eq 'Client Test SARL|Réticent|1')
    }

    Write-Host "`n=== Test 6 : intégrité des contrôles (OnAction) ==="
    $knownMacros = @('modNav.GoToProfilageQuestionnaire','modNav.GoToProfilageResultats',
                      'modProfilage.ValiderQuestionnaire','modProfilage.ExporterProfilPDF','modProfilage.EnregistrerProfil','modProfilage.NouveauProfilage','modProfilage.ExporterSynthese')
    Test-OnActionIntegrity -Workbook $wb -SheetNames @('Profilage_Questionnaire','Profilage_Resultats') -KnownMacros $knownMacros

} finally {
    if ($wb) { try { $wb.Close($false) } catch {} }
    if ($excel) { try { $excel.Quit() } catch {} }
    if ($wb) { [Runtime.InteropServices.Marshal]::ReleaseComObject($wb) | Out-Null }
    if ($excel) { [Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null }
    Remove-Variable wb, excel -ErrorAction SilentlyContinue
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
}

Write-Summary
