# Script maître : génère UN des classeurs (3 outils autonomes, ou le classeur Unified qui
# regroupe les 3 dans un seul fichier avec un référentiel partagé)
# Usage : powershell -File Build-Workbook.ps1 -Module Profilage|Balance|Rapport|Unified

param(
    [Parameter(Mandatory)][ValidateSet('Profilage','Balance','Rapport','Unified')][string]$Module
)

$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent $PSScriptRoot
$refDataPath = Join-Path $root 'referentiel'
$vbaRoot = Join-Path $root 'vba'
$outDir = Join-Path $root 'dist'

# Seul Balance conserve une (toute petite) macro, pour l'import de fichier - action
# impossible à réaliser par formule (aucun outil ne peut ouvrir un fichier tiers sans
# un minimum de code). Profilage et Rapport sont 100% formules : classeurs .xlsx,
# sans aucune macro, cliquables sans blocage sur PC, mobile et Excel Online.
$ModuleConfig = @{
    Profilage = @{
        OutFile = 'Outil-de-Profilage.xlsx'
        UseVba = $false
        SaveFormat = 51 # xlOpenXMLWorkbook (.xlsx, sans macro)
        ProtectedSheets = @('Profilage_Questionnaire','Profilage_Resultats')
    }
    Balance = @{
        OutFile = 'Integrateur-de-Balance.xlsm'
        UseVba = $true
        SaveFormat = 52 # xlOpenXMLWorkbookMacroEnabled (.xlsm)
        VbaPaths = @((Join-Path $vbaRoot 'common'), (Join-Path $vbaRoot 'balance'))
        ProtectedSheets = @('Balance_Import','Balance_Data','Balance_Analyse')
    }
    Rapport = @{
        OutFile = 'Trame-de-Rapport.xlsx'
        UseVba = $false
        SaveFormat = 51
        ProtectedSheets = @('Rapport_Trame')
    }
    Unified = @{
        # Regroupe les 3 outils dans un seul classeur, avec un référentiel partagé (une seule
        # feuille Référentiel au lieu d'une par classeur) et des champs de synthèse du Rapport
        # reliés par formule aux résultats du Profilage et de la Balance (plus de recopie manuelle
        # entre fichiers séparés). Conserve la macro d'import de balance (UseVba = Balance) : c'est
        # la seule action impossible à réaliser par formule, quel que soit le module qui l'embarque.
        OutFile = 'Fichier Excel Final.xlsm'
        UseVba = $true
        SaveFormat = 52
        VbaPaths = @((Join-Path $vbaRoot 'common'), (Join-Path $vbaRoot 'balance'))
        ProtectedSheets = @('Profilage_Questionnaire')
    }
}
$cfg = $ModuleConfig[$Module]
$outPath = Join-Path $outDir $cfg.OutFile

. (Join-Path $PSScriptRoot 'Common.ps1')
. (Join-Path $PSScriptRoot 'New-Sheets.ps1')
. (Join-Path $PSScriptRoot 'New-Referentiel.ps1')
. (Join-Path $PSScriptRoot 'New-Controls.ps1')
. (Join-Path $PSScriptRoot 'New-ProfilageSheets.ps1')
. (Join-Path $PSScriptRoot 'New-BalanceSheets.ps1')
. (Join-Path $PSScriptRoot 'New-RapportSheet.ps1')
. (Join-Path $PSScriptRoot 'Import-VbaModules.ps1')

if (-not (Test-Path $outDir)) { New-Item -ItemType Directory -Path $outDir -Force | Out-Null }
if (Test-Path $outPath) { Remove-Item $outPath -Force }

$existingExcel = Get-Process EXCEL -ErrorAction SilentlyContinue
if ($existingExcel) {
    Write-Warning "Des processus EXCEL.EXE sont déjà en cours d'exécution ($($existingExcel.Count)). S'ils proviennent d'une exécution précédente en échec, fermez-les avant de relancer."
}

$vbomState = $null
if ($cfg.UseVba) {
    if (-not (Test-AccessVBOM)) {
        Write-Host "Activation temporaire de 'Trust access to the VBA project object model' (HKCU)..."
        $vbomState = Enable-AccessVBOM
    } else {
        Write-Host "AccessVBOM déjà activé - aucune modification du registre nécessaire."
    }
}

$excel = $null
$wb = $null
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.ScreenUpdating = $false
    $excel.EnableEvents = $false

    $wb = $excel.Workbooks.Add()
    $excel.Calculation = -4135 # xlCalculationManual

    Write-Host "[$Module] Création des feuilles..."
    $sheets = New-AllSheets -Workbook $wb -Module $Module

    Write-Host "[$Module] Peuplement du référentiel..."
    New-ReferentielSheet -Sheet $sheets['Référentiel'] -Workbook $wb -RefDataPath $refDataPath -Module $Module

    switch ($Module) {
        'Profilage' {
            Write-Host "Construction du questionnaire de profilage..."
            New-ProfilageQuestionnaireSheet -Sheet $sheets['Profilage_Questionnaire'] -Workbook $wb -RefDataPath $refDataPath

            Write-Host "Construction des résultats de profilage..."
            New-ProfilageResultatsSheet -Sheet $sheets['Profilage_Resultats'] -Workbook $wb

            Write-Host "Construction de l'historique des profils..."
            New-ProfilsHistoriqueSheet -Sheet $sheets['Profils_Historique'] -Workbook $wb
        }
        'Balance' {
            Write-Host "Construction de l'import de balance..."
            New-BalanceImportSheet -Sheet $sheets['Balance_Import'] -Workbook $wb

            Write-Host "Construction des en-têtes Balance_Data..."
            New-BalanceDataSheet -Sheet $sheets['Balance_Data'] -Workbook $wb -RefDataPath $refDataPath

            Write-Host "Construction de l'analyse de balance..."
            New-BalanceAnalyseSheet -Sheet $sheets['Balance_Analyse'] -Workbook $wb -RefDataPath $refDataPath
        }
        'Rapport' {
            Write-Host "Construction de la trame de rapport..."
            New-RapportTrameSheet -Sheet $sheets['Rapport_Trame'] -Workbook $wb -RefDataPath $refDataPath
        }
        'Unified' {
            # Barre de navigation commune aux 7 onglets visibles - mêmes clés d'activation que
            # les classeurs autonomes (import/donnees/analyse/questionnaire/resultats), + 'trame'
            # pour le Rapport (absente des classeurs autonomes, où Rapport est seul dans son fichier).
            $unifiedButtons = @(
                @{ Key='questionnaire'; Label='Questionnaire'; TargetSheet='Profilage_Questionnaire'; TargetCell='B4' }
                @{ Key='resultats';     Label='Résultats';     TargetSheet='Profilage_Resultats';      TargetCell='B4' }
                @{ Key='import';        Label='Import';        TargetSheet='Balance_Import';           TargetCell='B4' }
                @{ Key='donnees';       Label='Données';       TargetSheet='Balance_Data';              TargetCell='B4' }
                @{ Key='analyse';       Label='Analyse';       TargetSheet='Balance_Analyse';           TargetCell='B4' }
                @{ Key='trame';         Label='Rapport';       TargetSheet='Rapport_Trame';             TargetCell='B4' }
            )
            $unifiedTitle = 'Expert Comptable'

            Write-Host "Construction du questionnaire de profilage..."
            New-ProfilageQuestionnaireSheet -Sheet $sheets['Profilage_Questionnaire'] -Workbook $wb -RefDataPath $refDataPath `
                -NavButtons $unifiedButtons -NavTitle $unifiedTitle

            Write-Host "Construction des résultats de profilage..."
            New-ProfilageResultatsSheet -Sheet $sheets['Profilage_Resultats'] -Workbook $wb `
                -NavButtons $unifiedButtons -NavTitle $unifiedTitle -Unified $true

            Write-Host "Construction de l'historique des profils..."
            New-ProfilsHistoriqueSheet -Sheet $sheets['Profils_Historique'] -Workbook $wb

            Write-Host "Construction de l'import de balance..."
            New-BalanceImportSheet -Sheet $sheets['Balance_Import'] -Workbook $wb `
                -NavButtons $unifiedButtons -NavTitle $unifiedTitle

            Write-Host "Construction des en-têtes Balance_Data..."
            New-BalanceDataSheet -Sheet $sheets['Balance_Data'] -Workbook $wb -RefDataPath $refDataPath `
                -NavButtons $unifiedButtons -NavTitle $unifiedTitle

            Write-Host "Construction de l'analyse de balance..."
            New-BalanceAnalyseSheet -Sheet $sheets['Balance_Analyse'] -Workbook $wb -RefDataPath $refDataPath `
                -NavButtons $unifiedButtons -NavTitle $unifiedTitle

            Write-Host "Construction de la trame de rapport..."
            New-RapportTrameSheet -Sheet $sheets['Rapport_Trame'] -Workbook $wb -RefDataPath $refDataPath `
                -NavButtons $unifiedButtons -NavTitle $unifiedTitle -Unified $true
        }
    }

    # Feuilles volontairement NON protégées : le client réutilise ces classeurs pour un
    # mémoire de soutenance et doit pouvoir tout modifier librement (mise en forme, contenu,
    # structure). $cfg.ProtectedSheets sert uniquement à désigner la feuille active à
    # l'enregistrement (voir plus bas) - aucun appel à .Protect() n'est fait.

    if ($cfg.UseVba) {
        Write-Host "Import des modules VBA..."
        Import-VbaModules -Workbook $wb -VbaPaths $cfg.VbaPaths
    }

    $firstSheetName = $cfg.ProtectedSheets[0]
    $sheets[$firstSheetName].Activate()
    $sheets[$firstSheetName].Range('A1').Select() | Out-Null

    Write-Host "Enregistrement en $($cfg.OutFile.Substring($cfg.OutFile.LastIndexOf('.')))..."
    # Sur les classeurs les plus lourds (VBA + centaines de formules/formes), SaveAs lève parfois
    # une COMException transitoire ("Impossible de lire la propriété SaveAs...") alors que
    # l'enregistrement aboutit en réalité (fichier valide, confirmé en le rouvrant à la main) -
    # vraisemblablement un appel COM rejeté pendant qu'Excel finit son rendu interne. Ré-appeler
    # SaveAs sur le MÊME objet Workbook après un premier échec aggrave la situation (le classeur
    # est déjà partiellement "attaché" au fichier cible) - donc après une exception, on vérifie si
    # le fichier a malgré tout été écrit avant de retenter, plutôt que de retenter aveuglément.
    $beforeSave = Get-Date
    $saved = $false
    try {
        $wb.SaveAs($outPath, $cfg.SaveFormat)
        $saved = $true
    } catch {
        $savedAnyway = (Test-Path $outPath) -and ((Get-Item $outPath).LastWriteTime -ge $beforeSave.AddSeconds(-2))
        if ($savedAnyway) {
            Write-Warning "SaveAs a levé une exception COM transitoire, mais le fichier a bien été écrit : $($_.Exception.Message)"
            $saved = $true
        } else {
            throw
        }
    }
    if (-not $saved) { throw "SaveAs a échoué et aucun fichier n'a été produit." }
    Write-Host "Classeur généré : $outPath"

} finally {
    if ($wb) { try { $wb.Close($false) } catch {} }
    if ($excel) { try { $excel.Quit() } catch {} }
    if ($wb) { [Runtime.InteropServices.Marshal]::ReleaseComObject($wb) | Out-Null }
    if ($excel) { [Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null }
    Remove-Variable wb, excel -ErrorAction SilentlyContinue
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()

    if ($vbomState) {
        Write-Host "Restauration du registre AccessVBOM à son état initial..."
        Restore-AccessVBOM -State $vbomState
    }
}
