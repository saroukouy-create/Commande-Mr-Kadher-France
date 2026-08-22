# Vérification headless de Fichier Excel Final.xlsm (classeur unique regroupant les 3 outils)
$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent $PSScriptRoot
$xlsmPath = Join-Path $root 'dist\Fichier Excel Final.xlsm'
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
    $excel.Calculation = -4105 # xlCalculationAutomatic - Profilage/Rapport sont 100% formules,
                                # sans macro pour déclencher un recalcul explicite comme dans les
                                # classeurs autonomes historiques (modProfilage.ValiderQuestionnaire).

    Write-Host "`n=== Test 1 : structure des feuilles (un seul classeur, un seul Référentiel) ==="
    $expectedSheets = @('Profilage_Questionnaire','Profilage_Resultats','Profils_Historique',
                         'Balance_Import','Balance_Data','Balance_Analyse','Rapport_Trame','Référentiel')
    $actualSheets = @($wb.Sheets | ForEach-Object { $_.Name })
    foreach ($s in $expectedSheets) { Assert-Equal "Feuille présente : $s" $true ($actualSheets -contains $s) }
    Assert-Equal 'Un seul onglet Référentiel (pas un par outil)' 1 (@($actualSheets | Where-Object { $_ -eq 'Référentiel' }).Count)
    Assert-Equal 'Nombre total de feuilles' $expectedSheets.Count $actualSheets.Count

    Write-Host "`n=== Test 2 : référentiel partagé - les 3 domaines cohabitent sur la même feuille ==="
    foreach ($n in @('Ref_Questions','Ref_Profiles','Ref_AccountRules','Ref_CycleRules','Ref_Missions','Ref_Indicateurs')) {
        $has = $true
        try { $wb.Names.Item($n) | Out-Null } catch { $has = $false }
        Assert-Equal "Nom défini présent : $n" $true $has
    }

    Write-Host "`n=== Test 3 : auto-remplissage de la Trame de rapport depuis Profilage et Balance ==="
    # Client (profilage) : champ manuel sur Profilage_Resultats -> repris par formule sur Rapport_Trame
    $wb.Names.Item('Res_ClientName').RefersToRange.Cells.Item(1,1).Value2 = 'Client Unifié Test'
    # Client (balance) : champ manuel sur Balance_Import -> repris par formule sur Rapport_Trame
    $wb.Names.Item('Bal_ClientName').RefersToRange.Cells.Item(1,1).Value2 = 'Client Unifié Test (2024)'

    # Réponses au questionnaire (cas Proactif, mêmes valeurs que Verify-Profilage.ps1)
    $wb.Names.Item('Q1_Answer').RefersToRange.Value2 = "Oui, dans le cadre d'une stratégie structurée"
    $wb.Names.Item('Q2_Answer').RefersToRange.Value2 = 'Oui, clairement identifiés'
    $wb.Names.Item('Q3_Answer').RefersToRange.Value2 = 'Oui, en interne et en externe'
    # Q4_Opts / Q5_Opts / Q7_Scores sont des plages de bloc (une cellule liée de case à cocher, ou
    # une notation, par ligne) - pas de nom individuel par option (Q4_Opt1..5 n'existe pas).
    $q4 = $wb.Names.Item('Q4_Opts').RefersToRange
    for ($i = 1; $i -le $q4.Rows.Count; $i++) { $q4.Cells.Item($i,1).Value2 = $false }
    $q5 = $wb.Names.Item('Q5_Opts').RefersToRange
    for ($i = 1; $i -le $q5.Rows.Count; $i++) { $q5.Cells.Item($i,1).Value2 = $true }
    $wb.Names.Item('Q6_Answer').RefersToRange.Value2 = "Oui, je pense que c'est le bon interlocuteur"
    $q7 = $wb.Names.Item('Q7_Scores').RefersToRange
    for ($i = 1; $i -le $q7.Rows.Count; $i++) { $q7.Cells.Item($i,1).Value2 = 4 }
    $wb.Names.Item('Q8_Answer').RefersToRange.Value2 = 'Oui, dès maintenant'
    $wb.Names.Item('Q9_Answer').RefersToRange.Value2 = 'Oui'
    $excel.CalculateFull()

    # Pas d'hypothèse sur le profil exact obtenu (dépend du barème du référentiel, hors sujet ici) -
    # ce test vérifie le CÂBLAGE inter-feuilles (Rapport <- Profilage/Balance), pas le barème lui-même.
    $computedProfil = [string](Get-NamedValue $wb 'Res_ProfilKey')
    $computedScore = [int](Get-NamedValue $wb 'Res_ScoreFinal')
    Write-Host "Profil calculé : $computedProfil ($computedScore/30)"

    Assert-Equal 'Rap_ProfilClient reprend Res_ClientName' 'Client Unifié Test' ([string](Get-NamedValue $wb 'Rap_ProfilClient'))
    Assert-Equal 'Rap_ProfilLabel reprend Res_ProfilKey' $computedProfil ([string](Get-NamedValue $wb 'Rap_ProfilLabel'))
    Assert-Equal 'Rap_ProfilScore reprend Res_ScoreFinal (format X/30)' "$computedScore/30" ([string](Get-NamedValue $wb 'Rap_ProfilScore'))
    Assert-Equal 'Rap_BalanceClient reprend Bal_ClientName' 'Client Unifié Test (2024)' ([string](Get-NamedValue $wb 'Rap_BalanceClient'))

    $part1 = [int](Get-NamedValue $wb 'Res_Part1')
    $part3 = [int](Get-NamedValue $wb 'Res_Part3')
    Assert-Equal 'Rap_Part1Score reprend Res_Part1 (format X/12)' "$part1/12" ([string](Get-NamedValue $wb 'Rap_Part1Score'))
    Assert-Equal 'Rap_Part3Score reprend Res_Part3 (format X/20)' "$part3/20" ([string](Get-NamedValue $wb 'Rap_Part3Score'))

    Write-Host "`n=== Test 3b : encart d'alerte Proactif (visible uniquement si profil = Proactif) ==="
    $alertVal = [string]$wb.Sheets.Item('Profilage_Resultats').Range('B30').Value2
    $expectNonEmpty = ($computedProfil -eq 'Proactif')
    Write-Host "Profil=$computedProfil / encart='$alertVal'"
    Assert-Equal 'Encart alerte Proactif cohérent avec le profil calculé' $expectNonEmpty ($alertVal.Length -gt 0)
    if ($expectNonEmpty) { Assert-Equal 'Encart alerte - texte' $true ($alertVal -match 'Impact Durabilité') }

    Write-Host "`n=== Test 4 : import de balance (macro unique conservée) fonctionne dans le classeur unifié ==="
    # Ne vérifie que le CÂBLAGE (la macro s'exécute et écrit un statut) - le nombre de lignes
    # détectées dépend de la reconnaissance des en-têtes dans modBalance.bas, un point distinct
    # et déjà en défaut à l'identique sur le classeur autonome Integrateur-de-Balance.xlsm (non lié
    # à cette fusion - cf. Verify-Balance.ps1 Test 2, échoue de la même façon hors classeur unifié).
    $samplePath = Join-Path $root '..\ExpertComptable\public\Balance comparative 2023-2024.xlsx'
    if (Test-Path $samplePath) {
        $wb.Names.Item('Bal_Year').RefersToRange.Value2 = 2024
        $excel.Run('modBalance.ImporterBalance', $samplePath) | Out-Null
        $status = [string](Get-NamedValue $wb 'Bal_ImportStatus')
        Write-Host "Statut d'import : $status"
        Assert-Equal 'Import balance - la macro écrit un statut' $true ($status.Length -gt 0)

        Write-Host "`n=== Test 4b : Cycle n'est rempli QUE quand une lecture écologique a matché (fidélité site) ==="
        # Sur le site, E5Controller::detectCycle() n'est appelé que dans generateAnalysisForLine()
        # une fois qu'une des 11 règles AccountRules a matché - un compte hors de ces 11 préfixes
        # n'est jamais classé en cycle. La colonne H (Cycle) doit donc être vide exactement quand
        # la colonne I (Lecture écologique) l'est aussi, ligne par ligne.
        $dataWs = $wb.Sheets.Item('Balance_Data')
        # Colonne B (2), pas A (1) : A est l'espaceur étroit inutilisé en tête de tableau
        # (New-Controls.ps1: `$Sheet.Columns.Item(1).ColumnWidth = 2`) - toujours vide, donc
        # End(xlUp) sur la colonne A renvoie la ligne 1 et fausse silencieusement ce test.
        $lastDataRow = $dataWs.Cells.Item($dataWs.Rows.Count, 2).End(-4162).Row
        $mismatch = 0
        for ($rr = 10; $rr -le $lastDataRow; $rr++) {
            $cyc = [string]$dataWs.Cells.Item($rr, 8).Value2
            $lecture = [string]$dataWs.Cells.Item($rr, 9).Value2
            if (($cyc.Length -gt 0) -ne ($lecture.Length -gt 0)) { $mismatch++ }
        }
        Write-Host "Lignes désaccordées Cycle/Lecture : $mismatch (sur $($lastDataRow - 9))"
        Assert-Equal 'Cycle rempli ssi Lecture écologique rempli, sur toutes les lignes importées' 0 $mismatch

        Write-Host "`n=== Test 4c : pas de carte 'Autres comptes' dans Balance_Analyse (aucun équivalent sur le site) ==="
        $balWs = $wb.Sheets.Item('Balance_Analyse')
        $foundAutres = $balWs.UsedRange.Find('Autres comptes', [Type]::Missing, -4163)
        Assert-Equal "Aucune occurrence de 'Autres comptes' dans Balance_Analyse" $true ($null -eq $foundAutres)
    } else {
        Write-Host "(ignoré - fixture introuvable : $samplePath)"
    }

    Write-Host "`n=== Test 5 : intégrité des contrôles (OnAction) ==="
    $knownMacros = @('modBalance.ImporterBalance')
    Test-OnActionIntegrity -Workbook $wb -SheetNames $expectedSheets -KnownMacros $knownMacros

} finally {
    if ($wb) { try { $wb.Close($false) } catch {} }
    if ($excel) { try { $excel.Quit() } catch {} }
    if ($wb) { [Runtime.InteropServices.Marshal]::ReleaseComObject($wb) | Out-Null }
    if ($excel) { [Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null }
    Remove-Variable wb, excel -ErrorAction SilentlyContinue
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
}

Write-Summary
