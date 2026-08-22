# Peuple la feuille Référentiel (très masquée) à partir des fichiers .psd1
# et crée les Noms définis que la VBA utilisera pour lire les données.
# Dot-sourcé par Build-Workbook.ps1

function Write-RefTable {
    <#
        Écrit un tableau de hashtables dans une feuille à partir de (StartRow,StartCol),
        avec une ligne d'en-tête (noms des colonnes), et crée un Nom défini couvrant
        uniquement les lignes de données (hors en-tête).
    #>
    param(
        $Sheet, $Workbook,
        [int]$StartRow, [int]$StartCol,
        [string[]]$Columns,
        [object[]]$Rows,
        [scriptblock]$RowMapper,
        [string]$DefinedName
    )
    for ($c = 0; $c -lt $Columns.Count; $c++) {
        $cell = $Sheet.Cells.Item($StartRow, $StartCol + $c)
        Set-CellValue2 -Cell $cell -Val $Columns[$c]
        $cell.Font.Bold = $true
    }
    $r = $StartRow + 1
    foreach ($row in $Rows) {
        $values = & $RowMapper $row
        for ($c = 0; $c -lt $values.Count; $c++) {
            Set-CellValue2 -Cell $Sheet.Cells.Item($r, $StartCol + $c) -Val $values[$c]
        }
        $r++
    }
    $lastRow = [Math]::Max($r - 1, $StartRow + 1)
    $rng = $Sheet.Range($Sheet.Cells.Item($StartRow + 1, $StartCol), $Sheet.Cells.Item($lastRow, $StartCol + $Columns.Count - 1))
    if ($DefinedName) { New-DefinedName -Workbook $Workbook -Name $DefinedName -RangeObj $rng }
    return @{ NextCol = $StartCol + $Columns.Count + 1; LastRow = $lastRow }
}

function New-ReferentielSheet {
    <#
        -Module : 'Profilage' | 'Balance' | 'Rapport' - chaque classeur étant désormais
        autonome, on ne peuple que les tables de référence dont ce module a besoin.
    #>
    param($Sheet, $Workbook, [string]$RefDataPath, [Parameter(Mandatory)][string]$Module)

    # 'Unified' (classeur Suite) a besoin des 3 blocs a la fois - les colonnes de depart de
    # chaque bloc (1, 29, 52) sont deja disjointes, donc les 3 s'ecrivent cote a cote sans
    # collision sur la meme feuille Référentiel, partagee par les 3 outils.
    $needsProfilage = ($Module -eq 'Profilage' -or $Module -eq 'Unified')
    $needsBalance = ($Module -eq 'Balance' -or $Module -eq 'Unified')
    $needsRapport = ($Module -eq 'Rapport' -or $Module -eq 'Unified')

    if ($needsProfilage) {
        $questions = (Import-PsdSafe (Join-Path $RefDataPath 'ProfilQuestions.psd1')).Questions
        # Trié par MinScore croissant : la feuille Résultats utilise LOOKUP() (recherche approchée,
        # exige un vecteur croissant) pour déterminer le profil sans aucune macro.
        # Sort-Object avec un nom de propriété littéral ne trie PAS correctement un tableau de
        # Hashtable (le type que produit Import-PowerShellDataFile pour un bloc @{...}) - il faut
        # un bloc de script explicite pour que l'accès à la clé soit résolu avant le tri.
        $profiles  = (Import-PsdSafe (Join-Path $RefDataPath 'ProfilProfiles.psd1')).Profiles | Sort-Object { $_.MinScore }
    }
    if ($needsBalance) {
        $cyclesData = Import-PsdSafe (Join-Path $RefDataPath 'Cycles.psd1')
        $accountRules = (Import-PsdSafe (Join-Path $RefDataPath 'AccountRules.psd1')).Rules
        $cycleRules = Import-PsdSafe (Join-Path $RefDataPath 'CycleRules.psd1')
    }
    if ($needsRapport) {
        $missions = (Import-PsdSafe (Join-Path $RefDataPath 'MissionTypes.psd1')).Missions
        $indicateurs = (Import-PsdSafe (Join-Path $RefDataPath 'Indicateurs.psd1')).Indicateurs
    }

    if ($needsProfilage) {
    # --- Bloc 1 : Questions (une ligne par question) ---
    Write-RefTable -Sheet $Sheet -Workbook $Workbook -StartRow 1 -StartCol 1 `
        -Columns @('QId','Part','Type','Text','Description') -Rows $questions `
        -RowMapper { param($q) @($q.Id, $q.Part, $q.Type, $q.Text, $q.Description) } `
        -DefinedName 'Ref_Questions' | Out-Null

    # --- Bloc 2 : Options (une ligne par option, questions radio/checkbox) ---
    $optionRows = @()
    foreach ($q in $questions) {
        if ($q.Type -eq 'radio' -or $q.Type -eq 'checkbox') {
            $ord = 1
            foreach ($opt in $q.Options) {
                $optionRows += [pscustomobject]@{ QId = $q.Id; Ord = $ord; Text = $opt.Text; Score = $opt.Score }
                $ord++
            }
        }
    }
    Write-RefTable -Sheet $Sheet -Workbook $Workbook -StartRow 1 -StartCol 7 `
        -Columns @('QId','Ord','OptionText','Score') -Rows $optionRows `
        -RowMapper { param($o) @($o.QId, $o.Ord, $o.Text, $o.Score) } `
        -DefinedName 'Ref_Options' | Out-Null

    # Noms dédiés par question radio : liste des libellés d'options, utilisée comme
    # source de la validation de données (liste déroulante) sur la feuille du questionnaire.
    $optStartRow = 2
    $r = $optStartRow
    foreach ($q in $questions) {
        if ($q.Type -eq 'radio') {
            $count = $q.Options.Count
            $listRng = $Sheet.Range($Sheet.Cells.Item($r, 9), $Sheet.Cells.Item($r + $count - 1, 9))
            New-DefinedName -Workbook $Workbook -Name "Ref_Q$($q.Id)_OptionList" -RangeObj $listRng
            $r += $count
        } elseif ($q.Type -eq 'checkbox') {
            $r += $q.Options.Count
        }
    }

    # --- Bloc 3 : Sujets de la question 7 (notation 1-4) ---
    $q7 = $questions | Where-Object { $_.Id -eq 7 }
    $subjRows = @()
    $ord = 1
    foreach ($s in $q7.Subjects) { $subjRows += [pscustomobject]@{ Ord = $ord; Text = $s }; $ord++ }
    Write-RefTable -Sheet $Sheet -Workbook $Workbook -StartRow 1 -StartCol 12 `
        -Columns @('Ord','SubjectText') -Rows $subjRows `
        -RowMapper { param($s) @($s.Ord, $s.Text) } `
        -DefinedName 'Ref_Q7Subjects' | Out-Null

    # --- Bloc 4 : Profils (Proactif/Ouvert/Prudent/Réticent) ---
    Write-RefTable -Sheet $Sheet -Workbook $Workbook -StartRow 1 -StartCol 15 `
        -Columns @('Key','MinScore','MaxScore','Label','Header','Positionnement','Frein1','Frein2','Frein3','Arg1','Arg2','Arg3','Ton') -Rows $profiles `
        -RowMapper {
            param($p)
            @($p.Key, $p.MinScore, $p.MaxScore, $p.Label, $p.Header, $p.Positionnement,
              $p.Freins[0], $p.Freins[1], $p.Freins[2], $p.Arguments[0], $p.Arguments[1], $p.Arguments[2], $p.Ton)
        } -DefinedName 'Ref_Profiles' | Out-Null
    }

    if ($needsBalance) {
    # --- Bloc 5a : Cycles (méta) ---
    Write-RefTable -Sheet $Sheet -Workbook $Workbook -StartRow 1 -StartCol 29 `
        -Columns @('Key','Name','Description') -Rows $cyclesData.Cycles `
        -RowMapper { param($c) @($c.Key, $c.Name, $c.Description) } `
        -DefinedName 'Ref_CyclesMeta' | Out-Null

    # --- Bloc 5b : Cycles - items (Questions/Issues/Actions), format long ---
    $itemRows = @()
    foreach ($c in $cyclesData.Cycles) {
        foreach ($q in $c.Questions) { $itemRows += [pscustomobject]@{ Cycle = $c.Key; Type = 'Q'; Text = $q } }
        foreach ($i in $c.Issues)    { $itemRows += [pscustomobject]@{ Cycle = $c.Key; Type = 'I'; Text = $i } }
        foreach ($a in $c.Actions)   { $itemRows += [pscustomobject]@{ Cycle = $c.Key; Type = 'A'; Text = $a } }
    }
    Write-RefTable -Sheet $Sheet -Workbook $Workbook -StartRow 1 -StartCol 33 `
        -Columns @('CycleKey','ItemType','ItemText') -Rows $itemRows `
        -RowMapper { param($it) @($it.Cycle, $it.Type, $it.Text) } `
        -DefinedName 'Ref_CyclesItems' | Out-Null

    # --- Bloc 5c : Fallback générique (comptes hors des 9 cycles nommés) ---
    $fbRows = @()
    foreach ($q in $cyclesData.GeneralQuestions)   { $fbRows += [pscustomobject]@{ Type = 'Q'; Text = $q } }
    foreach ($i in $cyclesData.CrossCuttingIssues) { $fbRows += [pscustomobject]@{ Type = 'I'; Text = $i } }
    foreach ($a in $cyclesData.CrossCuttingActions){ $fbRows += [pscustomobject]@{ Type = 'A'; Text = $a } }
    Write-RefTable -Sheet $Sheet -Workbook $Workbook -StartRow 1 -StartCol 37 `
        -Columns @('ItemType','ItemText') -Rows $fbRows `
        -RowMapper { param($it) @($it.Type, $it.Text) } `
        -DefinedName 'Ref_GeneralFallback' | Out-Null

    # --- Bloc 6 : Règles de lecture écologique par préfixe de compte ---
    Write-RefTable -Sheet $Sheet -Workbook $Workbook -StartRow 1 -StartCol 40 `
        -Columns @('Ord','Prefix','Reading','Question','Suggestion') -Rows $accountRules `
        -RowMapper { param($r) @($r.Order, $r.Prefix, $r.Reading, $r.Question, $r.Suggestion) } `
        -DefinedName 'Ref_AccountRules' | Out-Null

    # --- Bloc 7 : Règles de classification par cycle ---
    Write-RefTable -Sheet $Sheet -Workbook $Workbook -StartRow 1 -StartCol 46 `
        -Columns @('Ord','Prefix','Cycle') -Rows $cycleRules.Rules `
        -RowMapper { param($r) @($r.Order, $r.Prefix, $r.Cycle) } `
        -DefinedName 'Ref_CycleRules' | Out-Null
    $defCell = $Sheet.Cells.Item(1, 50)
    Set-CellValue2 -Cell $defCell -Val $cycleRules.DefaultCycle
    New-DefinedName -Workbook $Workbook -Name 'Ref_DefaultCycle' -RangeObj $defCell
    }

    if ($needsRapport) {
    # --- Bloc 8 : Types de mission (préconisations) ---
    Write-RefTable -Sheet $Sheet -Workbook $Workbook -StartRow 1 -StartCol 52 `
        -Columns @('Code','Name','Description','BasePrice','Services','Tools') -Rows $missions `
        -RowMapper { param($m) @($m.Code, $m.Name, $m.Description, $m.BasePrice, ($m.Services -join ' | '), ($m.Tools -join ' | ')) } `
        -DefinedName 'Ref_Missions' | Out-Null

    # --- Bloc 9 : Indicateurs environnementaux (E8) ---
    Write-RefTable -Sheet $Sheet -Workbook $Workbook -StartRow 1 -StartCol 59 `
        -Columns @('Nom','Unite') -Rows $indicateurs `
        -RowMapper { param($ind) @($ind.Nom, $ind.Unite) } `
        -DefinedName 'Ref_Indicateurs' | Out-Null
    }

    $Sheet.Visible = 2 # xlSheetVeryHidden
}
