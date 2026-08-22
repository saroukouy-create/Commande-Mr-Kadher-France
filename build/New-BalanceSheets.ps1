# Construit les feuilles Balance_Import et Balance_Analyse
# Dot-sourcé par Build-Workbook.ps1
#
# Sans macro, sauf UNE : l'import de fichier (impossible à faire par formule, quelle que
# soit la plateforme - aucun tableur ne peut ouvrir un fichier tiers sans un minimum de
# code). Tout le reste (classification, comptages, navigation) est en formules pures,
# calculées en direct - fonctionne aussi si l'utilisateur préfère coller ses données à la
# main plutôt que d'utiliser le bouton d'import.

$Script:BalanceNavButtons = @(
    @{ Key='import';   Label='Import';   TargetSheet='Balance_Import';  TargetCell='B4' }
    @{ Key='donnees';  Label='Données';  TargetSheet='Balance_Data';    TargetCell='B4' }
    @{ Key='analyse';  Label='Analyse';  TargetSheet='Balance_Analyse'; TargetCell='B4' }
)

function New-BalanceImportSheet {
    param($Sheet, $Workbook,
          [array]$NavButtons = $Script:BalanceNavButtons, [string]$NavTitle = 'Intégrateur de balance')

    Add-NavBar -Sheet $Sheet -Workbook $Workbook -Title $NavTitle -Buttons $NavButtons -ActiveKey 'import'
    for ($c = 2; $c -le 13; $c++) { $Sheet.Cells.Item(1, $c).ColumnWidth = 9.5 }

    Add-SectionHeader -Sheet $Sheet -CellAddr 'B4' -Text 'Intégrateur de balance comptable' -MergeCols 12 | Out-Null

    $Sheet.Cells.Item(7,2).Value2 = 'Nom du client :'
    Set-CellStyle -Range $Sheet.Cells.Item(7,2) -Bold $true -HAlign 'left'
    $cn = $Sheet.Range('D7:G7'); $cn.Merge() | Out-Null
    $cn.Interior.Color = ConvertTo-OleColor $Script:Palette.LightGray; $cn.Borders.Color = ConvertTo-OleColor $Script:Palette.Border; $cn.Locked = $false
    New-DefinedName -Workbook $Workbook -Name 'Bal_ClientName' -RangeObj $cn

    $Sheet.Cells.Item(9,2).Value2 = 'Email du client :'
    Set-CellStyle -Range $Sheet.Cells.Item(9,2) -Bold $true -HAlign 'left'
    $ce = $Sheet.Range('D9:G9'); $ce.Merge() | Out-Null
    $ce.Interior.Color = ConvertTo-OleColor $Script:Palette.LightGray; $ce.Borders.Color = ConvertTo-OleColor $Script:Palette.Border; $ce.Locked = $false
    New-DefinedName -Workbook $Workbook -Name 'Bal_ClientEmail' -RangeObj $ce

    $Sheet.Cells.Item(11,2).Value2 = 'Année (N) à analyser :'
    Set-CellStyle -Range $Sheet.Cells.Item(11,2) -Bold $true -HAlign 'left'
    $yr = $Sheet.Cells.Item(11,4)
    $yr.Interior.Color = ConvertTo-OleColor $Script:Palette.LightGray; $yr.Borders.Color = ConvertTo-OleColor $Script:Palette.Border; $yr.Locked = $false
    $yr.HorizontalAlignment = -4108
    New-DefinedName -Workbook $Workbook -Name 'Bal_Year' -RangeObj $yr

    $infoRng = $Sheet.Range('B13:M18')
    $infoRng.Merge() | Out-Null
    $infoRng.Value2 = "Deux façons d'importer :`r1) Bouton ci-dessous : choisir un fichier .xlsx/.xls/.csv - les colonnes Compte/Libellé/années sont détectées automatiquement (PC uniquement, nécessite les macros).`r2) Coller directement vos données dans l'onglet 'Données' (colonnes Compte/Libellé/Année N-1/Année N) - fonctionne partout, y compris sans macro."
    Set-CellStyle -Range $infoRng -Fill $Script:Palette.LightGray -FontColor '495057' -Size 9.5 -HAlign 'left' -VAlign 'top' -Wrap $true

    Add-NavButton -Sheet $Sheet -Left ($Sheet.Cells.Item(20,2).Left) -Top ($Sheet.Cells.Item(20,2).Top) -Width 240 -Height 34 `
        -Label 'Choisir un fichier...' -Macro 'modBalance.ImporterBalance' -Fill $Script:Palette.Green -TextColor $Script:Palette.White | Out-Null
    Add-NavHyperlink -Sheet $Sheet -Workbook $Workbook -Left ($Sheet.Cells.Item(20,2).Left + 250) -Top ($Sheet.Cells.Item(20,2).Top) -Width 240 -Height 34 `
        -Label 'Coller mes données à la main →' -TargetSheet 'Balance_Data' -TargetCell 'B4' -Fill '673AB7' -TextColor $Script:Palette.White | Out-Null

    $status = $Sheet.Range('B22:M25'); $status.Merge() | Out-Null
    Set-CellStyle -Range $status -FontColor $Script:Palette.DarkText -Size 9.5 -HAlign 'left' -VAlign 'top' -Wrap $true
    New-DefinedName -Workbook $Workbook -Name 'Bal_ImportStatus' -RangeObj $status

    Add-NavHyperlink -Sheet $Sheet -Workbook $Workbook -Left ($Sheet.Cells.Item(27,2).Left) -Top ($Sheet.Cells.Item(27,2).Top) -Width 260 -Height 34 `
        -Label 'Voir l''analyse écologique →' -TargetSheet 'Balance_Analyse' -TargetCell 'B4' -Fill $Script:Palette.Green -TextColor $Script:Palette.White | Out-Null
}

function New-BalanceAnalyseSheet {
    param($Sheet, $Workbook, [string]$RefDataPath,
          [array]$NavButtons = $Script:BalanceNavButtons, [string]$NavTitle = 'Intégrateur de balance')

    Add-NavBar -Sheet $Sheet -Workbook $Workbook -Title $NavTitle -Buttons $NavButtons -ActiveKey 'analyse'
    for ($c = 2; $c -le 13; $c++) { $Sheet.Cells.Item(1, $c).ColumnWidth = 9.5 }

    Add-SectionHeader -Sheet $Sheet -CellAddr 'B4' -Text 'Synthèse par cycle écologique' -MergeCols 12 | Out-Null

    $cyclesData = Import-PsdSafe (Join-Path $RefDataPath 'Cycles.psd1')
    # Pas de carte "autres" : le site (E6.blade.php) n'affiche que les 9 cycles configurés,
    # jamais de case fourre-tout pour les comptes non classés - cf. Cycle désormais conditionné
    # à une règle de lecture appariée (New-Controls.ps1), donc aucun compte n'atterrit plus dans
    # 'autres' de toute façon.
    $cardKeys = @($cyclesData.Cycles | ForEach-Object { $_.Key })
    $cardNames = @{}
    foreach ($c in $cyclesData.Cycles) { $cardNames[$c.Key] = $c.Name }

    # --- Cartes (comptage en direct par formule + lien hypertexte vers leur section) ---
    $perRow = 3
    $blockAnchorRow = @{}
    $anchorRowCursor = 18 # calculé après la boucle des cartes, réservé ici pour lisibilité
    for ($i = 0; $i -lt $cardKeys.Count; $i++) {
        $key = $cardKeys[$i]
        $g = [math]::Floor($i / $perRow)
        $colGroup = $i % $perRow
        $colStart = 2 + $colGroup * 4   # B / F / J
        $rowStart = 6 + $g * 3          # 6, 9, 12

        $cardRng = $Sheet.Range($Sheet.Cells.Item($rowStart, $colStart), $Sheet.Cells.Item($rowStart + 1, $colStart + 2))
        $cardRng.Merge() | Out-Null
        Set-CellStyle -Range $cardRng -Fill $Script:Palette.Green -FontColor $Script:Palette.White -Bold $true -Size 10 -HAlign 'center' -Wrap $true
        # COUNTIFS avec le critère joker "?*" (pas "<>") : BalData_Lecture est une FORMULE qui
        # renvoie "" quand aucune règle ne matche - le critère "<>" compte ces cellules comme
        # non vides (elles contiennent bien une formule), donnant un total identique à un simple
        # COUNTIF par cycle. "?*" ("au moins un caractère") exclut correctement les "" calculés.
        # Le site ne compte que les lignes réellement analysées (règle de lecture appariée) -
        # cf. E6Controller::index, $cycleCounts sur les lignes EcologicalAnalysis existantes uniquement.
        $cardRng.Formula = "=`"$($cardNames[$key])`"&CHAR(10)&COUNTIFS(BalData_Cycle,`"$key`",BalData_Lecture,`"?*`")&`" compte(s)`""
        Write-Output -NoEnumerate $cardRng | Out-Null
    }
    $anchorRowCursor = 6 + ([math]::Ceiling($cardKeys.Count / $perRow)) * 3 + 2

    # --- Blocs de référence par cycle (Constats / Questions / Actions - contenu fixe,
    #     issu du référentiel à la construction ; aucune saisie utilisateur ici). ---
    $r = $anchorRowCursor
    foreach ($key in $cardKeys) {
        $hdrRng = $Sheet.Range($Sheet.Cells.Item($r,2), $Sheet.Cells.Item($r,13)); $hdrRng.Merge() | Out-Null
        $hdrRng.Value2 = $cardNames[$key]
        Set-CellStyle -Range $hdrRng -Fill $Script:Palette.GreenDark -FontColor $Script:Palette.White -Bold $true -Size 11 -HAlign 'left'
        $Sheet.Rows.Item($r).RowHeight = 22
        $blockAnchorRow[$key] = $r
        $r++

        $cyc = $cyclesData.Cycles | Where-Object { $_.Key -eq $key }
        if ($cyc) {
            $qText = ($cyc.Issues   | ForEach-Object { "- $_" }) -join "`r"
            $iText = ($cyc.Questions | ForEach-Object { "- $_" }) -join "`r"
            $aText = ($cyc.Actions  | ForEach-Object { "- $_" }) -join "`r"
        } else {
            $qText = ($cyclesData.CrossCuttingIssues  | ForEach-Object { "- $_" }) -join "`r"
            $iText = ($cyclesData.GeneralQuestions    | ForEach-Object { "- $_" }) -join "`r"
            $aText = ($cyclesData.CrossCuttingActions | ForEach-Object { "- $_" }) -join "`r"
        }

        $hCols = @(
            @{ Label='Constats';           Text=$qText; Cols=2,5 }
            @{ Label='Questions à poser';  Text=$iText; Cols=6,9 }
            @{ Label='Actions proposées';  Text=$aText; Cols=10,13 }
        )
        foreach ($h in $hCols) {
            $lblCell = $Sheet.Range($Sheet.Cells.Item($r, $h.Cols[0]), $Sheet.Cells.Item($r, $h.Cols[1])); $lblCell.Merge() | Out-Null
            $lblCell.Value2 = $h.Label
            Set-CellStyle -Range $lblCell -Bold $true -FontColor $Script:Palette.DarkText -HAlign 'left'
        }
        $r++
        foreach ($h in $hCols) {
            $bodyCell = $Sheet.Range($Sheet.Cells.Item($r, $h.Cols[0]), $Sheet.Cells.Item($r + 4, $h.Cols[1])); $bodyCell.Merge() | Out-Null
            $bodyCell.Value2 = $h.Text
            Set-CellStyle -Range $bodyCell -Fill $Script:Palette.LightGray -FontColor $Script:Palette.DarkText -Size 9 -HAlign 'left' -VAlign 'top' -Wrap $true
        }
        $r += 6
    }

    # Liens hypertexte des cartes vers leur bloc (posés après coup : les lignes des blocs
    # ne sont connues qu'une fois tous les blocs construits).
    for ($i = 0; $i -lt $cardKeys.Count; $i++) {
        $key = $cardKeys[$i]
        $g = [math]::Floor($i / $perRow)
        $colGroup = $i % $perRow
        $colStart = 2 + $colGroup * 4
        $rowStart = 6 + $g * 3
        $cardRng = $Sheet.Range($Sheet.Cells.Item($rowStart, $colStart), $Sheet.Cells.Item($rowStart + 1, $colStart + 2))
        $Sheet.Hyperlinks.Add($cardRng, '', "'Balance_Analyse'!B$($blockAnchorRow[$key])", [Type]::Missing, $cardNames[$key]) | Out-Null
    }

    $noteRng = $Sheet.Range($Sheet.Cells.Item($r,2), $Sheet.Cells.Item($r+1,13)); $noteRng.Merge() | Out-Null
    $noteRng.Value2 = "Pour le détail compte par compte de chaque cycle : onglet 'Données', filtrez la colonne Cycle. Pour un PDF : Fichier -> Exporter -> Créer un PDF/XPS."
    Set-CellStyle -Range $noteRng -FontColor '6c757d' -Size 9 -HAlign 'left' -VAlign 'top' -Wrap $true
    $r += 3

    $Sheet.PageSetup.PrintArea = "`$A`$1:`$M`$$r"
}
