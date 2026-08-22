# Construit la feuille Balance_Data (données + classification automatique par formules)
# et Profils_Historique (repère manuel, sans macro).
# Dot-sourcé par Build-Workbook.ps1.

function New-BalanceDataSheet {
    <#
        Zone de collage manuel (Compte / Libellé / Année N-1 / Année N, colonnes B:E) +
        colonnes calculées par formule (Variation, Cycle, Lecture écologique, Question,
        Suggestion) - remplies pour 1000 lignes, qu'elles soient renseignées par le bouton
        d'import (macro, PC uniquement) ou collées à la main (fonctionne partout, y compris
        mobile). La classification (cycle comptable, lecture écologique) est une chaîne de
        SI imbriqués générée depuis le référentiel, dans l'ordre exact de priorité métier -
        aucune macro requise pour cette partie.
    #>
    param($Sheet, $Workbook, [string]$RefDataPath,
          [array]$NavButtons = $Script:BalanceNavButtons, [string]$NavTitle = 'Intégrateur de balance')

    Add-NavBar -Sheet $Sheet -Workbook $Workbook -Title $NavTitle -Buttons $NavButtons -ActiveKey 'donnees'
    for ($c = 2; $c -le 12; $c++) { $Sheet.Cells.Item(1, $c).ColumnWidth = 13 }

    Add-SectionHeader -Sheet $Sheet -CellAddr 'B4' -Text 'Données de balance' -MergeCols 10 | Out-Null

    $noteRng = $Sheet.Range('B6:K7'); $noteRng.Merge() | Out-Null
    $noteRng.Value2 = "Collez ici vos comptes dans les colonnes Compte / Libellé / Année N-1 / Année N (ou utilisez le bouton `"Choisir un fichier...`" sur l'onglet Import). Les colonnes Variation, Cycle, Lecture écologique, Question et Suggestion se remplissent automatiquement. Collez les numéros de compte en texte (pas en nombre) pour conserver les zéros de tête."
    Set-CellStyle -Range $noteRng -Fill $Script:Palette.LightGray -FontColor '495057' -Size 9.5 -HAlign 'left' -VAlign 'top' -Wrap $true

    $headerRow = 9
    $headers = @('Compte','Libellé','Année N-1','Année N','Variation (€)','Variation (%)','Cycle','Lecture écologique','Question à poser','Suggestion')
    for ($c = 0; $c -lt $headers.Count; $c++) {
        $cell = $Sheet.Cells.Item($headerRow, 2 + $c)
        $cell.Value2 = $headers[$c]
        Set-CellStyle -Range $cell -Fill $Script:Palette.Green -FontColor $Script:Palette.White -Bold $true -HAlign 'left'
    }

    $firstDataRow = $headerRow + 1
    $maxRows = 1000
    $lastDataRow = $firstDataRow + $maxRows - 1

    # Colonnes de saisie (B:E) - déverrouillées, saisissables/collables à la main.
    $rawRng = $Sheet.Range($Sheet.Cells.Item($firstDataRow,2), $Sheet.Cells.Item($lastDataRow,5))
    $rawRng.Locked = $false
    $Sheet.Range($Sheet.Cells.Item($firstDataRow,2), $Sheet.Cells.Item($lastDataRow,2)).NumberFormat = '@' # Compte en texte

    # --- Génération des chaînes de SI imbriqués (ordre de priorité = référentiel) ---
    $cycleRulesOrdered = (Import-PsdSafe (Join-Path $RefDataPath 'CycleRules.psd1'))
    # Sort-Object avec un nom de propriété littéral ne trie pas correctement un tableau de
    # Hashtable (type produit par Import-PowerShellDataFile) - bloc de script requis (cf. note
    # similaire dans New-Referentiel.ps1).
    $cycleChainRules = $cycleRulesOrdered.Rules | Sort-Object { $_.Order } | ForEach-Object { [pscustomobject]@{ Prefix = $_.Prefix; Value = $_.Cycle } }
    $accountRulesOrdered = (Import-PsdSafe (Join-Path $RefDataPath 'AccountRules.psd1')).Rules | Sort-Object { $_.Order }
    $ruleOrdChainRules = $accountRulesOrdered | ForEach-Object { [pscustomobject]@{ Prefix = $_.Prefix; Value = $_.Order } }

    $cycleFormulaExpr = Build-PrefixChainFormula -CellRef "`$B$firstDataRow" -Rules $cycleChainRules -Default $cycleRulesOrdered.DefaultCycle
    $ruleOrdFormulaExpr = Build-PrefixChainFormula -CellRef "`$B$firstDataRow" -Rules $ruleOrdChainRules -Default '""' -ValueIsLiteral $true

    $r = $firstDataRow
    $Sheet.Cells.Item($r,6).Formula  = "=IF(`$B$r=`"`",`"`",E$r-D$r)"
    $Sheet.Cells.Item($r,7).Formula  = "=IF(`$B$r=`"`",`"`",IF(D$r=0,0,(E$r-D$r)/ABS(D$r)*100))"
    $Sheet.Cells.Item($r,12).Formula = "=IF(`$B$r=`"`",`"`",$ruleOrdFormulaExpr)"
    # Cycle conditionné à `$L$r (règle de lecture écologique appariée), PAS à `$B$r seul : sur le
    # site, E5Controller::detectCycle() n'est appelé QUE dans generateAnalysisForLine() une fois
    # qu'une règle AccountRules a matché (cf. E5Controller.php ~391-415) - un compte qui ne
    # correspond à aucune des 11 règles de lecture n'est jamais classé en cycle côté site, même
    # s'il matche par ailleurs un préfixe CycleRules. Sans cette garde, l'Excel classait TOUS les
    # comptes (bien plus permissif que le site).
    $Sheet.Cells.Item($r,8).Formula  = "=IF(`$L$r=`"`",`"`",$cycleFormulaExpr)"
    $Sheet.Cells.Item($r,9).Formula  = "=IF(`$L$r=`"`",`"`",INDEX(Ref_AccountRules,`$L$r,3))"
    $Sheet.Cells.Item($r,10).Formula = "=IF(`$L$r=`"`",`"`",INDEX(Ref_AccountRules,`$L$r,4))"
    $Sheet.Cells.Item($r,11).Formula = "=IF(`$L$r=`"`",`"`",INDEX(Ref_AccountRules,`$L$r,5))"

    foreach ($col in 6,7,8,9,10,11,12) {
        $Sheet.Range($Sheet.Cells.Item($firstDataRow,$col), $Sheet.Cells.Item($lastDataRow,$col)).FillDown() | Out-Null
    }

    # Noms définis sur les colonnes calculées - utilisés par les formules de Balance_Analyse.
    $namesByCol = @{ 2='BalData_Compte'; 3='BalData_Libelle'; 6='BalData_VarAbs'; 7='BalData_VarPct'; 8='BalData_Cycle'; 9='BalData_Lecture'; 11='BalData_Suggestion' }
    foreach ($col in $namesByCol.Keys) {
        $rng = $Sheet.Range($Sheet.Cells.Item($firstDataRow,$col), $Sheet.Cells.Item($lastDataRow,$col))
        New-DefinedName -Workbook $Workbook -Name $namesByCol[$col] -RangeObj $rng
    }

    # Colonne L (12) = ordre de règle apparié, aide au calcul uniquement - masquée.
    $Sheet.Columns.Item(12).Hidden = $true
    $Sheet.Columns.Item(1).ColumnWidth = 2

    New-DefinedName -Workbook $Workbook -Name 'BalanceData_FirstRow' -RangeObj $Sheet.Cells.Item($firstDataRow, 2)
}

function New-ProfilsHistoriqueSheet {
    param($Sheet, $Workbook)
    $headers = @('Date','ClientName','TotalScore','Profile','Part1','Part2','Part3','Part4','Coefficient','ExpertNotes')
    for ($c = 0; $c -lt $headers.Count; $c++) {
        $Sheet.Cells.Item(1, $c + 1).Value2 = $headers[$c]
        $Sheet.Cells.Item(1, $c + 1).Font.Bold = $true
    }
    $hdrRng = $Sheet.Range($Sheet.Cells.Item(1,1), $Sheet.Cells.Item(1,$headers.Count))
    New-DefinedName -Workbook $Workbook -Name 'ProfilsHistorique_Header' -RangeObj $hdrRng
    $Sheet.Visible = 2
}
