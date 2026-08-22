# Construit la feuille Rapport_Trame (indicateurs + préconisations)
# Dot-sourcé par Build-Workbook.ps1
# Classeur autonome, sans macro : les champs "Client / Profil / Résumé du bilan" sont saisis
# manuellement par l'expert-comptable à partir de ce qu'il a obtenu dans les deux autres
# outils. Les missions sont des menus déroulants Oui/Non ; le budget total est une formule
# (SUMPRODUCT-like), recalculée en direct - fonctionne sur PC, mobile et Excel Online.

function New-RapportTrameSheet {
    param($Sheet, $Workbook, [string]$RefDataPath,
          [array]$NavButtons = @(), [string]$NavTitle = 'Trame de rapport', [bool]$Unified = $false)

    Add-NavBar -Sheet $Sheet -Workbook $Workbook -Title $NavTitle -Buttons $NavButtons -ActiveKey 'trame'

    for ($c = 2; $c -le 13; $c++) { $Sheet.Cells.Item(1, $c).ColumnWidth = 9.5 }

    Add-SectionHeader -Sheet $Sheet -CellAddr 'B4' -Text 'Trame de rapport et préconisations' -MergeCols 12 | Out-Null

    # --- Synthèse profilage (saisie manuelle, à reporter depuis Outil-de-Profilage.xlsx) ---
    $Sheet.Cells.Item(6,2).Value2 = 'Client (profilage) :'
    Set-CellStyle -Range $Sheet.Cells.Item(6,2) -Bold $true -HAlign 'left'
    $pc = $Sheet.Range('D6:F6'); $pc.Merge() | Out-Null
    $pc.Interior.Color = ConvertTo-OleColor $Script:Palette.White
    $pc.Borders.Color = ConvertTo-OleColor $Script:Palette.Border
    $pc.Locked = $false
    # INDEX(Nom,1,1) et non "=Res_ClientName" : Res_ClientName pointe vers une plage fusionnée
    # multi-cellules (D31:G31 sur Profilage_Resultats), et cette formule vit sur une AUTRE ligne -
    # une référence directe au nom subirait l'intersection implicite d'Excel et renverrait
    # #VALEUR! (même piège que documenté dans New-ProfilageSheets.ps1 pour Res_Part1 etc.).
    if ($Unified) { $pc.Formula = '=INDEX(Res_ClientName,1,1)' }
    New-DefinedName -Workbook $Workbook -Name 'Rap_ProfilClient' -RangeObj $pc

    $Sheet.Cells.Item(6,8).Value2 = 'Profil :'
    Set-CellStyle -Range $Sheet.Cells.Item(6,8) -Bold $true -HAlign 'left'
    $pl = $Sheet.Range('I6:J6'); $pl.Merge() | Out-Null
    $pl.Interior.Color = ConvertTo-OleColor $Script:Palette.White
    $pl.Borders.Color = ConvertTo-OleColor $Script:Palette.Border
    $pl.Locked = $false
    $pl.Validation.Delete()
    $pl.Validation.Add(3, 1, 1, "Proactif,Ouvert,Prudent / Intermédiaire,Réticent") | Out-Null
    $pl.Validation.InCellDropdown = $true
    if ($Unified) { $pl.Formula = '=Res_ProfilKey' } # nom d'ancrage mono-cellule (Q15) - référence directe sûre
    New-DefinedName -Workbook $Workbook -Name 'Rap_ProfilLabel' -RangeObj $pl

    $Sheet.Cells.Item(6,11).Value2 = 'Score :'
    Set-CellStyle -Range $Sheet.Cells.Item(6,11) -Bold $true -HAlign 'left'
    $ps = $Sheet.Cells.Item(6,12)
    # Format Texte SEULEMENT en saisie manuelle (évite qu'Excel interprète "1/30" comme une date) :
    # si on le pose avant d'assigner une .Formula (cas $Unified plus bas), Excel affiche la formule
    # telle quelle au lieu de la calculer - un cell pré-formaté Texte ne calcule pas les formules
    # qu'on y assigne ensuite par COM/VBA.
    if (-not $Unified) { $ps.NumberFormat = '@' }
    $ps.Interior.Color = ConvertTo-OleColor $Script:Palette.White
    $ps.Borders.Color = ConvertTo-OleColor $Script:Palette.Border
    $ps.Locked = $false
    $ps.HorizontalAlignment = -4108
    if ($Unified) { $ps.Formula = '=IF(Res_ScoreFinal="","",Res_ScoreFinal&"/30")' } # Res_ScoreFinal : nom mono-cellule (Q14)
    New-DefinedName -Workbook $Workbook -Name 'Rap_ProfilScore' -RangeObj $ps

    # Sous-scores Partie 1 (État des lieux, /12) et Partie 3 (Perception, /20) - affichés dans le
    # rapport final du site (E8.blade.php: colonnes "État des lieux" / "Perception", à côté du
    # score total /30) mais absents du classeur jusqu'ici, qui ne reprenait que le score total.
    $Sheet.Cells.Item(7,2).Value2 = 'État des lieux (Partie 1) :'
    Set-CellStyle -Range $Sheet.Cells.Item(7,2) -Bold $true -HAlign 'left'
    $p1 = $Sheet.Cells.Item(7,4)
    if (-not $Unified) { $p1.NumberFormat = '@' } # cf. note ci-dessus sur Rap_ProfilScore
    $p1.Interior.Color = ConvertTo-OleColor $Script:Palette.White
    $p1.Borders.Color = ConvertTo-OleColor $Script:Palette.Border
    $p1.Locked = $false
    $p1.HorizontalAlignment = -4108
    if ($Unified) { $p1.Formula = '=IF(Res_Part1="","",Res_Part1&"/12")' } # Res_Part1 : nom mono-cellule (G6 Profilage_Resultats)
    New-DefinedName -Workbook $Workbook -Name 'Rap_Part1Score' -RangeObj $p1

    $Sheet.Cells.Item(7,8).Value2 = 'Perception (Partie 3) :'
    Set-CellStyle -Range $Sheet.Cells.Item(7,8) -Bold $true -HAlign 'left'
    $p3 = $Sheet.Cells.Item(7,11)
    if (-not $Unified) { $p3.NumberFormat = '@' }
    $p3.Interior.Color = ConvertTo-OleColor $Script:Palette.White
    $p3.Borders.Color = ConvertTo-OleColor $Script:Palette.Border
    $p3.Locked = $false
    $p3.HorizontalAlignment = -4108
    if ($Unified) { $p3.Formula = '=IF(Res_Part3="","",Res_Part3&"/20")' } # Res_Part3 : nom mono-cellule (G10 Profilage_Resultats)
    New-DefinedName -Workbook $Workbook -Name 'Rap_Part3Score' -RangeObj $p3

    # --- Synthèse bilan écologique (saisie manuelle, à reporter depuis Integrateur-de-Balance.xlsm) ---
    $Sheet.Cells.Item(8,2).Value2 = 'Client (balance) :'
    Set-CellStyle -Range $Sheet.Cells.Item(8,2) -Bold $true -HAlign 'left'
    $bc = $Sheet.Range('D8:F8'); $bc.Merge() | Out-Null
    $bc.Interior.Color = ConvertTo-OleColor $Script:Palette.White
    $bc.Borders.Color = ConvertTo-OleColor $Script:Palette.Border
    $bc.Locked = $false
    # INDEX(Nom,1,1) : Bal_ClientName pointe vers D7:G7 fusionné sur Balance_Import (multi-cellules) -
    # même précaution que Rap_ProfilClient ci-dessus contre l'intersection implicite.
    if ($Unified) { $bc.Formula = '=INDEX(Bal_ClientName,1,1)' }
    New-DefinedName -Workbook $Workbook -Name 'Rap_BalanceClient' -RangeObj $bc

    $brRng = $Sheet.Range('B9:M10'); $brRng.Merge() | Out-Null
    $brRng.Interior.Color = ConvertTo-OleColor $Script:Palette.White
    $brRng.Borders.Color = ConvertTo-OleColor $Script:Palette.Border
    $brRng.VerticalAlignment = -4160
    $brRng.WrapText = $true
    $brRng.Locked = $false
    New-DefinedName -Workbook $Workbook -Name 'Rap_BalanceResume' -RangeObj $brRng

    $noteRng = $Sheet.Range('B11:M11'); $noteRng.Merge() | Out-Null
    $noteRng.Value2 = if ($Unified) {
        'Client, profil et score sont repris automatiquement des onglets Profilage et Balance (modifiables ici si besoin). Détail compte par compte : onglet Analyse.'
    } else {
        'Champs à recopier depuis les résultats des deux autres classeurs (Outil de profilage, Intégrateur de balance).'
    }
    Set-CellStyle -Range $noteRng -FontColor '6c757d' -Size 9 -HAlign 'left'

    # --- Indicateurs environnementaux ---
    # Colonnes alignées sur E8.blade.php:43-48 (Indicateur Clé / Valeur actuelle / Objectif /
    # Réalisation / Tendance / Dernière mise à jour) - pas de colonne "Unité" séparée : le site
    # n'en affiche jamais une (l'unité, quand elle apparaît, est incluse dans le nom lui-même).
    Add-SectionHeader -Sheet $Sheet -CellAddr 'B12' -Text 'Indicateurs environnementaux' -MergeCols 12 | Out-Null
    $indHeaders = @('Indicateur','Valeur actuelle','Objectif','Réalisation','Tendance','Dernière mise à jour')
    for ($c = 0; $c -lt $indHeaders.Count; $c++) {
        $cell = $Sheet.Cells.Item(14, 2 + $c)
        $cell.Value2 = $indHeaders[$c]
        Set-CellStyle -Range $cell -Bold $true -Fill $Script:Palette.CardGray -HAlign 'left'
    }
    $indicateurs = (Import-PsdSafe (Join-Path $RefDataPath 'Indicateurs.psd1')).Indicateurs
    $r = 15
    foreach ($ind in $indicateurs) {
        $Sheet.Cells.Item($r,2).Value2 = $ind.Nom
        foreach ($col in 3,4,5) {
            $cell = $Sheet.Cells.Item($r, $col)
            $cell.Interior.Color = ConvertTo-OleColor $Script:Palette.White
            $cell.Borders.Color = ConvertTo-OleColor $Script:Palette.Border
            $cell.Locked = $false
        }
        $trend = $Sheet.Cells.Item($r, 6)
        $trend.Interior.Color = ConvertTo-OleColor $Script:Palette.White
        $trend.Borders.Color = ConvertTo-OleColor $Script:Palette.Border
        $trend.Locked = $false
        $trend.Validation.Delete()
        $trend.Validation.Add(3, 1, 1, "Hausse,Baisse,Stable") | Out-Null
        $trend.Validation.InCellDropdown = $true
        $lastUpdate = $Sheet.Cells.Item($r, 7)
        $lastUpdate.Interior.Color = ConvertTo-OleColor $Script:Palette.White
        $lastUpdate.Borders.Color = ConvertTo-OleColor $Script:Palette.Border
        $lastUpdate.Locked = $false
        # Barres obliques échappées : "dd/mm/yyyy" (non échappé) est rejeté par COM sur cette
        # installation (Excel.Application lève "Impossible de définir la propriété NumberFormat")
        # dès que le jour précède le mois avec un "/" non échappé - confirmé empiriquement
        # ("m/d/yyyy" passe, "dd/mm/yyyy" échoue, "dd\/mm\/yyyy" passe et affiche pareil).
        $lastUpdate.NumberFormat = 'dd\/mm\/yyyy'
        $lastUpdate.HorizontalAlignment = -4108
        $lastUpdate.Formula = '=TODAY()' # même défaut que le site (date('Y-m-d') si non renseigné) - modifiable en tapant par-dessus
        $r++
    }

    # --- Préconisations : sélection de missions (menus déroulants Oui/Non, sans macro) ---
    Add-SectionHeader -Sheet $Sheet -CellAddr 'B22' -Text 'Préconisations - sélection de missions' -MergeCols 12 | Out-Null
    $missions = (Import-PsdSafe (Join-Path $RefDataPath 'MissionTypes.psd1')).Missions
    $r = 24
    $mIdx = 1
    foreach ($m in $missions) {
        $lbl = $Sheet.Range($Sheet.Cells.Item($r,2), $Sheet.Cells.Item($r,4)); $lbl.Merge() | Out-Null
        $lbl.Value2 = "$($m.Name) - $($m.BasePrice) EUR"
        Set-CellStyle -Range $lbl -FontColor $Script:Palette.DarkText -HAlign 'left'

        $sel = $Sheet.Cells.Item($r, 5)
        $sel.Interior.Color = ConvertTo-OleColor $Script:Palette.White
        $sel.Borders.Color = ConvertTo-OleColor $Script:Palette.Border
        $sel.Locked = $false
        $sel.HorizontalAlignment = -4108
        $sel.Value2 = 'Non'
        $sel.Validation.Delete()
        $sel.Validation.Add(3, 1, 1, "Oui,Non") | Out-Null
        $sel.Validation.InCellDropdown = $true
        New-DefinedName -Workbook $Workbook -Name "Rap_Mission$mIdx" -RangeObj $sel

        $descRng = $Sheet.Range($Sheet.Cells.Item($r,7), $Sheet.Cells.Item($r,13)); $descRng.Merge() | Out-Null
        $descRng.Value2 = $m.Description
        Set-CellStyle -Range $descRng -FontColor '6c757d' -Size 9 -HAlign 'left'
        $Sheet.Rows.Item($r).RowHeight = 18
        $r++
        $mIdx++
    }

    $r++
    $Sheet.Cells.Item($r,2).Value2 = 'Budget total estimé :'
    Set-CellStyle -Range $Sheet.Cells.Item($r,2) -Bold $true -HAlign 'left'
    $budgetCell = $Sheet.Cells.Item($r, 4)
    Set-CellStyle -Range $budgetCell -Fill $Script:Palette.Green -FontColor $Script:Palette.White -Bold $true -HAlign 'center'
    New-DefinedName -Workbook $Workbook -Name 'Rap_BudgetTotal' -RangeObj $budgetCell

    # Formule du budget total : identique à base_price * (1 + nb_services * 0.2) par mission
    # cochée, additionné - générée pour le nombre exact de missions du référentiel.
    $terms = for ($i = 1; $i -le $missions.Count; $i++) {
        "IF(Rap_Mission$i=""Oui"",INDEX(Ref_Missions,$i,4)*(1+(LEN(INDEX(Ref_Missions,$i,5))-LEN(SUBSTITUTE(INDEX(Ref_Missions,$i,5),""|"",""""))+1)*0.2),0)"
    }
    # "# ##0" (espace) et non "#,##0" : sous Excel en locale française, le code de format de
    # TEXT() suit la convention locale (séparateur de milliers = espace) - cf. note similaire
    # dans New-ProfilageSheets.ps1.
    $budgetCell.Formula = '=TEXT(' + ($terms -join '+') + ',"# ##0")&" EUR"'
    $r += 2

    $exportNoteRng = $Sheet.Range($Sheet.Cells.Item($r,2), $Sheet.Cells.Item($r+1,13)); $exportNoteRng.Merge() | Out-Null
    $exportNoteRng.Value2 = 'Pour un PDF : Fichier -> Exporter -> Créer un PDF/XPS.'
    Set-CellStyle -Range $exportNoteRng -FontColor '6c757d' -Size 9 -HAlign 'left' -VAlign 'top' -Wrap $true
    $r += 2

    $Sheet.PageSetup.PrintArea = "`$A`$1:`$M`$$($r + 2)"
}
