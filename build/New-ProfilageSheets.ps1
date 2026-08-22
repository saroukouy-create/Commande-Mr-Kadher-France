# Construit les feuilles Profilage_Questionnaire et Profilage_Resultats
# Dot-sourcé par Build-Workbook.ps1
#
# Sans macro : navigation par liens hypertexte natifs, calculs par formules (SUMPRODUCT/LOOKUP)
# qui se recalculent en direct dès qu'une réponse change. Fonctionne sur PC, Mac, mobile, Excel Online.

$Script:PartTitles = @{
    1 = "Partie 1 – État des lieux de la démarche environnementale"
    2 = "Partie 2 – Freins, motivations et niveau d’intérêt"
    3 = "Partie 3 – Perception du rôle du cabinet comptable"
    4 = "Partie 4 – Niveau d’intérêt pour un accompagnement"
}

function Get-CellAddress {
    param($Cell)
    return $Cell.Address($true, $true, 1, $true)
}

$Script:ProfilageNavButtons = @(
    @{ Key='questionnaire'; Label='Questionnaire'; TargetSheet='Profilage_Questionnaire'; TargetCell='B4' }
    @{ Key='resultats';     Label='Résultats';      TargetSheet='Profilage_Resultats';     TargetCell='B4' }
)

function New-ProfilageQuestionnaireSheet {
    param($Sheet, $Workbook, [string]$RefDataPath,
          [array]$NavButtons = $Script:ProfilageNavButtons, [string]$NavTitle = 'Outil de profilage')

    Add-NavBar -Sheet $Sheet -Workbook $Workbook -Title $NavTitle -Buttons $NavButtons -ActiveKey 'questionnaire'
    $questions = (Import-PsdSafe (Join-Path $RefDataPath 'ProfilQuestions.psd1')).Questions

    for ($c = 2; $c -le 13; $c++) { $Sheet.Cells.Item(1, $c).ColumnWidth = 9.5 }

    $Sheet.Range('B4:M4').Merge() | Out-Null
    $Sheet.Range('B4').Value2 = 'Outil de profilage client'
    Set-CellStyle -Range $Sheet.Range('B4:M4') -FontColor $Script:Palette.DarkText -Bold $true -Size 16 -HAlign 'left'
    $Sheet.Rows.Item(4).RowHeight = 26

    $row = 6
    $currentPart = 0
    $checkboxLinkCol = 16 # colonne P, hors zone visible du formulaire

    foreach ($q in $questions) {
        if ($q.Part -ne $currentPart) {
            $currentPart = $q.Part
            $partRng = $Sheet.Range($Sheet.Cells.Item($row,2), $Sheet.Cells.Item($row,13))
            $partRng.Merge() | Out-Null
            $partRng.Value2 = $Script:PartTitles[$currentPart]
            Set-CellStyle -Range $partRng -Fill $Script:Palette.Green -FontColor $Script:Palette.White -Bold $true -Size 11 -HAlign 'left'
            $Sheet.Rows.Item($row).RowHeight = 22
            $row += 2
        }

        $qRng = $Sheet.Range($Sheet.Cells.Item($row,2), $Sheet.Cells.Item($row,13))
        $qRng.Merge() | Out-Null
        $qRng.Value2 = $q.Text
        Set-CellStyle -Range $qRng -Fill $Script:Palette.CardGray -FontColor $Script:Palette.DarkText -Bold $true -Size 10.5 -HAlign 'left' -Wrap $true
        $Sheet.Rows.Item($row).RowHeight = 30
        $row++

        if ($q.Description) {
            $dRng = $Sheet.Range($Sheet.Cells.Item($row,2), $Sheet.Cells.Item($row,13))
            $dRng.Merge() | Out-Null
            $dRng.Value2 = $q.Description
            Set-CellStyle -Range $dRng -FontColor '6c757d' -Size 9 -HAlign 'left'
            $Sheet.Rows.Item($row).RowHeight = 16
            $row++
        }

        switch ($q.Type) {
            'radio' {
                $lbl = $Sheet.Cells.Item($row, 3); $lbl.Value2 = 'Votre réponse :'
                Set-CellStyle -Range $lbl -FontColor $Script:Palette.DarkText -HAlign 'left'
                $ans = $Sheet.Range($Sheet.Cells.Item($row,5), $Sheet.Cells.Item($row,10))
                $ans.Merge() | Out-Null
                $ans.Interior.Color = ConvertTo-OleColor $Script:Palette.White
                $ans.Borders.Color = ConvertTo-OleColor $Script:Palette.Border
                $ans.Locked = $false
                $Sheet.Rows.Item($row).RowHeight = 20
                $listName = "Ref_Q$($q.Id)_OptionList"
                $ans.Validation.Delete()
                $ans.Validation.Add(3, 1, 1, "=$listName") | Out-Null # 3=xlValidateList, 1=xlValidAlertStop
                $ans.Validation.InCellDropdown = $true
                # Nom défini sur la SEULE cellule d'ancrage (première cellule de la fusion),
                # pas sur toute la plage fusionnée - cf. note sur l'intersection implicite plus bas.
                New-DefinedName -Workbook $Workbook -Name "Q$($q.Id)_Answer" -RangeObj $Sheet.Cells.Item($row,5)
                $row += 2
            }
            'checkbox' {
                $blockStartRow = $row
                foreach ($opt in $q.Options) {
                    $linkCell = $Sheet.Cells.Item($row, $checkboxLinkCol)
                    $linkCell.Value2 = $false
                    $linkCell.Locked = $false # sinon Excel refuse de cocher la case sur une feuille protégée
                    $linkAddr = Get-CellAddress $linkCell
                    $cb = $Sheet.CheckBoxes().Add($Sheet.Cells.Item($row,4).Left, $Sheet.Cells.Item($row,4).Top, 420, 18)
                    $cb.Caption = $opt.Text
                    $cb.LinkedCell = $linkAddr # cellule liée native - aucune macro requise
                    $cb.Value = -4146 # xlOff
                    $Sheet.Rows.Item($row).RowHeight = 18
                    $row++
                }
                # Plage couvrant tout le bloc de cases à cocher - utilisée par une formule COUNTIF
                # sur la feuille Résultats (compte des cases cochées), sans macro.
                $blockRng = $Sheet.Range($Sheet.Cells.Item($blockStartRow, $checkboxLinkCol), $Sheet.Cells.Item($row - 1, $checkboxLinkCol))
                New-DefinedName -Workbook $Workbook -Name "Q$($q.Id)_Opts" -RangeObj $blockRng
                $row++
            }
            'rating' {
                $blockStartRow = $row
                foreach ($subj in $q.Subjects) {
                    $sLbl = $Sheet.Range($Sheet.Cells.Item($row,3), $Sheet.Cells.Item($row,9))
                    $sLbl.Merge() | Out-Null
                    $sLbl.Value2 = $subj
                    Set-CellStyle -Range $sLbl -FontColor $Script:Palette.DarkText -HAlign 'left' -Wrap $true
                    $rate = $Sheet.Cells.Item($row, 11)
                    $rate.Interior.Color = ConvertTo-OleColor $Script:Palette.White
                    $rate.Borders.Color = ConvertTo-OleColor $Script:Palette.Border
                    $rate.HorizontalAlignment = -4108
                    $rate.Locked = $false
                    $Sheet.Rows.Item($row).RowHeight = 20
                    $rate.Validation.Delete()
                    $rate.Validation.Add(3, 1, 1, "1,2,3,4") | Out-Null
                    $rate.Validation.InCellDropdown = $true
                    $row++
                }
                # Plage couvrant les 5 notations - utilisée par une formule SUM sur Résultats.
                $blockRng = $Sheet.Range($Sheet.Cells.Item($blockStartRow, 11), $Sheet.Cells.Item($row - 1, 11))
                New-DefinedName -Workbook $Workbook -Name 'Q7_Scores' -RangeObj $blockRng
                $row++
            }
        }
        $row++
    }

    Add-NavHyperlink -Sheet $Sheet -Workbook $Workbook -Left ($Sheet.Cells.Item($row,5).Left) -Top ($Sheet.Cells.Item($row,5).Top) `
        -Width 260 -Height 38 -Label 'Voir les résultats →' -TargetSheet 'Profilage_Resultats' -TargetCell 'B4' `
        -Fill $Script:Palette.Success -TextColor $Script:Palette.White -FontSize 11 | Out-Null

    $noteRng = $Sheet.Range($Sheet.Cells.Item($row + 2, 2), $Sheet.Cells.Item($row + 2, 10))
    $noteRng.Merge() | Out-Null
    $noteRng.Value2 = 'Le résultat se met à jour automatiquement dès que vous répondez - pas besoin de valider.'
    Set-CellStyle -Range $noteRng -FontColor '6c757d' -Size 9 -HAlign 'left'

    # La colonne P (16) porte les cellules liées des cases à cocher (hors zone visible du formulaire) -
    # on la masque et on borne la zone d'impression avant elle pour ne pas afficher VRAI/FAUX.
    $Sheet.Columns.Item(15).ColumnWidth = 2
    $Sheet.Columns.Item(16).Hidden = $true
    $Sheet.PageSetup.PrintArea = "`$A`$1:`$N`$$($row + 3)"
}

function New-ProfilageResultatsSheet {
    param($Sheet, $Workbook,
          [array]$NavButtons = $Script:ProfilageNavButtons, [string]$NavTitle = 'Outil de profilage', [bool]$Unified = $false)

    Add-NavBar -Sheet $Sheet -Workbook $Workbook -Title $NavTitle -Buttons $NavButtons -ActiveKey 'resultats'
    for ($c = 2; $c -le 13; $c++) { $Sheet.Cells.Item(1, $c).ColumnWidth = 9.5 }

    Add-SectionHeader -Sheet $Sheet -CellAddr 'B4' -Text 'Résultats du questionnaire' -MergeCols 12 | Out-Null

    # --- Détail du calcul (toutes les cellules "valeur" ci-dessous sont des FORMULES,
    #     recalculées en direct depuis Profilage_Questionnaire - aucune macro). ---
    $labels = @(
        @{ R=6;  L='Partie 1 (État des lieux) /12 :';       N='Res_Part1' }
        @{ R=7;  L='Partie 2 (Freins/Motivations) /10 :';   N='Res_Part2' }
        @{ R=8;  L='Partie 4 (Accompagnement) /8 :';        N='Res_Part4' }
        @{ R=9;  L='Total Brut /30 :';                      N='Res_Brut' }
        @{ R=10; L='Score Partie 3 - Perception /20 :';     N='Res_Part3' }
        @{ R=11; L='Niveau de perception :';                N='Res_Niveau' }
        @{ R=12; L='Coefficient appliqué :';                N='Res_Coefficient' }
    )
    foreach ($item in $labels) {
        $lc = $Sheet.Range($Sheet.Cells.Item($item.R,2), $Sheet.Cells.Item($item.R,6)); $lc.Merge() | Out-Null
        $lc.Value2 = $item.L
        Set-CellStyle -Range $lc -FontColor $Script:Palette.DarkText -HAlign 'left'
        $vc = $Sheet.Range($Sheet.Cells.Item($item.R,7), $Sheet.Cells.Item($item.R,9)); $vc.Merge() | Out-Null
        Set-CellStyle -Range $vc -Fill $Script:Palette.CardGray -FontColor $Script:Palette.DarkText -Bold $true -HAlign 'center'
        # Nom défini sur la SEULE cellule d'ancrage (G), pas sur toute la plage fusionnée G:I :
        # une formule d'une AUTRE ligne référençant un nom pointant vers une plage fusionnée
        # multi-cellules subit l'intersection implicite d'Excel et renvoie #VALEUR!.
        New-DefinedName -Workbook $Workbook -Name $item.N -RangeObj $Sheet.Cells.Item($item.R,7)
    }

    $finalRng = $Sheet.Range('B14:I15'); $finalRng.Merge() | Out-Null
    Set-CellStyle -Range $finalRng -Fill 'e9ecef' -FontColor $Script:Palette.Green -Bold $true -Size 16 -HAlign 'center'
    New-DefinedName -Workbook $Workbook -Name 'Res_FormuleFinale' -RangeObj $finalRng

    # Cellules techniques (formules intermédiaires, hors zone imprimée)
    New-DefinedName -Workbook $Workbook -Name 'Res_ScoreFinal' -RangeObj $Sheet.Cells.Item(14, 17)
    New-DefinedName -Workbook $Workbook -Name 'Res_ProfilKey' -RangeObj $Sheet.Cells.Item(15, 17)
    New-DefinedName -Workbook $Workbook -Name 'Res_CoefBase' -RangeObj $Sheet.Cells.Item(16, 17)
    New-DefinedName -Workbook $Workbook -Name 'Res_CoefRaw' -RangeObj $Sheet.Cells.Item(17, 17)

    # --- Profil ---
    $profRng = $Sheet.Range('B17:M17'); $profRng.Merge() | Out-Null
    Set-CellStyle -Range $profRng -Fill $Script:Palette.Green -FontColor $Script:Palette.White -Bold $true -Size 14 -HAlign 'center'
    $Sheet.Rows.Item(17).RowHeight = 28
    New-DefinedName -Workbook $Workbook -Name 'Res_ProfilTitre' -RangeObj $profRng

    $posRng = $Sheet.Range('B18:M19'); $posRng.Merge() | Out-Null
    Set-CellStyle -Range $posRng -Fill $Script:Palette.LightGray -FontColor $Script:Palette.DarkText -HAlign 'left' -Wrap $true
    New-DefinedName -Workbook $Workbook -Name 'Res_Positionnement' -RangeObj $posRng

    $freinsHdr = $Sheet.Range('B21:F21'); $freinsHdr.Merge() | Out-Null
    # Le site affiche "Freins potentiels :" pour le profil Proactif, "Freins à lever :" pour les 3 autres.
    $freinsHdr.Formula = '=IF(Res_ProfilKey="Proactif","Freins potentiels :","Freins à lever :")'
    Set-CellStyle -Range $freinsHdr -Bold $true -FontColor $Script:Palette.DarkText -HAlign 'left'
    $freinsRng = $Sheet.Range('B22:F25'); $freinsRng.Merge() | Out-Null
    Set-CellStyle -Range $freinsRng -FontColor $Script:Palette.DarkText -HAlign 'left' -VAlign 'top' -Wrap $true
    New-DefinedName -Workbook $Workbook -Name 'Res_Freins' -RangeObj $freinsRng

    $argHdr = $Sheet.Range('H21:M21'); $argHdr.Merge() | Out-Null
    $argHdr.Value2 = 'Arguments à utiliser :'
    Set-CellStyle -Range $argHdr -Bold $true -FontColor $Script:Palette.DarkText -HAlign 'left'
    $argRng = $Sheet.Range('H22:M25'); $argRng.Merge() | Out-Null
    Set-CellStyle -Range $argRng -FontColor $Script:Palette.DarkText -HAlign 'left' -VAlign 'top' -Wrap $true
    New-DefinedName -Workbook $Workbook -Name 'Res_Arguments' -RangeObj $argRng

    $tonHdr = $Sheet.Range('B27:M27'); $tonHdr.Merge() | Out-Null
    $tonHdr.Value2 = 'Ton recommandé :'
    Set-CellStyle -Range $tonHdr -Bold $true -FontColor $Script:Palette.DarkText -HAlign 'left'
    $tonRng = $Sheet.Range('B28:M29'); $tonRng.Merge() | Out-Null
    Set-CellStyle -Range $tonRng -Fill $Script:Palette.LightGray -FontColor $Script:Palette.DarkText -HAlign 'left' -VAlign 'top' -Wrap $true
    New-DefinedName -Workbook $Workbook -Name 'Res_Ton' -RangeObj $tonRng

    # Encart d'alerte visible uniquement pour le profil Proactif - cf. E4.blade.php:270-274.
    # Toujours présent (cellule vide sinon) : sans macro, on ne peut pas masquer une ligne selon
    # une condition, seulement vider son contenu par formule - même technique que le libellé
    # Freins potentiels/à lever ci-dessus.
    $alertRng = $Sheet.Range('B30:M30'); $alertRng.Merge() | Out-Null
    $alertRng.Formula = '=IF(Res_ProfilKey="Proactif","⚠ Le client a peut-être les données pour l''utilisation de l''outil Impact Durabilité By la Source","")'
    Set-CellStyle -Range $alertRng -Fill $Script:Palette.Warning -FontColor $Script:Palette.DarkText -Bold $true -HAlign 'center' -Wrap $true

    # --- Infos client ---
    $Sheet.Cells.Item(31,2).Value2 = 'Nom du client :'
    Set-CellStyle -Range $Sheet.Cells.Item(31,2) -Bold $true -FontColor $Script:Palette.DarkText -HAlign 'left'
    $cn = $Sheet.Range('D31:G31'); $cn.Merge() | Out-Null
    $cn.Interior.Color = ConvertTo-OleColor $Script:Palette.White
    $cn.Borders.Color = ConvertTo-OleColor $Script:Palette.Border
    $cn.Locked = $false
    New-DefinedName -Workbook $Workbook -Name 'Res_ClientName' -RangeObj $cn

    $Sheet.Cells.Item(33,2).Value2 = 'Notes de l''expert-comptable :'
    Set-CellStyle -Range $Sheet.Cells.Item(33,2) -Bold $true -FontColor $Script:Palette.DarkText -HAlign 'left'
    $notes = $Sheet.Range('B34:M37'); $notes.Merge() | Out-Null
    $notes.Interior.Color = ConvertTo-OleColor $Script:Palette.White
    $notes.Borders.Color = ConvertTo-OleColor $Script:Palette.Border
    $notes.VerticalAlignment = -4160
    $notes.WrapText = $true
    $notes.Locked = $false
    New-DefinedName -Workbook $Workbook -Name 'Res_Notes' -RangeObj $notes

    $exportNoteRng = $Sheet.Range('B39:M40'); $exportNoteRng.Merge() | Out-Null
    $exportNoteRng.Value2 = if ($Unified) {
        'Pour un PDF : Fichier -> Exporter -> Créer un PDF/XPS. Le client, le profil et le score sont repris automatiquement dans l''onglet Rapport_Trame.'
    } else {
        'Pour un PDF : Fichier -> Exporter -> Créer un PDF/XPS. Pour reporter ces résultats dans la Trame de rapport, recopiez le profil et le score dans les champs prévus de ce classeur.'
    }
    Set-CellStyle -Range $exportNoteRng -FontColor '6c757d' -Size 9 -HAlign 'left' -VAlign 'top' -Wrap $true

    # --- Formules (calculées en direct depuis Profilage_Questionnaire) ---
    $Sheet.Range('G6').Formula = '=SUMPRODUCT((INDEX(Ref_Options,0,1)=1)*(INDEX(Ref_Options,0,3)=Q1_Answer)*INDEX(Ref_Options,0,4))+SUMPRODUCT((INDEX(Ref_Options,0,1)=2)*(INDEX(Ref_Options,0,3)=Q2_Answer)*INDEX(Ref_Options,0,4))+SUMPRODUCT((INDEX(Ref_Options,0,1)=3)*(INDEX(Ref_Options,0,3)=Q3_Answer)*INDEX(Ref_Options,0,4))'
    $Sheet.Range('G7').Formula = '=-COUNTIF(Q4_Opts,TRUE)+COUNTIF(Q5_Opts,TRUE)+SUMPRODUCT((INDEX(Ref_Options,0,1)=6)*(INDEX(Ref_Options,0,3)=Q6_Answer)*INDEX(Ref_Options,0,4))'
    $Sheet.Range('G8').Formula = '=SUMPRODUCT((INDEX(Ref_Options,0,1)=8)*(INDEX(Ref_Options,0,3)=Q8_Answer)*INDEX(Ref_Options,0,4))+SUMPRODUCT((INDEX(Ref_Options,0,1)=9)*(INDEX(Ref_Options,0,3)=Q9_Answer)*INDEX(Ref_Options,0,4))'
    $Sheet.Range('G9').Formula = '=Res_Part1+Res_Part2+Res_Part4'
    $Sheet.Range('G10').Formula = '=SUM(Q7_Scores)'
    $Sheet.Range('G11').Formula = '=IF(Res_Part3>=16,"Très bonne",IF(Res_Part3>=11,"Moyenne","Faible"))'
    $Sheet.Range('G12').Formula = '=ROUND(Res_CoefRaw,2)'

    $Sheet.Range('Q16').Formula = '=IF(Res_Part3>=16,1,IF(Res_Part3>=11,0.75,0.6))'
    $Sheet.Range('Q17').Formula = '=IF(Res_Brut*Res_CoefBase>30,30/Res_Brut,IF(Res_Brut*Res_CoefBase<1,1/MAX(1,Res_Brut),Res_CoefBase))'
    $Sheet.Range('Q14').Formula = '=MIN(30,MAX(1,ROUND(Res_Brut*Res_CoefRaw,0)))'
    $Sheet.Range('Q15').Formula = '=LOOKUP(Res_ScoreFinal,INDEX(Ref_Profiles,0,2),INDEX(Ref_Profiles,0,1))'

    # Code de format "0,00" (virgule) et non "0.00" : sous Excel en locale française, le code
    # de format passé à TEXT() suit la convention décimale locale (virgule), pas la syntaxe
    # US de .Formula - "0.00" produit un résultat aberrant (le point est lu comme un séparateur
    # de milliers). Confirmé empiriquement sur cette installation.
    $Sheet.Range('B14').Formula = '=IF(Q1_Answer="","Complétez le questionnaire pour voir votre score.",Res_Brut&" x "&TEXT(Res_CoefRaw,"0,00")&" = "&Res_ScoreFinal&"/30")'
    $Sheet.Range('B17').Formula = '=IF(Q1_Answer="","En attente de réponses...",LOOKUP(Res_ScoreFinal,INDEX(Ref_Profiles,0,2),INDEX(Ref_Profiles,0,5))&" - "&LOOKUP(Res_ScoreFinal,INDEX(Ref_Profiles,0,2),INDEX(Ref_Profiles,0,4)))'
    $Sheet.Range('B18').Formula = '=IF(Q1_Answer="","",LOOKUP(Res_ScoreFinal,INDEX(Ref_Profiles,0,2),INDEX(Ref_Profiles,0,6)))'
    $Sheet.Range('B22').Formula = '=IF(Q1_Answer="","","- "&LOOKUP(Res_ScoreFinal,INDEX(Ref_Profiles,0,2),INDEX(Ref_Profiles,0,7))&CHAR(10)&"- "&LOOKUP(Res_ScoreFinal,INDEX(Ref_Profiles,0,2),INDEX(Ref_Profiles,0,8))&CHAR(10)&"- "&LOOKUP(Res_ScoreFinal,INDEX(Ref_Profiles,0,2),INDEX(Ref_Profiles,0,9)))'
    $Sheet.Range('H22').Formula = '=IF(Q1_Answer="","",CHAR(34)&LOOKUP(Res_ScoreFinal,INDEX(Ref_Profiles,0,2),INDEX(Ref_Profiles,0,10))&CHAR(34)&CHAR(10)&CHAR(34)&LOOKUP(Res_ScoreFinal,INDEX(Ref_Profiles,0,2),INDEX(Ref_Profiles,0,11))&CHAR(34)&CHAR(10)&CHAR(34)&LOOKUP(Res_ScoreFinal,INDEX(Ref_Profiles,0,2),INDEX(Ref_Profiles,0,12))&CHAR(34))'
    $Sheet.Range('B28').Formula = '=IF(Q1_Answer="","",LOOKUP(Res_ScoreFinal,INDEX(Ref_Profiles,0,2),INDEX(Ref_Profiles,0,13)))'
}
