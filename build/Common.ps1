# Helpers communs pour la génération du classeur ExpertComptable
# Dot-sourcé par Build-Workbook.ps1

$Script:Palette = @{
    Green      = '055c2c'
    GreenDark  = '033d1d'
    DarkText   = '02131b'
    White      = 'FFFFFF'
    LightGray  = 'F8F9FA'
    CardGray   = 'F0F2F1'
    Success    = '28a745'
    Danger     = 'dc3545'
    Warning    = 'ffc107'
    Info       = '17a2b8'
    Border     = 'DEE2E6'
}

function Set-CellValue2 {
    <#
        Écrit dans Range.Value2 via InvokeMember plutôt que l'affectation directe PowerShell.
        Contourne un bug de marshaling COM (Windows PowerShell 5.1 / .NET Framework) où des
        affectations consécutives de types différents (Int32 puis String) sur Range.Value2
        lèvent parfois une InvalidCastException factice ("Le cast spécifié n'est pas valide").
        Utilisé partout où une boucle écrit des valeurs de types hétérogènes ligne par ligne
        (cf. Write-RefTable dans New-Referentiel.ps1).
    #>
    param([Parameter(Mandatory)]$Cell, $Val)
    $Cell.GetType().InvokeMember('Value2', [System.Reflection.BindingFlags]::SetProperty, $null, $Cell, @($Val)) | Out-Null
}

function ConvertTo-OleColor {
    param([Parameter(Mandatory)][string]$Hex)
    $h = $Hex.TrimStart('#')
    $r = [Convert]::ToInt32($h.Substring(0,2),16)
    $g = [Convert]::ToInt32($h.Substring(2,2),16)
    $b = [Convert]::ToInt32($h.Substring(4,2),16)
    return [int]($r + $g*256 + $b*65536)
}

function Set-CellStyle {
    param(
        [Parameter(Mandatory)]$Range,
        [string]$Fill,
        [string]$FontColor,
        [bool]$Bold = $false,
        [double]$Size = 10,
        [string]$HAlign,
        [string]$VAlign = 'center',
        [bool]$Wrap = $false
    )
    if ($Fill) {
        $Range.Interior.Color = ConvertTo-OleColor $Fill
    }
    if ($FontColor) { $Range.Font.Color = ConvertTo-OleColor $FontColor }
    $Range.Font.Bold = $Bold
    $Range.Font.Size = $Size
    $Range.Font.Name = 'Segoe UI'
    if ($HAlign) {
        $Range.HorizontalAlignment = switch ($HAlign) {
            'left'   { -4131 }
            'center' { -4108 }
            'right'  { -4152 }
            default  { -4131 }
        }
    }
    $Range.VerticalAlignment = if ($VAlign -eq 'center') { -4108 } else { -4160 }
    $Range.WrapText = $Wrap
}

function New-DefinedName {
    param([Parameter(Mandatory)]$Workbook, [Parameter(Mandatory)][string]$Name, [Parameter(Mandatory)]$RangeObj)
    $existing = $null
    try { $existing = $Workbook.Names.Item($Name) } catch {}
    if ($existing) { $existing.Delete() }
    $Workbook.Names.Add($Name, $RangeObj) | Out-Null
}

function Add-NavBar {
    <#
        Dessine la barre de navigation verte persistante en haut de la feuille donnée.
        - Title : nom du module affiché à côté du logo (ex: "Outil de profilage").
        - Buttons : tableau de @{ Key; Label; TargetSheet; TargetCell } propres au fichier -
          liens hypertexte natifs (pas de macro), donc cliquables sur PC, mobile et Excel Online.
    #>
    param([Parameter(Mandatory)]$Sheet, [Parameter(Mandatory)]$Workbook, [string]$Title = 'Expert Comptable',
          [array]$Buttons = @(), [string]$ActiveKey = '')

    # Bande verte dimensionnee sur 14 colonnes (comportement d'origine, inchange) tant qu'il y a
    # au plus 3 boutons. Au-dela (barre de nav unifiee du classeur Suite, jusqu'a 6 boutons), on
    # elargit la plage jusqu'a couvrir la position reelle du dernier bouton (mesuree via .Left,
    # en points) - sinon le fond vert s'arrete avant la fin des boutons et la barre parait coupee.
    $endCol = 14
    if ($Buttons.Count -gt 3) {
        $requiredWidth = 280 + $Buttons.Count * 146 + 30
        while ($endCol -lt 100 -and $Sheet.Cells.Item(1, $endCol).Left -lt $requiredWidth) { $endCol++ }
    }
    $navRange = $Sheet.Range($Sheet.Cells.Item(1,1), $Sheet.Cells.Item(2,$endCol))
    $navRange.Interior.Color = ConvertTo-OleColor $Script:Palette.Green

    $logoPath = Join-Path $PSScriptRoot '..\..\ExpertComptable\public\images\Expert comptable.jpg'
    if (Test-Path $logoPath) {
        $pic = $Sheet.Shapes.AddPicture($logoPath, $false, $true, 6, 4, 30, 30)
        $pic.Placement = 2
    }

    $titleShape = $Sheet.Shapes.AddTextbox(1, 42, 2, 220, 34)
    $titleShape.TextFrame2.TextRange.Text = $Title
    $titleShape.TextFrame2.TextRange.Font.Bold = $true
    $titleShape.TextFrame2.TextRange.Font.Size = 14
    $titleShape.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = ConvertTo-OleColor $Script:Palette.White
    $titleShape.Fill.Visible = 0
    $titleShape.Line.Visible = 0
    $titleShape.TextFrame2.VerticalAnchor = 3
    $titleShape.Placement = 2

    $x = 280
    foreach ($b in $Buttons) {
        $isActive = ($b.Key -eq $ActiveKey)
        $fillColor = if ($isActive) { $Script:Palette.GreenDark } else { $Script:Palette.Green }
        $targetCell = if ($b.TargetCell) { $b.TargetCell } else { 'A1' }
        Add-NavHyperlink -Sheet $Sheet -Workbook $Workbook -Left $x -Top 4 -Width 140 -Height 30 -Label $b.Label `
            -TargetSheet $b.TargetSheet -TargetCell $targetCell -Fill $fillColor -TextColor $Script:Palette.White -Border $true | Out-Null
        $x += 146
    }
}

function Add-NavHyperlink {
    <#
        Bouton de navigation SANS macro : un lien hypertexte natif Excel ancré sur une forme,
        qui saute vers une cellule d'une autre feuille du MEME classeur. Fonctionne sur
        PC, Mac, Excel Online et Excel mobile - aucune macro requise.
    #>
    param($Sheet, $Workbook, [double]$Left, [double]$Top, [double]$Width, [double]$Height, [string]$Label,
          [string]$TargetSheet, [string]$TargetCell = 'A1',
          [string]$Fill = '055c2c', [string]$TextColor = 'FFFFFF', [bool]$Border = $false, [double]$FontSize = 10)
    $shp = $Sheet.Shapes.AddShape(5, $Left, $Top, $Width, $Height) # 5 = msoShapeRoundedRectangle
    $shp.Fill.ForeColor.RGB = ConvertTo-OleColor $Fill
    if ($Border) { $shp.Line.ForeColor.RGB = ConvertTo-OleColor 'FFFFFF'; $shp.Line.Weight = 0.75 } else { $shp.Line.Visible = 0 }
    $tr = $shp.TextFrame2.TextRange
    $tr.Text = $Label
    $tr.Font.Bold = $true
    $tr.Font.Size = $FontSize
    $tr.Font.Fill.ForeColor.RGB = ConvertTo-OleColor $TextColor
    $tr.ParagraphFormat.Alignment = 2 # msoAlignCenter
    $shp.TextFrame2.VerticalAnchor = 3 # msoAnchorMiddle
    $shp.TextFrame2.WordWrap = 0
    $Sheet.Hyperlinks.Add($shp, '', "'$TargetSheet'!$TargetCell", [Type]::Missing, $Label) | Out-Null
    Write-Output -NoEnumerate $shp
}

function Add-NavButton {
    param($Sheet, [double]$Left, [double]$Top, [double]$Width, [double]$Height, [string]$Label, [string]$Macro,
          [string]$Fill = '055c2c', [string]$TextColor = 'FFFFFF', [bool]$Border = $false, [double]$FontSize = 10)
    $shp = $Sheet.Shapes.AddShape(5, $Left, $Top, $Width, $Height) # 5 = msoShapeRoundedRectangle
    $shp.Fill.ForeColor.RGB = ConvertTo-OleColor $Fill
    if ($Border) { $shp.Line.ForeColor.RGB = ConvertTo-OleColor 'FFFFFF'; $shp.Line.Weight = 0.75 } else { $shp.Line.Visible = 0 }
    $shp.OnAction = $Macro
    $tr = $shp.TextFrame2.TextRange
    $tr.Text = $Label
    $tr.Font.Bold = $true
    $tr.Font.Size = $FontSize
    $tr.Font.Fill.ForeColor.RGB = ConvertTo-OleColor $TextColor
    $tr.ParagraphFormat.Alignment = 2 # msoAlignCenter
    $shp.TextFrame2.VerticalAnchor = 3 # msoAnchorMiddle
    $shp.TextFrame2.WordWrap = 0
    Write-Output -NoEnumerate $shp
}

function Add-SectionHeader {
    param($Sheet, [string]$CellAddr, [string]$Text, [int]$MergeCols = 8)
    $c = $Sheet.Range($CellAddr)
    $endCol = $c.Column + $MergeCols - 1
    $rng = $Sheet.Range($c, $Sheet.Cells.Item($c.Row, $endCol))
    $rng.Merge() | Out-Null
    $rng.Value2 = $Text
    Set-CellStyle -Range $rng -Fill $Script:Palette.Green -FontColor $Script:Palette.White -Bold $true -Size 12 -HAlign 'left'
    $rng.RowHeight = 24
    Write-Output -NoEnumerate $rng
}

function Add-Card {
    param($Sheet, [string]$RangeAddr)
    $rng = $Sheet.Range($RangeAddr)
    $rng.Interior.Color = ConvertTo-OleColor $Script:Palette.LightGray
    $rng.Borders.LineStyle = 0 # xlLineStyleNone - retire toute grille interne héritée
    # BorderAround(LineStyle, Weight, ColorIndex, Color) : contour extérieur uniquement.
    # Ne PAS suivre cet appel d'un .Borders.Color= séparé : cela réactiverait les bordures internes.
    $rng.BorderAround(1, -4138, [Type]::Missing, (ConvertTo-OleColor $Script:Palette.Border)) | Out-Null
    Write-Output -NoEnumerate $rng
}

function Build-PrefixChainFormula {
    <#
        Génère une chaîne de SI imbriqués testant, dans l'ordre exact de priorité fourni,
        si CellRef commence par chaque préfixe - reproduit exactement la logique VBA
        "premier préfixe qui correspond gagne" (E5Controller::detectCycle /
        generateAnalysisForLine), sans macro : formule pure, calculée à la construction.
        - Rules : tableau ordonné d'objets avec .Prefix et .Value (texte ou nombre littéral
          à retourner tel quel, sans guillemets si $ValueIsLiteral).
        - Default : valeur retournée si aucun préfixe ne correspond.
    #>
    param([Parameter(Mandatory)][string]$CellRef, [Parameter(Mandatory)][array]$Rules,
          [Parameter(Mandatory)]$Default, [bool]$ValueIsLiteral = $false)
    $expr = if ($ValueIsLiteral) { "$Default" } else { '"' + $Default + '"' }
    for ($i = $Rules.Count; $i -ge 1; $i--) {
        $rule = $Rules[$i-1]
        $val = if ($ValueIsLiteral) { "$($rule.Value)" } else { '"' + $rule.Value + '"' }
        $prefix = $rule.Prefix -replace '"', '""'
        $expr = "IF(LEFT($CellRef,$($rule.Prefix.Length))=""$prefix"",$val,$expr)"
    }
    return $expr
}

function Import-PsdSafe {
    param([Parameter(Mandatory)][string]$Path)
    return Import-PowerShellDataFile -Path $Path
}

function Get-ExcelColLetter {
    param([int]$ColIndex)
    $dividend = $ColIndex
    $columnName = ''
    while ($dividend -gt 0) {
        $modulo = ($dividend - 1) % 26
        $columnName = [char](65 + $modulo) + $columnName
        $dividend = [int](($dividend - $modulo) / 26)
    }
    return $columnName
}
