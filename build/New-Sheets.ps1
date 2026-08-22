# Création des feuilles du classeur, dans l'ordre des onglets, avec habillage de base
# Dot-sourcé par Build-Workbook.ps1 - suppose $wb et $excel déjà initialisés

$Script:SheetNamesByModule = @{
    Profilage = @('Profilage_Questionnaire', 'Profilage_Resultats', 'Profils_Historique', 'Référentiel')
    Balance   = @('Balance_Import', 'Balance_Data', 'Balance_Analyse', 'Référentiel')
    Rapport   = @('Rapport_Trame', 'Référentiel')
    # Classeur unique regroupant les 3 outils : memes feuilles que les 3 classeurs autonomes,
    # UN SEUL onglet Référentiel partagé (au lieu d'un par classeur) - cf. New-ReferentielSheet.
    Unified   = @('Profilage_Questionnaire', 'Profilage_Resultats', 'Profils_Historique',
                  'Balance_Import', 'Balance_Data', 'Balance_Analyse',
                  'Rapport_Trame', 'Référentiel')
}

function New-AllSheets {
    param($Workbook, [Parameter(Mandatory)][string]$Module)

    $sheetNames = $Script:SheetNamesByModule[$Module]
    if (-not $sheetNames) { throw "Module inconnu : $Module" }

    while ($Workbook.Sheets.Count -gt 1) {
        $Workbook.Sheets.Item($Workbook.Sheets.Count).Delete()
    }
    $Workbook.Sheets.Item(1).Name = $sheetNames[0]

    for ($i = 1; $i -lt $sheetNames.Count; $i++) {
        $prev = $Workbook.Sheets.Item($i)
        $sh = $Workbook.Sheets.Add([Type]::Missing, $prev)
        $sh.Name = $sheetNames[$i]
    }

    $sheets = @{}
    foreach ($name in $sheetNames) {
        $sh = $Workbook.Sheets.Item($name)
        $sh.Cells.Font.Name = 'Segoe UI'
        $sh.Cells.Font.Size = 10
        $sh.Activate()
        $Workbook.Application.ActiveWindow.DisplayGridlines = $false
        $sheets[$name] = $sh
    }
    $sheets[$sheetNames[0]].Activate()

    return $sheets
}
