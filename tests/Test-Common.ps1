# Helpers communs aux 3 scripts de vérification (Dot-sourcé)

$script:results = @()

function Get-NamedValue {
    # Lit via Cells.Item(1,1) : une plage nommée fusionnée renvoie un tableau 2D via .Value2 direct.
    param($Workbook, [string]$Name)
    return $Workbook.Names.Item($Name).RefersToRange.Cells.Item(1, 1).Value2
}

function Assert-Equal {
    param([string]$Label, $Expected, $Actual)
    $ok = ($Expected -eq $Actual)
    $script:results += [pscustomobject]@{ Label = $Label; Expected = $Expected; Actual = $Actual; OK = $ok }
    $status = if ($ok) { 'OK  ' } else { 'FAIL' }
    Write-Host "[$status] $Label -- attendu: $Expected / obtenu: $Actual"
}

function Test-OnActionIntegrity {
    param($Workbook, [string[]]$SheetNames, [string[]]$KnownMacros)
    $badShapes = @()
    foreach ($sheetName in $SheetNames) {
        $sh = $Workbook.Sheets.Item($sheetName)
        foreach ($shp in $sh.Shapes) {
            if ($shp.OnAction -and $shp.OnAction -ne '') {
                if ($KnownMacros -notcontains $shp.OnAction) { $badShapes += "$sheetName!$($shp.Name) -> $($shp.OnAction)" }
            }
        }
    }
    if ($badShapes.Count -gt 0) { Write-Host "Boutons avec macro inconnue : $($badShapes -join '; ')" }
    Assert-Equal 'Tous les OnAction pointent vers une macro connue' 0 $badShapes.Count
}

function Write-Summary {
    Write-Host "`n=== RÉSUMÉ ==="
    $failed = @($script:results | Where-Object { -not $_.OK })
    Write-Host "$($script:results.Count) assertions, $($failed.Count) échec(s)."
    if ($failed.Count -gt 0) {
        $failed | ForEach-Object { Write-Host "  ECHEC: $($_.Label) (attendu $($_.Expected), obtenu $($_.Actual))" }
        exit 1
    } else {
        Write-Host "Toutes les vérifications sont passées."
        exit 0
    }
}
