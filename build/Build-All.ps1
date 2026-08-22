# Génère les 3 classeurs autonomes (un par mission) + le classeur Unified (les 3 outils
# regroupés, référentiel partagé) dans dist\
# Usage : powershell -File Build-All.ps1

$ErrorActionPreference = 'Stop'
foreach ($m in @('Profilage','Balance','Rapport','Unified')) {
    Write-Host "`n================ Génération : $m ================"
    & powershell -NoProfile -ExecutionPolicy Bypass -File (Join-Path $PSScriptRoot 'Build-Workbook.ps1') -Module $m
    if ($LASTEXITCODE -and $LASTEXITCODE -ne 0) { throw "Échec de la génération du module $m (code $LASTEXITCODE)" }
}
Write-Host "`nLes 4 classeurs ont été générés dans dist\."
