# Lance les 3 scripts de vérification headless (un par classeur)
$ErrorActionPreference = 'Continue'
$overallExit = 0
foreach ($t in @('Verify-Profilage.ps1','Verify-Balance.ps1','Verify-Rapport.ps1','Verify-Unified.ps1')) {
    Write-Host "`n########## $t ##########"
    & powershell -NoProfile -ExecutionPolicy Bypass -File (Join-Path $PSScriptRoot $t)
    if ($LASTEXITCODE -ne 0) { $overallExit = 1 }
}
exit $overallExit
