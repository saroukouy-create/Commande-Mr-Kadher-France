$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent $PSScriptRoot
$outDir = Join-Path $root 'dist\preview'
if (-not (Test-Path $outDir)) { New-Item -ItemType Directory -Path $outDir -Force | Out-Null }

$files = @(
    @{ Path = Join-Path $root 'dist\Outil-de-Profilage.xlsm'; Sheets = @('Profilage_Questionnaire','Profilage_Resultats') }
    @{ Path = Join-Path $root 'dist\Integrateur-de-Balance.xlsm'; Sheets = @('Balance_Import','Balance_Analyse') }
    @{ Path = Join-Path $root 'dist\Trame-de-Rapport.xlsm'; Sheets = @('Rapport_Trame') }
)

$excel = $null
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false

    foreach ($f in $files) {
        $wb = $excel.Workbooks.Open($f.Path, $false, $true)
        foreach ($name in $f.Sheets) {
            $sh = $wb.Sheets.Item($name)
            $sh.Activate()
            $outPath = Join-Path $outDir "$name.pdf"
            $sh.PageSetup.Zoom = 65
            $sh.PageSetup.Orientation = 2
            $sh.ExportAsFixedFormat(0, $outPath)
            Write-Host "Exporte : $outPath"
        }
        $wb.Close($false)
        [Runtime.InteropServices.Marshal]::ReleaseComObject($wb) | Out-Null
    }
} finally {
    if ($excel) { $excel.Quit() }
    if ($excel) { [Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null }
}
