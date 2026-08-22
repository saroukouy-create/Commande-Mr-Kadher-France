# Vérifie/active temporairement AccessVBOM, importe les modules .bas, puis restaure le registre.
# Dot-sourcé par Build-Workbook.ps1

function Test-AccessVBOM {
    $key = 'HKCU:\Software\Microsoft\Office\16.0\Excel\Security'
    try {
        $val = Get-ItemProperty -Path $key -Name AccessVBOM -ErrorAction Stop
        return ($val.AccessVBOM -eq 1)
    } catch {
        return $false
    }
}

function Enable-AccessVBOM {
    <#
        Active HKCU:...\Security\AccessVBOM=1 si nécessaire, et retourne un objet décrivant
        l'état initial (pour restauration ultérieure). N'agit QUE sur HKCU (pas besoin d'admin).
    #>
    $key = 'HKCU:\Software\Microsoft\Office\16.0\Excel\Security'
    $hadValue = $false
    $originalValue = $null
    try {
        $existing = Get-ItemProperty -Path $key -Name AccessVBOM -ErrorAction Stop
        $hadValue = $true
        $originalValue = $existing.AccessVBOM
    } catch {
        $hadValue = $false
    }

    if (-not (Test-Path $key)) {
        New-Item -Path $key -Force | Out-Null
    }
    New-ItemProperty -Path $key -Name AccessVBOM -PropertyType DWord -Value 1 -Force | Out-Null

    return @{ Key = $key; HadValue = $hadValue; OriginalValue = $originalValue }
}

function Restore-AccessVBOM {
    param($State)
    if (-not $State) { return }
    if ($State.HadValue) {
        Set-ItemProperty -Path $State.Key -Name AccessVBOM -Value $State.OriginalValue -Force
    } else {
        Remove-ItemProperty -Path $State.Key -Name AccessVBOM -ErrorAction SilentlyContinue
    }
}

function Test-BasHeader {
    param([string]$Path)
    $expectedName = [IO.Path]::GetFileNameWithoutExtension($Path)
    $firstLine = Get-Content -Path $Path -TotalCount 1
    $expected = "Attribute VB_Name = `"$expectedName`""
    return ($firstLine.Trim() -eq $expected)
}

function Import-VbaModules {
    <#
        -VbaPaths : liste ordonnée de dossiers à importer (ex: common + dossier du module).
        Chaque dossier peut contenir des .bas et/ou un ThisWorkbook.txt.
    #>
    param($Workbook, [string[]]$VbaPaths)

    $allBasFiles = @()
    foreach ($path in $VbaPaths) {
        $allBasFiles += Get-ChildItem -Path $path -Filter '*.bas' | Sort-Object Name
    }
    foreach ($f in $allBasFiles) {
        if (-not (Test-BasHeader -Path $f.FullName)) {
            throw "En-tête VB_Name manquant/invalide dans $($f.Name) - la première ligne doit être : Attribute VB_Name = `"$($f.BaseName)`""
        }
    }
    foreach ($f in $allBasFiles) {
        $Workbook.VBProject.VBComponents.Import($f.FullName) | Out-Null
        Write-Host "  Module VBA importé : $($f.Name)"
    }

    # ThisWorkbook est un module document déjà existant : on ne peut pas l'Import()-er comme
    # un module standard, on injecte son code via CodeModule.AddFromString.
    foreach ($path in $VbaPaths) {
        $thisWbCodePath = Join-Path $path 'ThisWorkbook.txt'
        if (Test-Path $thisWbCodePath) {
            $code = Get-Content -Path $thisWbCodePath -Raw
            $comp = $Workbook.VBProject.VBComponents.Item('ThisWorkbook')
            if ($comp.CodeModule.CountOfLines -gt 0) {
                $comp.CodeModule.DeleteLines(1, $comp.CodeModule.CountOfLines)
            }
            $comp.CodeModule.AddFromString($code)
            Write-Host "  Code injecté dans ThisWorkbook (Workbook_Open -> ReprotectAllSheets)"
        }
    }
}
