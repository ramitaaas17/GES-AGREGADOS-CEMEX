$ErrorActionPreference = 'Stop'
try {
    $carpeta = Split-Path -Parent $PSScriptRoot
    $escritorio = [Environment]::GetFolderPath('DesktopDirectory')
    $shell = New-Object -ComObject WScript.Shell
    $acceso = $shell.CreateShortcut((Join-Path $escritorio 'Generador de Matrices.lnk'))
    $acceso.TargetPath = Join-Path $carpeta 'Generador de Matrices.cmd'
    $acceso.WorkingDirectory = $carpeta
    $acceso.IconLocation = (Join-Path $PSScriptRoot 'generador-matrices.ico') + ',0'
    $acceso.Description = 'CEMEX - Generador de matrices y dashboard'
    $acceso.Save()
    Write-Host 'Acceso Generador de Matrices creado en tu escritorio.'
} catch {
    Write-Host $_.Exception.Message
    exit 1
}
