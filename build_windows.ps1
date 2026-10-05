param(
    [string]$PythonExe = "python",
    [string]$OutputDir = (Join-Path $PSScriptRoot "dist"),
    [string]$WorkDir = (Join-Path $PSScriptRoot "build")
)
$ErrorActionPreference = "Stop"
$env:PYINSTALLER_CONFIG_DIR = Join-Path $WorkDir "cache"
& $PythonExe -m PyInstaller --noconfirm --clean --onedir --windowed --name ImportadorGoalfy --icon (Join-Path $PSScriptRoot "assets\app.ico") --add-data "$(Join-Path $PSScriptRoot 'assets\app.ico');assets" --collect-all customtkinter --hidden-import pythoncom --hidden-import pywintypes --hidden-import win32com.client --exclude-module matplotlib --exclude-module scipy --exclude-module IPython --exclude-module pytest --distpath $OutputDir --workpath (Join-Path $WorkDir "pyinstaller") --specpath (Join-Path $WorkDir "spec") (Join-Path $PSScriptRoot "app.py")
if ($LASTEXITCODE -ne 0) { throw "Falha ao gerar o executável." }
$pastaApp = Join-Path $OutputDir "ImportadorGoalfy"
$configEquipe = Get-Content -LiteralPath (Join-Path $PSScriptRoot "config.json") -Raw -Encoding utf8 | ConvertFrom-Json
$configEquipe.pasta_destino = ""
foreach ($matriz in $configEquipe.matrizes) { $matriz.caminho = "" }
$textoConfig = $configEquipe | ConvertTo-Json -Depth 30
[System.IO.File]::WriteAllText((Join-Path $pastaApp "config.json"), $textoConfig, [System.Text.UTF8Encoding]::new($false))
Copy-Item -LiteralPath (Join-Path $PSScriptRoot "GUIA_EQUIPE.txt") -Destination (Join-Path $pastaApp "LEIA-ME.txt")
Compress-Archive -LiteralPath $pastaApp -DestinationPath (Join-Path $OutputDir "ImportadorGoalfy-Windows.zip") -Force
Write-Output "Pronto: $(Join-Path $OutputDir 'ImportadorGoalfy-Windows.zip')"
