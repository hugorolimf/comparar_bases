param()

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

$ProjectRoot = $PSScriptRoot
Set-Location $ProjectRoot

$VenvPython = Join-Path $ProjectRoot '.venv\Scripts\python.exe'
$RequirementsFile = Join-Path $ProjectRoot 'requirements.txt'

function Get-PythonLauncher {
    if (Test-Path $VenvPython) {
        return @{ FilePath = $VenvPython; Arguments = @() }
    }

    $pythonCommand = Get-Command python -ErrorAction SilentlyContinue
    if ($pythonCommand) {
        return @{ FilePath = $pythonCommand.Source; Arguments = @() }
    }

    $pyCommand = Get-Command py -ErrorAction SilentlyContinue
    if ($pyCommand) {
        return @{ FilePath = $pyCommand.Source; Arguments = @('-3') }
    }

    throw 'Python não foi encontrado. Instale o Python 3.14+ ou configure o launcher py no Windows.'
}

function Ensure-VirtualEnvironment {
    if (Test-Path $VenvPython) {
        return
    }

    $launcher = Get-PythonLauncher
    Write-Host 'Criando ambiente virtual .venv...' -ForegroundColor Cyan
    & $launcher.FilePath @($launcher.Arguments + @('-m', 'venv', '.venv'))
    if (-not (Test-Path $VenvPython)) {
        throw 'Não foi possível criar o ambiente virtual .venv.'
    }
}

function Test-ProjectDependencies {
    param(
        [Parameter(Mandatory = $true)]
        [string]$PythonExecutable
    )

    & $PythonExecutable -c 'import openpyxl, InquirerPy, rich'
    return $LASTEXITCODE -eq 0
}

function Install-ProjectDependencies {
    param(
        [Parameter(Mandatory = $true)]
        [string]$PythonExecutable
    )

    if (-not (Test-Path $RequirementsFile)) {
        throw "Arquivo requirements.txt não encontrado em $ProjectRoot."
    }

    Write-Host 'Dependências ausentes ou desatualizadas. Instalando/atualizando pacotes...' -ForegroundColor Yellow
    & $PythonExecutable -m pip install --upgrade pip
    if ($LASTEXITCODE -ne 0) {
        throw 'Falha ao atualizar o pip.'
    }

    & $PythonExecutable -m pip install -r $RequirementsFile
    if ($LASTEXITCODE -ne 0) {
        throw 'Falha ao instalar as dependências do projeto.'
    }
}

Ensure-VirtualEnvironment

$PythonExecutable = $VenvPython

if (-not (Test-ProjectDependencies -PythonExecutable $PythonExecutable)) {
    Install-ProjectDependencies -PythonExecutable $PythonExecutable
}

if (-not (Test-ProjectDependencies -PythonExecutable $PythonExecutable)) {
    throw 'As dependências continuam indisponíveis após a instalação.'
}

Write-Host 'Iniciando o comparador...' -ForegroundColor Cyan
& $PythonExecutable .\main.py
