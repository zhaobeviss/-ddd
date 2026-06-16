$ErrorActionPreference = "Stop"

$Root = Split-Path -Parent $MyInvocation.MyCommand.Path
$Python = "C:\Users\zhaob\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe"
$Builder = Join-Path $Root "build_dashboard.py"
$LogDir = Join-Path $Root "logs"
$LogFile = Join-Path $LogDir ("update-{0}.log" -f (Get-Date -Format "yyyyMMdd-HHmmss"))

New-Item -ItemType Directory -Force -Path $LogDir | Out-Null

function Write-LogLine {
  param([string]$Message)
  $line = "[{0}] {1}" -f (Get-Date -Format "yyyy-MM-dd HH:mm:ss"), $Message
  $line | Tee-Object -FilePath $LogFile -Append
}

Write-LogLine "Starting customs export BI dashboard update."
Write-LogLine "Builder: $Builder"

if (-not (Test-Path -LiteralPath $Python)) {
  Write-LogLine "Python runtime not found: $Python"
  exit 1
}

if (-not (Test-Path -LiteralPath $Builder)) {
  Write-LogLine "Builder script not found: $Builder"
  exit 1
}

Push-Location $Root
try {
  & $Python $Builder *>&1 | Tee-Object -FilePath $LogFile -Append
  $exitCode = $LASTEXITCODE
}
finally {
  Pop-Location
}

if ($exitCode -ne 0) {
  Write-LogLine "Update failed with exit code $exitCode."
  exit $exitCode
}

Write-LogLine "Dashboard updated successfully: $(Join-Path $Root 'index.html')"
