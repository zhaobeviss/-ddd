$ErrorActionPreference = "Stop"

$Root = Split-Path -Parent $MyInvocation.MyCommand.Path

Push-Location $Root
try {
  & (Join-Path $Root "update_dashboard.ps1")

  git add index.html public/index.html
  git diff --cached --quiet
  if ($LASTEXITCODE -eq 0) {
    Write-Host "No dashboard changes to commit."
    exit 0
  }

  $message = "Update customs dashboard data {0}" -f (Get-Date -Format "yyyy-MM-dd HH:mm")
  git commit -m $message
  git push
}
finally {
  Pop-Location
}
