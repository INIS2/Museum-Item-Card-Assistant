param(
    [int]$Port = 8080
)

$root = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $root

Write-Host "Serving $root on http://127.0.0.1:$Port"
Write-Host "Press Ctrl+C to stop."

python -m http.server $Port
