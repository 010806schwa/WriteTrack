param(
    [Parameter(ValueFromRemainingArguments = $true)]
    [string[]]$RemainingArgs
)

$ErrorActionPreference = "Continue"

$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$ProjectDir = Split-Path -Parent $ScriptDir
$OutputDir = Join-Path $ProjectDir "outputs"
$OutputFile = Join-Path $OutputDir "latest_word_diff.txt"
$CoreScript = Join-Path $ScriptDir "count_writing_diff.ps1"

New-Item -ItemType Directory -Force -Path $OutputDir | Out-Null
Set-Location $ProjectDir

$output = & $CoreScript @RemainingArgs 2>&1
$exitCode = $LASTEXITCODE

$output | Tee-Object -FilePath $OutputFile

Write-Host ""
Write-Host "Result saved to:"
Write-Host $OutputFile
Write-Host ""
Read-Host "Press Enter to close this window"

exit $exitCode
