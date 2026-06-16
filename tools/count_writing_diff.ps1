param(
    [Parameter(Mandatory = $false, Position = 0)]
    [string]$Before,

    [Parameter(Mandatory = $false, Position = 1)]
    [string]$After,

    [Parameter(Mandatory = $false)]
    [string]$Csv,

    [switch]$Json
)

$ErrorActionPreference = "Stop"

$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$ProjectDir = Split-Path -Parent $ScriptDir
$BundledPython = Join-Path $env:USERPROFILE ".cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe"
$SystemPython = "python"
$Tool = Join-Path $ScriptDir "word_diff_counter.py"

if (Test-Path $BundledPython) {
    $Python = $BundledPython
} else {
    $Python = $SystemPython
}

$ArgsList = @($Tool)
if ($Before) {
    $ArgsList += $Before
}
if ($After) {
    $ArgsList += $After
}
if ($Csv) {
    $ArgsList += @("--csv", $Csv)
}
if ($Json) {
    $ArgsList += "--json"
}

& $Python @ArgsList
