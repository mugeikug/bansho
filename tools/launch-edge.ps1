# Open index.html in Microsoft Edge "app mode" (a window without address bar or tabs).
# Called from the launcher .cmd in the parent folder. Extra arguments are passed to Edge.
# The page path is converted to a percent-encoded file:// URL because Edge's --app
# does not accept raw non-ASCII characters (e.g. Japanese folder names) in the URL.
$html = Join-Path (Split-Path $PSScriptRoot -Parent) 'index.html'
$url = ([Uri]$html).AbsoluteUri

$candidates = @(
  (Join-Path ${env:ProgramFiles(x86)} 'Microsoft\Edge\Application\msedge.exe'),
  (Join-Path $env:ProgramFiles 'Microsoft\Edge\Application\msedge.exe'),
  (Join-Path $env:LOCALAPPDATA 'Microsoft\Edge\Application\msedge.exe')
)
$edge = $candidates | Where-Object { $_ -and (Test-Path $_) } | Select-Object -First 1
if (-not $edge) {
  Add-Type -AssemblyName System.Windows.Forms
  [System.Windows.Forms.MessageBox]::Show('Microsoft Edge was not found.', 'Bansho') | Out-Null
  exit 1
}

$argList = @("--app=$url") + $args
Start-Process -FilePath $edge -ArgumentList $argList
