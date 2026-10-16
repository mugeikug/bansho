# Add the Bansho app to Explorer's "Send to" menu.
# Run once after installing the app from Edge (Apps > Install this site as an app).
# It reuses the command Edge registered for opening .board files with the installed app.
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Windows.Forms
function Say($msg) { [System.Windows.Forms.MessageBox]::Show($msg, 'Bansho') | Out-Null }

# ProgIDs that can open .board (registered by Edge when the app was installed)
$progIds = @()
$key = 'HKCU:\Software\Classes\.board\OpenWithProgids'
if (Test-Path $key) { $progIds = (Get-Item $key).Property }
$cmd = $null; $icon = $null
foreach ($p in $progIds) {
  $c = (Get-ItemProperty -LiteralPath "HKCU:\Software\Classes\$p\shell\open\command" -ErrorAction SilentlyContinue).'(default)'
  if ($c -and $c -match '--app-id=') {
    $cmd = $c
    $icon = (Get-ItemProperty -LiteralPath "HKCU:\Software\Classes\$p\DefaultIcon" -ErrorAction SilentlyContinue).'(default)'
    break
  }
}
if (-not $cmd) {
  Say ("The installed app was not found.`n`nFirst open https://mugeikug.github.io/bansho/ in Microsoft Edge and install it as an app, then run this again.")
  exit 1
}

# Split "exe" args ... "%1" into exe and args (Send to appends the file path)
if ($cmd -match '^\s*"([^"]+)"\s*(.*)$') { $exe = $Matches[1]; $args0 = $Matches[2] }
else { $parts = $cmd -split '\s+', 2; $exe = $parts[0]; $args0 = $parts[1] }
$args0 = ($args0 -replace '"?%1"?', '' -replace '\s+--\s*$', ' --').Trim()

$sendTo = [Environment]::GetFolderPath('SendTo')
$name = [string][char]0x30C7 + [char]0x30B8 + [char]0x30BF + [char]0x30EB + [char]0x677F + [char]0x66F8   # app name in Japanese
$lnkPath = Join-Path $sendTo ($name + '.lnk')
$sh = New-Object -ComObject WScript.Shell
$lnk = $sh.CreateShortcut($lnkPath)
$lnk.TargetPath = $exe
$lnk.Arguments = $args0
if ($icon) { $lnk.IconLocation = $icon }
$lnk.Save()
Say ("Added to the Send to menu:`n" + $lnkPath)
