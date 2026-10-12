# Zee AI for Word diagnostics - read-only. Writes diag-log.txt next to this script.
$ErrorActionPreference = 'Continue'
$log = Join-Path $PSScriptRoot 'diag-log.txt'
"===== Zee AI for Word diagnostics $(Get-Date) =====" | Out-File $log -Encoding utf8
function W($t) { $t | Out-File $log -Append -Encoding utf8 }

W '--- node / npm'
W (& node -v 2>&1)
W (& npm -v 2>&1)

W '--- processes listening on ports 3000, 3100 (Excel add-in) and 3200 (Word add-in)'
Get-NetTCPConnection -LocalPort 3000,3100,3200 -State Listen -ErrorAction SilentlyContinue | ForEach-Object {
  $p = Get-CimInstance Win32_Process -Filter "ProcessId=$($_.OwningProcess)"
  W ("{0}:{1} pid={2} name={3}" -f $_.LocalAddress, $_.LocalPort, $_.OwningProcess, $p.Name)
  W ("   cmd: " + $p.CommandLine)
}

foreach ($h in @('localhost:3200','127.0.0.1:3200','[::1]:3200')) {
  W "--- GET https://$h/taskpane.html (insecure / strict TLS)"
  W (& curl.exe -k -s -m 5 -o NUL -w "%{http_code}" "https://$h/taskpane.html")
  W (& curl.exe -s -m 5 -o NUL -w "%{http_code} %{errormsg}" "https://$h/taskpane.html")
}
W '--- hosted version (manifest.prod.xml)'
$prod = Join-Path $PSScriptRoot 'manifest.prod.xml'
if (Test-Path $prod) {
  $m = Select-String -Path $prod -Pattern 'id="Taskpane\.Url" DefaultValue="([^"]+)"' | Select-Object -First 1
  if ($m) { $u = $m.Matches[0].Groups[1].Value; W ($u + ' -> ' + (& curl.exe -s -m 10 -o NUL -w "%{http_code}" $u)) }
} else { W '(not published yet)' }
W '--- GET office.js status'
W (& curl.exe -s -m 10 -o NUL -w "%{http_code}" "https://appsforoffice.microsoft.com/lib/1/hosted/office.js")

W '--- dev certificate'
& npx --no office-addin-dev-certs verify 2>&1 | Out-String | W

W '--- built files'
W ("dist\taskpane.html exists: " + (Test-Path (Join-Path $PSScriptRoot 'dist\taskpane.html')))

W '--- registered developer add-ins'
Get-ItemProperty 'HKCU:\Software\Microsoft\Office\16.0\WEF\Developer' -ErrorAction SilentlyContinue | Out-String | W
W '--- loopback exemption for WebView'
& CheckNetIsolation.exe LoopbackExempt -s 2>&1 | Select-String -Pattern 'win32webviewhost|cw5n1h2txyewy' -Context 0,1 | Out-String | W
W '--- last 30 lines of client-log.txt'
$cl = Join-Path $PSScriptRoot 'client-log.txt'
if (Test-Path $cl) { Get-Content $cl -Tail 30 | Out-String | W } else { W '(none)' }
W '===== done ====='
