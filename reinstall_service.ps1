#Requires -RunAsAdministrator
# PronMCP service reinstall: remove the delete-marked service, install fresh, start.
# Run from an ELEVATED (Administrator) PowerShell.
# NOTE: this stops the temporary server on port 8000 and replaces it with the
#       real service (a few seconds of downtime; SSE clients auto-reconnect).
#
# PREREQUISITE: download nssm.exe from https://nssm.cc/download and place it at
#   the path set in $nssm below (the binary is not bundled in this repo).
#
# BEFORE RUNNING: close services.msc and Task Manager's "Services" tab.
#   (If a handle is held open, the delete-marked state will not clear and
#    reinstall will fail.)
#
# Adjust the paths below to your install location before running.

$svc       = 'PronMCP'
$appdir    = 'C:\path\to\pron-mcp'                          # repo / install directory
$nssm      = "$appdir\nssm.exe"                             # downloaded from nssm.cc
$py        = "$appdir\.venv\Scripts\python.exe"
$logdir    = "$appdir\logs"
$cachePath = "$env:USERPROFILE\.pron-mcp\token_cache.bin"   # verified user token cache

function Show-State {
    Get-CimInstance Win32_Service -Filter "Name='$svc'" -ErrorAction SilentlyContinue |
        Select-Object Name,State,StartMode,ProcessId | Format-List
}

Write-Host '=== 0) initial state ===' -ForegroundColor Cyan
Show-State

# 1) remove any leftover service (finish the pending delete)
Write-Host '=== 1) remove existing service ===' -ForegroundColor Cyan
& $nssm stop $svc 2>$null
& sc.exe stop $svc 2>$null | Out-Null
& $nssm remove $svc confirm 2>$null
& sc.exe delete $svc 2>$null | Out-Null
Start-Sleep -Seconds 2

# 2) confirm it is fully gone
if (Get-Service -Name $svc -ErrorAction SilentlyContinue) {
    Write-Host '[ABORT] Service still present (delete pending). Close services.msc / Task Manager Services tab, then re-run.' -ForegroundColor Yellow
    return
}
Write-Host '  old service fully removed' -ForegroundColor Green

# 3) log directory
if (-not (Test-Path $logdir)) { New-Item -ItemType Directory -Path $logdir | Out-Null }

# 4) free port 8000 (kills the temporary server)
Write-Host '=== 4) free port 8000 ===' -ForegroundColor Cyan
$pids = (Get-NetTCPConnection -LocalPort 8000 -State Listen -ErrorAction SilentlyContinue).OwningProcess | Select-Object -Unique
foreach ($p in $pids) { if ($p -and $p -ne 0) { Write-Host "  kill PID $p"; Stop-Process -Id $p -Force -ErrorAction SilentlyContinue } }
Start-Sleep -Seconds 2

# 5) install service + configure
Write-Host '=== 5) install service ===' -ForegroundColor Cyan
& $nssm install $svc $py "-m pron_mcp"
& $nssm set $svc AppDirectory $appdir
& $nssm set $svc DisplayName "PronMCP"
& $nssm set $svc Start SERVICE_AUTO_START
& $nssm set $svc AppStdout "$logdir\stdout.log"
& $nssm set $svc AppStderr "$logdir\stderr.log"
& $nssm set $svc AppRotateFiles 1
& $nssm set $svc AppRotateBytes 10485760
# point the LocalSystem service at the verified user token cache (avoids auth failure/hang)
& $nssm set $svc AppEnvironmentExtra "TOKEN_CACHE_PATH=$cachePath"

# 6) start
Write-Host '=== 6) start service ===' -ForegroundColor Cyan
& $nssm start $svc
Start-Sleep -Seconds 5

# 7) verify
Write-Host '=== 7) verify ===' -ForegroundColor Cyan
Show-State
$nssmPid = (Get-CimInstance Win32_Service -Filter "Name='$svc'" -ErrorAction SilentlyContinue).ProcessId
if ($nssmPid) {
    Write-Host "nssm PID=$nssmPid / child python (must be present):"
    Get-CimInstance Win32_Process -Filter "ParentProcessId=$nssmPid" | Select-Object ProcessId,Name | Format-Table -Auto
}
Write-Host 'port 8000 listen:'
Get-NetTCPConnection -LocalPort 8000 -State Listen -ErrorAction SilentlyContinue | Select-Object LocalAddress,OwningProcess | Format-Table -Auto
Write-Host 'recent stderr log:'
if (Test-Path "$logdir\stderr.log") { Get-Content "$logdir\stderr.log" -Tail 15 }
Write-Host '=== done. SUCCESS if State=Running + child python + port 8000 listening ===' -ForegroundColor Green
