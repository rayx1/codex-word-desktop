# Setup checks only. Never creates Word, reads document content or changes policy.
param([switch]$Json, [string]$Python = 'python',
      [string]$WorkDir = "$env:LOCALAPPDATA/CodexWord/work",
      [switch]$SkipWord, [switch]$WordProbe)
$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = New-Object Text.UTF8Encoding($false)
if ($WordProbe) {
    try {
        $word = [Runtime.InteropServices.Marshal]::GetActiveObject('Word.Application')
        @{connected=$true; version=$word.Version; build=$word.Build; documents=$word.Documents.Count} | ConvertTo-Json -Compress
    } catch {
        $failure = $_.Exception
        while ($failure.InnerException) { $failure = $failure.InnerException }
        @{connected=$false; error=$failure.Message; hresult=('0x{0:X8}' -f $failure.HResult)} | ConvertTo-Json -Compress
    }
    exit 0
}
function RunProbe([string]$exe, [string[]]$arguments, [int]$timeout = 15000) {
    # Arguments are local paths or fixed flags, never shell expressions.
    $quoted = foreach ($a in $arguments) {
        if ($a.Contains('"')) { throw 'A probe path contains a double quote.' }
        '"' + $a.TrimEnd('\') + '"'
    }
    $info = New-Object Diagnostics.ProcessStartInfo
    $info.FileName = $exe
    $info.Arguments = $quoted -join ' '
    $info.UseShellExecute = $false
    $info.CreateNoWindow = $true
    $info.RedirectStandardOutput = $true
    $info.RedirectStandardError = $true
    $process = New-Object Diagnostics.Process
    $process.StartInfo = $info
    try {
        [void]$process.Start()
        $out = $process.StandardOutput.ReadToEndAsync()
        $err = $process.StandardError.ReadToEndAsync()
        if (!$process.WaitForExit($timeout)) {
            $process.Kill() # Only this probe child; never a Word process.
            $process.WaitForExit()
            throw 'Probe timed out. Inspect Word for a dialog; no Word process was stopped.'
        }
        return @{code=$process.ExitCode; stdout=$out.GetAwaiter().GetResult(); stderr=$err.GetAwaiter().GetResult()}
    } finally { $process.Dispose() }
}
$checks = @()
function Check([string]$name, [string]$status, [string]$message, [string]$action = '', $details = $null) {
    $script:checks += @{name=$name; status=$status; message=$message; action=$action; details=$details}
}
$identity = [Security.Principal.WindowsIdentity]::GetCurrent().Name
$session = (Get-Process -Id $PID).SessionId
$desktop = [Environment]::UserInteractive -and $session -gt 0 -and $identity -notmatch '(?i)\\(CodexSandbox[^\\]*|SYSTEM|LOCAL SERVICE|NETWORK SERVICE)$'
Check 'interactive_user' $(if($desktop){'pass'}else{'fail'}) "Account $identity; Windows session $session." 'Run this check as your normal signed-in desktop user. A Codex sandbox/service account cannot share Word COM with your desktop account.'
$hostOK = $PSVersionTable.PSEdition -eq 'Desktop' -and $PSVersionTable.PSVersion -ge [version]'5.1' -and [Threading.Thread]::CurrentThread.ApartmentState -eq 'STA'
Check 'powershell' $(if($hostOK){'pass'}else{'fail'}) "PowerShell $($PSVersionTable.PSVersion), $($PSVersionTable.PSEdition), $([Threading.Thread]::CurrentThread.ApartmentState)." 'Use powershell.exe -NoProfile -STA -ExecutionPolicy RemoteSigned -File .\plugin\diagnose.ps1. Do not change Group Policy or Office security.'
try {
    $type = [type]::GetTypeFromProgID('Word.Application')
    if (!$type) { throw 'Word.Application is not registered.' }
    Check 'word_installation' 'pass' 'Word COM registration exists. This does not establish activation or a live connection.'
} catch { Check 'word_installation' 'fail' $_.Exception.Message 'Install or repair licensed desktop Microsoft Word through its normal installer.' }
try {
    $WorkDir = [IO.Path]::GetFullPath($WorkDir)
    [void][IO.Directory]::CreateDirectory($WorkDir)
    $probeFile = Join-Path $WorkDir ('setup-probe-' + [guid]::NewGuid().ToString('N'))
    $stream = [IO.File]::Open($probeFile, [IO.FileMode]::CreateNew, [IO.FileAccess]::Write, [IO.FileShare]::None)
    $stream.Dispose()
    [IO.File]::Delete($probeFile)
    Check 'work_directory' 'pass' 'Dedicated work directory is writable; the temporary write probe was removed.' '' $WorkDir
} catch { Check 'work_directory' 'fail' $_.Exception.Message 'Choose a dedicated writable -WorkDir. No document permissions will be changed.' }
try {
    $pythonPath = (Get-Command $Python -CommandType Application -ErrorAction Stop | Select-Object -First 1).Source
    if ($pythonPath -match '(?i)\\Microsoft\\WindowsApps\\') { throw 'Python resolves to a Windows Store alias rather than a verified interpreter.' }
    $probe = RunProbe $pythonPath @("$PSScriptRoot/renderer.py", '--diagnose') 30000
    $data = $probe.stdout | ConvertFrom-Json
    if ($data.checks.Count -ne 3 -or $null -eq $data.ok) { throw 'Renderer returned an incomplete diagnostic report.' }
    foreach ($item in $data.checks) {
        $message = $(if($item.ok){"$($item.version) at $($item.path)"}else{$item.error})
        Check $item.name $(if($item.ok){'pass'}else{'fail'}) $message $item.action
    }
    if ($probe.code -ne 0 -and $data.ok) { throw 'Renderer diagnostic exited unsuccessfully.' }
} catch { Check 'python_and_renderer' 'fail' $_.Exception.Message 'Install Python 3.10+ or supply -Python with its full executable path; then run plugin/renderer.py --diagnose.' }
try {
    $codex = (Get-Command codex -CommandType Application -ErrorAction Stop | Select-Object -First 1).Source
    $version = RunProbe $codex @('--version')
    if ($version.code -ne 0) { throw $version.stderr }
    Check 'codex' 'pass' $version.stdout.Trim() '' $codex
} catch { Check 'codex' 'fail' $_.Exception.Message 'Install/update Codex and ensure its CLI is on PATH before registration.' }
if ($SkipWord) {
    Check 'word_connection' 'warn' 'Live Word attachment explicitly skipped.' 'Open Word normally and rerun without -SkipWord before claiming a live connection.'
} elseif (!$hostOK -or !$desktop) {
    Check 'word_connection' 'warn' 'Attachment skipped because this is not the required interactive Windows PowerShell host.' 'Run as the same desktop account as Word; do not start Word in the background or change Office security.'
} else {
    try {
        $probe = RunProbe "$env:SystemRoot/System32/WindowsPowerShell/v1.0/powershell.exe" @('-NoLogo','-NoProfile','-STA','-ExecutionPolicy','RemoteSigned','-File',$PSCommandPath,'-WordProbe')
        if ($probe.code -ne 0) { throw $probe.stderr }
        $data = $probe.stdout | ConvertFrom-Json
        if ($data.connected) { Check 'word_connection' 'pass' "Attached to Word $($data.version), build $($data.build); $($data.documents) open documents. No content read." '' $data }
        else { Check 'word_connection' 'warn' "No usable Word attachment: $($data.hresult) $($data.error)" 'Open Word normally in this account and resolve any Office dialog, then retry. Error 0x80070520 commonly requires the interactive user context; do not configure a service/DCOM workaround.' $data }
    } catch { Check 'word_connection' 'warn' $_.Exception.Message 'Resolve dialogs in Word and rerun. Only the timed-out probe helper is stopped.' }
}
$failures = @($checks | Where-Object status -eq 'fail')
$connection = $checks | Where-Object name -eq 'word_connection'
$result = @{ok=($failures.Count -eq 0); ready_for_word=($failures.Count -eq 0 -and $connection.status -eq 'pass'); checks=$checks; note='Local checks only. Codex MCP initialization, document operations and PDF export are separate verification steps.'}
if ($Json) { $result | ConvertTo-Json -Depth 10 -Compress }
else {
    foreach ($check in $checks) {
        Write-Host ("[{0}] {1}: {2}" -f $check.status.ToUpperInvariant(), $check.name, $check.message)
        if ($check.action -and $check.status -ne 'pass') { Write-Host ("  Next: " + $check.action) }
    }
    Write-Host "Prerequisites OK: $($result.ok). Live Word ready: $($result.ready_for_word)."
}
if (!$result.ok) { exit 1 }
exit 0
