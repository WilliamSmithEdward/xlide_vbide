# Standing PASS/FAIL probe for the shim's debug api: every route, and the break-mode round
# trip that is the whole point of it (push code, set a breakpoint, run, read live locals).
# Run tools\dev.ps1 -KeepOpen first, or any harness session; Debug builds only.
# Dev-harness script: uses Application.VBE per decision 10's harness exception.
$ErrorActionPreference = 'Continue'

$checks = [ordered] @{}
function Check([string] $name, [scriptblock] $test) {
    try {
        $result = & $test
        $checks[$name] = if ($result) { 'PASS' } else { 'FAIL' }
    } catch {
        $checks[$name] = "FAIL ($($_.Exception.Message))"
    }
}

$excel = Get-Process EXCEL -ErrorAction SilentlyContinue | Select-Object -First 1
if (-not $excel) { Write-Output 'RESULT: FAIL - no Excel is running'; return }

$discoveryPath = Join-Path $env:LOCALAPPDATA "xlide_vbide\debug-api-$($excel.Id).json"
if (-not (Test-Path $discoveryPath)) {
    Write-Output "RESULT: FAIL - no discovery file for Excel $($excel.Id) (Release build, or the editor was never opened)"
    return
}

$d = Get-Content $discoveryPath | ConvertFrom-Json
$api = "http://127.0.0.1:$($d.port)/$($d.token)"
Write-Output "api: $api (pid $($d.pid), devtools $($d.devtoolsPort))"

Check 'discovery carries a schema, port, token, and devtools port' {
    $d.api -ge 1 -and $d.port -gt 0 -and $d.token.Length -ge 8 -and $d.devtoolsPort -gt 0
}

Check 'the token gates: a wrong one is refused' {
    try {
        Invoke-RestMethod "http://127.0.0.1:$($d.port)/wrongtoken/state" -TimeoutSec 5 | Out-Null
        $false
    } catch {
        $true
    }
}

Check 'state answers with a live surface' {
    $s = Invoke-RestMethod "$api/state" -TimeoutSec 8
    $s.surfaceReady -and $s.configuration -eq 'debug'
}

Check 'windows lists the editor own windows' {
    (Invoke-RestMethod "$api/windows" -TimeoutSec 8).windows.Count -gt 0
}

Check 'stats reports counters' {
    $t = Invoke-RestMethod "$api/stats" -TimeoutSec 8
    $t.uptimeSeconds -ge 0 -and $t.handleCount -gt 0 -and $t.logLines -gt 0
}

Check 'log slices and filters server side' {
    $l = Invoke-RestMethod "$api/log?match=surface&max=5" -TimeoutSec 8
    $l.next -gt 0 -and $l.lines.Count -le 5
}

Check 'messages tap holds page traffic' {
    (Invoke-RestMethod "$api/messages?last=5" -TimeoutSec 8).messages.Count -ge 0
}

Check 'problems answers the analyzer findings' {
    $null -ne (Invoke-RestMethod "$api/problems" -TimeoutSec 8).findings
}

Check 'capture renders the frame to a bitmap' {
    $bmp = Join-Path $env:TEMP 'xlide-api-probe.bmp'
    Invoke-WebRequest "$api/capture" -OutFile $bmp -TimeoutSec 20
    (Get-Item $bmp).Length -gt 10000
}

Check 'module reads through the session reader' {
    (Invoke-RestMethod "$api/module?name=CleanModule" -TimeoutSec 8).text.Length -gt 0
}

# The round trip: write a runner, break inside it, read locals, reset. This is the shape
# every debugger-milestone regression will take.
$runner = @'
Option Explicit

Public Function Total(ByVal values As Variant) As Double
    Dim i As Long
    Dim sum As Double

    For i = LBound(values) To UBound(values)
        sum = sum + values(i)
    Next i

    Total = sum
End Function

Public Sub RunTotal()
    Dim answer As Double
    answer = Total(Array(1, 2, 3))
End Sub
'@

Check 'module writes through the session writer' {
    Invoke-RestMethod "$api/module?name=CleanModule" -Method Post -Body $runner -TimeoutSec 10 | Out-Null
    Start-Sleep -Milliseconds 1500
    (Invoke-RestMethod "$api/module?name=CleanModule" -TimeoutSec 8).text -match 'RunTotal'
}

Check 'breakpoint is idempotent with state=on' {
    $first = (Invoke-RestMethod "$api/breakpoint?module=CleanModule&line=8&state=on" -Method Post -TimeoutSec 8).ran
    $again = (Invoke-RestMethod "$api/breakpoint?module=CleanModule&line=8&state=on" -Method Post -TimeoutSec 8).ran
    $first -and $again
}

$reachedBreak = $false
Check 'a run reaches the breakpoint and state says break' {
    $app = [Runtime.InteropServices.Marshal]::GetActiveObject('Excel.Application')
    $book = Split-Path -Leaf $app.ActiveWorkbook.FullName
    $app.OnTime([DateTime]::Now.AddSeconds(1), "'$book'!RunTotal")
    $deadline = (Get-Date).AddSeconds(25)
    do {
        Start-Sleep -Milliseconds 400
        $s = Invoke-RestMethod "$api/state" -TimeoutSec 5
    } while ($s.debugMode -ne 'break' -and (Get-Date) -lt $deadline)
    $script:reachedBreak = $s.debugMode -eq 'break'
    $script:reachedBreak
}

Check 'locals answers the stopped scope' {
    if (-not $script:reachedBreak) { throw 'never reached break' }
    Start-Sleep -Milliseconds 1500
    $rows = (Invoke-RestMethod "$api/locals" -TimeoutSec 8).rows
    ($rows | Where-Object { $_.expression -eq 'sum' }) -and ($rows | Where-Object { $_.expression -eq 'i' })
}

Check 'watches answers stopped with no watches set' {
    if (-not $script:reachedBreak) { throw 'never reached break' }
    (Invoke-RestMethod "$api/watches" -TimeoutSec 8).stopped
}

Check 'command by name resets out of the break' {
    Invoke-RestMethod "$api/command?name=reset" -Method Post -TimeoutSec 8 | Out-Null
    Start-Sleep -Milliseconds 1200
    (Invoke-RestMethod "$api/state" -TimeoutSec 8).debugMode -ne 'break'
}

Check 'breakpoint clears with state=off' {
    (Invoke-RestMethod "$api/breakpoint?module=CleanModule&line=8&state=off" -Method Post -TimeoutSec 8).ran -eq $false
}

Check 'an unknown route is named, not crashed on' {
    (Invoke-RestMethod "$api/nonesuch" -TimeoutSec 8).error -match 'unknown route'
}

Write-Output ''
foreach ($name in $checks.Keys) { "  {0,-52} {1}" -f $name, $checks[$name] }
$failed = @($checks.Values | Where-Object { $_ -ne 'PASS' })
Write-Output ''
if ($failed.Count -eq 0) {
    Write-Output "RESULT: PASS - $($checks.Count) checks"
} else {
    Write-Output "RESULT: FAIL - $($failed.Count) of $($checks.Count) checks failed"
}
