<#
.SYNOPSIS
  Fake LIS for testing Lab-I middleware mode: listens for HL7 over MLLP,
  prints each message and answers with an ACK.

.EXAMPLE
  .\tools\mllp-test-listener.ps1                  # port 6661, answers AA (accepted)
  .\tools\mllp-test-listener.ps1 -Port 7000
  .\tools\mllp-test-listener.ps1 -Reply AE        # reject everything (tests retries / failed\)
  .\tools\mllp-test-listener.ps1 -NoReply         # never answer (tests the ACK timeout)

  Received messages are also appended to -LogFile (default: %TEMP%\mllp-test-listener.log).
  They contain patient data: the log is kept out of the repository on purpose.
  Stop with Ctrl+C.
#>
param(
    [int]$Port = 6661,
    [ValidateSet('AA', 'AE', 'AR')][string]$Reply = 'AA',
    [switch]$NoReply,
    [string]$LogFile = (Join-Path $env:TEMP 'mllp-test-listener.log')
)

$VT = [char]0x0B; $FS = [char]0x1C; $CR = [char]0x0D
$enc = [System.Text.Encoding]::GetEncoding(1252)

$listener = [System.Net.Sockets.TcpListener]::new([System.Net.IPAddress]::Any, $Port)
$listener.Start()
Write-Host "MLLP test listener on port $Port (reply: $(if ($NoReply) {'none'} else {$Reply})). Log: $LogFile. Ctrl+C to stop." -ForegroundColor Cyan

try {
    while ($true) {
        $client = $listener.AcceptTcpClient()
        Write-Host "`n[$(Get-Date -f 'HH:mm:ss')] Connected: $($client.Client.RemoteEndPoint)" -ForegroundColor Green
        $stream = $client.GetStream()
        $buffer = New-Object byte[] 65536
        $pending = ''
        try {
            while ($client.Connected) {
                $n = $stream.Read($buffer, 0, $buffer.Length)
                if ($n -le 0) { break }
                $pending += $enc.GetString($buffer, 0, $n)
                while (($end = $pending.IndexOf($FS)) -ge 0) {
                    $start = $pending.IndexOf($VT)
                    $msg = if ($start -ge 0 -and $start -lt $end) { $pending.Substring($start + 1, $end - $start - 1) } else { $pending.Substring(0, $end) }
                    $pending = $pending.Substring([Math]::Min($end + 2, $pending.Length))

                    $stamp = Get-Date -f 'yyyy-MM-dd HH:mm:ss'
                    Write-Host "[$stamp] Message received:" -ForegroundColor Yellow
                    $msg -split "`r" | Where-Object { $_ } | ForEach-Object { Write-Host "  $_" }
                    Add-Content -Path $LogFile -Value "===== $stamp =====`r`n$($msg -replace "`r", "`r`n")"

                    if (-not $NoReply) {
                        $ctrl = ''
                        $msh = ($msg -split "`r") | Where-Object { $_ -like 'MSH|*' } | Select-Object -First 1
                        if ($msh) { $f = $msh -split '\|'; if ($f.Count -gt 9) { $ctrl = $f[9] } }
                        $ack = "MSH|^~\&|TESTLIS||LAB-I||$(Get-Date -f 'yyyyMMddHHmmss')||ACK^R01|ACK$ctrl|P|2.5.1$CR" +
                               "MSA|$Reply|$ctrl$(if ($Reply -ne 'AA') {'|Rejected by test listener'})$CR"
                        $bytes = $enc.GetBytes("$VT$ack$FS$CR")
                        $stream.Write($bytes, 0, $bytes.Length)
                        Write-Host "  -> replied MSA|$Reply|$ctrl" -ForegroundColor DarkGray
                    }
                }
            }
        } catch {
            Write-Host "Connection error: $($_.Exception.Message)" -ForegroundColor Red
        } finally {
            $client.Close()
            Write-Host "[$(Get-Date -f 'HH:mm:ss')] Disconnected" -ForegroundColor DarkGray
        }
    }
} finally {
    $listener.Stop()
}
