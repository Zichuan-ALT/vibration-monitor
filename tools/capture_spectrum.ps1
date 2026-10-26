# 抓取一帧频谱并保存为 CSV（供 plot_spectra.py 绘图）
# 用法：
#   .\tools\capture_spectrum.ps1 -Port COM8 -Label "工况1-无配重"
param(
    [string]$Port  = "COM8",
    [string]$Label = "未命名工况",
    [int]$TimeoutSec = 30
)

$ErrorActionPreference = "Stop"
$root = Split-Path $PSScriptRoot -Parent
$dir  = Join-Path $root "experiments\data"
New-Item -ItemType Directory -Force -Path $dir | Out-Null

$sp = New-Object System.IO.Ports.SerialPort
$sp.PortName = $Port
$sp.BaudRate = 115200
$sp.ReadTimeout = 500
$sp.Encoding = [System.Text.Encoding]::UTF8

try { $sp.Open() } catch { Write-Error "串口打不开: $($_.Exception.Message)"; exit 1 }
$sp.DtrEnable = $false
$sp.RtsEnable = $false
Start-Sleep -Seconds 2
$sp.DiscardInBuffer()

$sp.Write("s")
Start-Sleep -Milliseconds 200

$lines = New-Object System.Collections.ArrayList
$buf = ""
$done = $false
$sw = [System.Diagnostics.Stopwatch]::StartNew()

while ($sw.Elapsed.TotalSeconds -lt $TimeoutSec -and -not $done) {
    try {
        $chunk = $sp.ReadExisting()
        if ($chunk) {
            $buf += $chunk
            while ($buf.Contains("`n")) {
                $i = $buf.IndexOf("`n")
                $line = $buf.Substring(0, $i).Trim()
                $buf = $buf.Substring($i + 1)
                if ($line -match '^#BEGIN') {
                    [void]$lines.Add($line)
                } elseif ($line -eq '#END') {
                    [void]$lines.Add($line)
                    $done = $true
                    break
                } elseif ($line -match '^\d') {
                    [void]$lines.Add($line)
                }
            }
        }
    } catch { }
    Start-Sleep -Milliseconds 30
}
$sp.Close()

if (-not $done) { Write-Error "超时：未收到完整频谱（检查板子是否在运行 08_fan_measure）"; exit 1 }

$begin = $lines | Where-Object { $_ -match '^#BEGIN' } | Select-Object -First 1
$stamp = Get-Date -Format 'yyyyMMdd_HHmmss'
$safe  = ($Label -replace '[^\w\u4e00-\u9fff-]+', '_')
$file  = Join-Path $dir ("spectrum_{0}_{1}.csv" -f $safe, $stamp)

$header = @(
    "# label: $Label",
    "# captured: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')",
    $begin
)
$body = $lines | Where-Object { $_ -notmatch '^#' }

($header + $body) | Set-Content -Path $file -Encoding UTF8

Write-Host ("已保存: {0}" -f $file)
Write-Host ("  {0}" -f $begin)
Write-Host ("  频谱点数: {0}" -f $body.Count)
