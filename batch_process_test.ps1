# Send N requests to the convert service and report results.
# Assumes fake_ocr + service are already running.
#
# Usage: powershell .\batch_process_test.ps1 <file.pptx> [total=100] [concurrency=5]

param(
    [Parameter(Mandatory=$true)]
    [string]$Pptx,
    
    [int]$Total = 100,
    [int]$Concurrency = 5
)

$Url = if ($env:API_URL) { $env:API_URL } else { "http://localhost:8000/ppt" }
$TmpDir = Join-Path $env:TEMP "pptx2pdf_$(New-Guid | Select-Object -ExpandProperty Guid)"

if (-not (Test-Path $Pptx)) {
    Write-Host "Usage: powershell .\batch_process_test.ps1 <file.pptx> [total] [concurrency]"
    exit 1
}

# 转换为绝对路径，确保子作业能访问
$PptxAbsolute = (Resolve-Path $Pptx).Path

New-Item -ItemType Directory -Path $TmpDir -Force | Out-Null

Write-Host "file=$Pptx  total=$Total  concurrency=$Concurrency  target=$Url"
Write-Host ""

$scriptBlock = {
    param($idx, $PptxPath, $Url, $TmpDir)
    $outFile = Join-Path $TmpDir $idx
    $output = & curl.exe -s -w "%{http_code} %{time_total}" -X POST "$Url" -F "file=@$PptxPath" 2>&1
    # 提取最后一行的状态码和耗时
    $trimmed = $output.Trim()
    if ($trimmed -match '(\d{3})\s+([\d.]+)$') {
        "$($matches[1]) $($matches[2])" | Out-File $outFile -Encoding ASCII
    } else {
        "999 0" | Out-File $outFile -Encoding ASCII
    }
}

# fire requests in batches of CONCURRENCY
$jobs = @()
$startTime = Get-Date
for ($i = 1; $i -le $Total; $i++) {
    $job = Start-Job -ScriptBlock $scriptBlock -ArgumentList $i, $PptxAbsolute, $Url, $TmpDir
    $jobs += $job
    [math]::Round(100 * $i / $Total)
    Write-Progress -Activity "Dispatching requests" -PercentComplete ([math]::Round(100 * $i / $Total)) -Status "$i / $Total"
    
    if ($jobs.Count -ge $Concurrency) {
        $jobs | Wait-Job | Out-Null
        $jobs = @()
    }
}
if ($jobs.Count -gt 0) {
    $jobs | Wait-Job | Out-Null
}
Write-Host ""
Write-Host ""

# tally results
$ok = 0
$fail = 0
$elapsedVals = @()

for ($i = 1; $i -le $Total; $i++) {
    $resFile = Join-Path $TmpDir $i
    $parts = (Get-Content $resFile -Raw).Split(' ', [System.StringSplitOptions]::RemoveEmptyEntries)
    $code = $parts[0]
    $elapsed = [double]$parts[1]
    
    if ($code -eq "200") {
        $ok++
    } else {
        $fail++
        Write-Host "  [FAIL] #$i -> HTTP $code"
    }
    $elapsedVals += $elapsed
}

# calculate average
if ($elapsedVals.Count -gt 0) {
    $avg = ($elapsedVals | Measure-Object -Average).Average
    $avg = [math]::Round($avg, 1)
} else {
    $avg = 0
}

Write-Host "total=$Total  success=$ok  fail=$fail  avg_latency=${avg}s"

# cleanup
Remove-Item -Path $TmpDir -Recurse -Force -ErrorAction SilentlyContinue
