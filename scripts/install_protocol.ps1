\
param(
  [string]$ExePath = ""
)

Write-Host "=== PptTplAgent 프로토콜(ppttpl://) 등록 (HKCU) ==="

if ([string]::IsNullOrWhiteSpace($ExePath)) {
  $ExePath = Read-Host "PptTplAgent.exe 전체 경로를 입력하세요 (예: C:\Tools\PptTplAgent\PptTplAgent.exe)"
}

if (!(Test-Path $ExePath)) {
  Write-Error "파일이 존재하지 않습니다: $ExePath"
  exit 1
}

# 레지스트리 경로
$base = "HKCU:\Software\Classes\ppttpl"
New-Item -Path $base -Force | Out-Null
New-ItemProperty -Path $base -Name "(Default)" -Value "URL:PptTpl Protocol" -Force | Out-Null
New-ItemProperty -Path $base -Name "URL Protocol" -Value "" -Force | Out-Null

$cmdKey = Join-Path $base "shell\open\command"
New-Item -Path $cmdKey -Force | Out-Null

# "%1" 전달
$cmd = "`"$ExePath`" `"%1`""
New-ItemProperty -Path $cmdKey -Name "(Default)" -Value $cmd -Force | Out-Null

Write-Host "등록 완료!"
Write-Host "테스트: 크롬 주소창에  ppttpl://ping"
