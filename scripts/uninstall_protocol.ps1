\
Write-Host "=== PptTplAgent 프로토콜(ppttpl://) 제거 (HKCU) ==="
$base = "HKCU:\Software\Classes\ppttpl"
if (Test-Path $base) {
  Remove-Item -Path $base -Recurse -Force
  Write-Host "제거 완료!"
} else {
  Write-Host "이미 제거되어 있습니다."
}
