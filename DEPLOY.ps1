$OutputEncoding = [Console]::InputEncoding = [Console]::OutputEncoding = [System.Text.UTF8Encoding]::new($false)
Set-Location $PSScriptRoot
$msg = Read-Host '커밋 메시지 입력'
if (-not $msg) { Write-Host '취소됨'; exit }
git add .
git commit -m $msg
git push
Write-Host '완료! Vercel 자동 반영됩니다'