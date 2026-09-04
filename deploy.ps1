# 自動同期スクリプト (deploy.ps1)
Write-Host "1. GASへコードを反映中..." -ForegroundColor Green
clasp push

Write-Host "`n2. GitHubへ変更を保存中..." -ForegroundColor Green
git add .
$msg = Read-Host "コミットメッセージを入力してください（空欄の場合は'update'）"
if ([string]::IsNullOrWhiteSpace($msg)) { $msg = "update" }
git commit -m $msg
git push

Write-Host "`n全自動同期が完了しました！" -ForegroundColor Cyan