Write-Host "RPAシステム用のconda環境を作成しています..." -ForegroundColor Green

# conda環境の作成
conda env create -f environment.yml

Write-Host ""
Write-Host "環境の作成が完了しました！" -ForegroundColor Green
Write-Host ""
Write-Host "以下のコマンドで環境をアクティベートしてください：" -ForegroundColor Yellow
Write-Host "conda activate rpa-system" -ForegroundColor Cyan
Write-Host ""
Write-Host "その後、以下のコマンドでRPAシステムを実行できます：" -ForegroundColor Yellow
Write-Host "python rpa_system.py" -ForegroundColor Cyan
Write-Host "python rpa_excel_system.py" -ForegroundColor Cyan
Write-Host ""

Read-Host "Enterキーを押して終了"
