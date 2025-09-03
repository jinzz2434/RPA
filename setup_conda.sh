#!/bin/bash

echo "RPAシステム用のconda環境を作成しています..."

# conda環境の作成
conda env create -f environment.yml

echo ""
echo "環境の作成が完了しました！"
echo ""
echo "以下のコマンドで環境をアクティベートしてください："
echo "conda activate rpa-system"
echo ""
echo "その後、以下のコマンドでRPAシステムを実行できます："
echo "python rpa_system.py"
echo "python rpa_excel_system.py"
echo ""
