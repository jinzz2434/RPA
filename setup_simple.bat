@echo off
echo RPAシステム用のパッケージをインストールしています...

REM 必要なパッケージのインストール
pip install pandas openpyxl pyautogui pyperclip pillow

echo.
echo インストールが完了しました！
echo.
echo 以下のコマンドでRPAシステムを実行できます：
echo python rpa_system.py
echo python rpa_excel_system.py
echo.
pause
