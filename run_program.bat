@echo off
REM 仮想環境を有効化
call venv\Scripts\activate
REM Pythonスクリプトを実行
python kgcareer016.py
REM 終了時に仮想環境を無効化（必要に応じて）
deactivate
pause
