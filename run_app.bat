@echo off
title VIEW SUMMARY STUDIO
echo ====================================================
echo  Starting VIEW SUMMARY STUDIO ^& Backend Engine
echo ====================================================
echo.
cd /d "%~dp0backend"
echo Initializing knowledge database from data\raw_boq...
python -c "import server; res = server.scan_workspace_folder(); print('Scanned:', res['scanned_files'], '| Trained:', len(res['trained_files']))"
echo.
echo Launching Server at http://127.0.0.1:8000 ...
echo Opening application in your default browser...
start http://127.0.0.1:8000
python -m uvicorn server:app --host 127.0.0.1 --port 8000
pause
