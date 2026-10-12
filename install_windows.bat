@echo off
cd /d %~dp0
py -m venv .venv
call .venv\Scripts\activate.bat
python -m pip install --upgrade pip
pip install -r requirements.txt
if not exist uploads mkdir uploads
if not exist portal_updates mkdir portal_updates
if not exist portal_backups mkdir portal_backups
echo Installation completed. Configure environment variables before start_portal.bat.
pause
