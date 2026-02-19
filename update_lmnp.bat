@echo off
cd /d "%~dp0"
git pull origin main
git add .
git commit -m "Mise a jour automatique"
git push origin main
pause