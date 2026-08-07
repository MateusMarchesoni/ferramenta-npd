@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo.
echo  GRAVAR - lanca os produtos marcados no Funil e na Priorizacao
echo.
"%~dp0programa\npd-tool.exe" gravar
echo.
echo ============================================================
echo  Foi feito um backup antes de gravar, na pasta "backups".
echo  O relatorio da importacao esta na pasta "relatorios".
echo ============================================================
echo.
pause
