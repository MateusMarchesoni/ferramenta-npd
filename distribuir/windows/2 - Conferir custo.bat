@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo.
echo  CONFERIR CUSTO - calcula o custo dos produtos marcados
echo.
"%~dp0programa\npd-tool.exe" conferir
echo.
echo ============================================================
echo  O custo tambem aparece na coluna M da aba Candidatos.
echo  Se estiver certo, clique em "3 - Gravar no funil".
echo ============================================================
echo.
pause
