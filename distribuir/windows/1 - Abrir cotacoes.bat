@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo.
echo  ABRIR COTACOES - le os arquivos da pasta "cotacoes" e monta a aba Candidatos
echo.
"%~dp0programa\npd-tool.exe" abrir
echo.
echo ============================================================
echo  Agora abra a planilha, va na aba Candidatos, marque com x
echo  os produtos e preencha NCM e Marca. Depois salve, feche,
echo  e clique em "2 - Conferir custo".
echo ============================================================
echo.
pause
