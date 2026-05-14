@echo off
REM Wrapper de duplo-clique para Apply-Reports.ps1
REM Executa o script PowerShell com bypass de execution policy, mantem a janela aberta no fim.
setlocal
set "SCRIPT_DIR=%~dp0"
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%SCRIPT_DIR%Apply-Reports.ps1"
echo.
echo --- Fim. Pressione qualquer tecla para fechar ---
pause >nul
endlocal
