@echo off
REM Atualiza data.json a partir das APIs + manual.xlsx, comita e empurra ao GitHub.
REM Uso:  update.bat                       (mensagem padrao com data/hora)
REM       update.bat "minha mensagem"      (mensagem custom)

setlocal

REM 1) Rodar o build (gera data.json)
echo === Rodando o build ===
python scripts\build_data.py
if errorlevel 1 (
    echo.
    echo Build falhou. Nada foi commitado.
    exit /b 1
)

REM 2) Stage dos arquivos esperados
echo.
echo === Preparando commit ===
git add manual\panorama_manual.xlsx data.json

REM Sair se nao houve nenhuma mudanca
git diff --cached --quiet
if not errorlevel 1 (
    echo Nenhuma mudanca em data.json ou manual.xlsx. Nada para commitar.
    exit /b 0
)

REM 3) Commit
if "%~1"=="" (
    REM Mensagem padrao com data/hora
    for /f "tokens=2 delims==" %%i in ('wmic os get localdatetime /value') do set DT=%%i
    set "STAMP=%DT:~0,4%-%DT:~4,2%-%DT:~6,2% %DT:~8,2%:%DT:~10,2%"
    git commit -m "Atualiza dados (%STAMP%)"
) else (
    git commit -m "%~1"
)

REM 4) Push
echo.
echo === Empurrando pro GitHub ===
git push
if errorlevel 1 (
    echo.
    echo Push falhou. Commit local feito; rode 'git push' manualmente quando puder.
    exit /b 1
)

echo.
echo OK. GitHub Pages republica em 30-90s. Ctrl+Shift+R no navegador para recarregar.
