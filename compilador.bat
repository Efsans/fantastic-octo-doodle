@echo off
setlocal enabledelayedexpansion


echo ======================
echo   #       #   #######
echo  ###      #        #
echo  # #      #       #
echo #####     #      #
echo #   #     #     #
echo #   #     #    #######
echo ======================= 

echo ===============================
echo Limpando anterior... 
echo ===============================

cd /d C:\TEMP\Solidcod

if exist build rmdir /s /q build
if exist dist rmdir /s /q dist
if exist __pycache__ rmdir /s /q __pycache__
del /q *.spec 2>nul

set FILES=unificacaoteste.py interface.py

echo ===============================
echo Finalizando os executáveis anteriores...
for %%f in (%FILES%) do (
    set EXE_NAME=%%~nf.exe
    echo Tentando finalizar !EXE_NAME!...
    taskkill /f /im !EXE_NAME! >nul 2>nul
)
echo ===============================

echo ===========  
echo Compilando arquivos...  
echo ===========

for %%f in (%FILES%) do (
    echo Compilando %%f...
    pyinstaller --onefile --clean --noconsole %%f
)

echo ===============================
echo Deu certo aparentemente... :\/
pause