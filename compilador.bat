@echo off
echo ===============================
echo Limpando anterior...
echo ===============================

cd /d C:\TEMP\Solidcod


rmdir /s /q build
rmdir /s /q dist
rmdir /s /q __pycache__
del /q *.spec

set FILES=unificacaoteste.py interface.py 

echo ===========
echo Compilando arquivos...?!
echo ===========

for %%f in (%FILES%) do (
    echo Compilando %%f...
    pyinstaller --onefile --clean --noconsole %%f
)

echo ===============================
echo Deu Certo Aparemtemente
pause
