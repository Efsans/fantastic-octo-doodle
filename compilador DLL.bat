@echo off
setlocal enabledelayedexpansion

echo ======================================
echo Compilando suplemento SolidWorks C#
echo ======================================

cd /d "C:\TEMP\Solidcod\C#solidworks\"

set FILE=exemplo.cs interface.cs
set DLL_NAME=SW_teste.dll


if not exist exemplo.cs (
    echo Arquivo exemplo.cs nao encontrado!
    pause
    exit /b
)
if not exist interface.cs (
    echo Arquivo interface.cs nao encontrado!
    pause
    exit /b
)

:: Deleta compilações antigas
if exist "%DLL_NAME%" del "%DLL_NAME%"

:: Compilação
"C:\Windows\Microsoft.NET\Framework64\v4.0.30319\csc.exe" /target:library /out:%DLL_NAME% ^
/reference:"C:\Program Files\SOLIDWORKS Corp\SOLIDWORKS\SolidWorks.Interop.sldworks.dll" ^
/reference:"C:\Program Files\SOLIDWORKS Corp\SOLIDWORKS\SolidWorks.Interop.swconst.dll" ^
/reference:"C:\Program Files\SOLIDWORKS Corp\SOLIDWORKS\SolidWorks.Interop.swpublished.dll" ^
%FILE%

if exist "%DLL_NAME%" (
    echo ===============================
    echo Compilado com sucesso: %DLL_NAME%
) else (
    echo ===============================
    echo Erro na compilacao
)

pause