chcp 1252>nul
@echo off
cls

echo.
echo Ajout de la variable en cours..
setx ConcatFiles "T:\- 4 Suivi Appuis\26_MACROS\GO\ConcatFiles\ConcatFiles.exe"
echo.
start "" "T:\- 4 Suivi Appuis\26_MACROS\GO\ConcatFiles\ConcatFiles_REG.reg"

pause