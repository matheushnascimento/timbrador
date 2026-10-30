@echo off
rem Abre o Timbrador no Edge ou no Chrome, que sao os navegadores que gravam direto na pasta.
rem Deixe este arquivo na mesma pasta do timbrador.html.
setlocal
set "PAGINA=%~dp0timbrador.html"
if not exist "%PAGINA%" (
  echo Nao encontrei o timbrador.html nesta pasta.
  pause
  exit /b 1
)
reg query "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\msedge.exe" >nul 2>&1 && (
  start "" msedge "%PAGINA%"
  exit /b
)
reg query "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\chrome.exe" >nul 2>&1 && (
  start "" chrome "%PAGINA%"
  exit /b
)
echo Edge ou Chrome nao encontrados; abrindo no navegador padrao.
echo Sem um deles, o Timbrador nao grava direto na pasta: baixa o arquivo e a lista.
start "" "%PAGINA%"
