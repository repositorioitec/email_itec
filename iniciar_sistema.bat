@echo off
title Sistema Outlook IMAP - Webmail
echo ========================================================
echo  Iniciando Sistema de Gerenciamento de Email (IMAP/SMTP)
echo  Interface baseada no Microsoft Outlook
echo ========================================================
echo.

if not exist node_modules (
  echo Instalando dependencias necessarias...
  call npm install
)

echo Abrindo o navegador em http://localhost:4000 ...
start http://localhost:4000

echo Iniciando servidor Node.js...
node server.js

pause

