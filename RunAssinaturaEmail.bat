REM Atualizado em: 02/03/2025 por Jefferson Aparecido Sadoski
REM Script para execuatar o script para aplicação de assinatura no Outlook WEB, New Outlook e Outlook Desktop
REM Contato: jefferson_sadoski@hotmail.com -WindowStyle Hidden

REM %windir%\system32\WindowsPowerShell\v1.0\powershell.exe -ExecutionPolicy Bypass -File \\contoso.sa\NETLOGON\Assinatura\assinaturaWebDesktop.ps1

@echo off
%windir%\system32\WindowsPowerShell\v1.0\powershell.exe -NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -Command "Start-Process powershell.exe -ArgumentList '-NoProfile -ExecutionPolicy Bypass -File \"\\contoso.sa\NETLOGON\Assinatura\assinaturaWebDesktop.ps1\"' -WindowStyle Hidden"
