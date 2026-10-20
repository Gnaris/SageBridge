@echo off
setlocal
chcp 65001 >nul
title Compilation de SageBridge
cd /d "%~dp0"
rem ───────────────────────────────────────────────────────────────────────
rem  Compile SageBridge.exe puis l'installateur SageBridge-Setup.exe (dans bin\)
rem  avec le csc.exe du .NET Framework 4 fourni par Windows : rien a installer.
rem  Resultat recopie dans ..\Installateur et ..\Documentation s'ils existent.
rem ───────────────────────────────────────────────────────────────────────
set CSC=%WINDIR%\Microsoft.NET\Framework64\v4.0.30319\csc.exe
if not exist "%CSC%" set CSC=%WINDIR%\Microsoft.NET\Framework\v4.0.30319\csc.exe
if not exist bin mkdir bin
set LOG=%~dp0build.log
echo Compilation du %date% %time% avec %CSC% > "%LOG%"
set REFS=/r:System.dll /r:System.Core.dll /r:System.Windows.Forms.dll /r:System.Drawing.dll /r:Microsoft.CSharp.dll
set OPTS=/nologo /target:winexe /platform:anycpu /optimize+ /codepage:65001 /win32icon:logo\SageBridge.ico /win32manifest:app.manifest

echo [1/3] SageBridge.exe
echo [1/3] SageBridge.exe >> "%LOG%"
"%CSC%" %OPTS% /out:bin\SageBridge.exe %REFS% /r:System.Web.Extensions.dll SageBridge.cs >> "%LOG%" 2>&1
if errorlevel 1 goto :echec

for /f "usebackq delims=" %%v in (`powershell -NoProfile -Command "$v=[Reflection.AssemblyName]::GetAssemblyName('%~dp0bin\SageBridge.exe').Version; '{0}.{1}.{2}' -f $v.Major,$v.Minor,$v.Build"`) do set VER=%%v
echo Version %VER% >> "%LOG%"
> bin\SetupVersion.cs echo [assembly: System.Reflection.AssemblyVersion("%VER%.0")]
>> bin\SetupVersion.cs echo [assembly: System.Reflection.AssemblyFileVersion("%VER%.0")]
>> bin\SetupVersion.cs echo [assembly: System.Reflection.AssemblyInformationalVersion("%VER%")]

copy /y logo\SageBridge-256.png bin\logo.png >nul
set DOCRES=
if exist "docs\SageBridge - Documentation.pdf" copy /y "docs\SageBridge - Documentation.pdf" bin\Documentation.pdf >nul
if exist bin\Documentation.pdf set DOCRES=/resource:bin\Documentation.pdf,Documentation.pdf

echo [2/3] SageBridge-Setup.exe (version %VER%)
echo [2/3] SageBridge-Setup.exe >> "%LOG%"
"%CSC%" %OPTS% /out:bin\SageBridge-Setup.exe %REFS% /resource:bin\SageBridge.exe,SageBridge.exe /resource:bin\logo.png,logo.png %DOCRES% Setup.cs bin\SetupVersion.cs >> "%LOG%" 2>&1
if errorlevel 1 goto :echec

echo [3/3] Empreinte SHA-256
powershell -NoProfile -Command "$h=(Get-FileHash -Algorithm SHA256 '%~dp0bin\SageBridge-Setup.exe').Hash.ToLower(); Set-Content -Encoding ASCII -NoNewline -Path '%~dp0bin\SageBridge-Setup.exe.sha256' -Value ($h + '  SageBridge-Setup.exe')" >> "%LOG%" 2>&1

if exist "..\Installateur" copy /y bin\SageBridge-Setup.exe "..\Installateur\SageBridge-Setup.exe" >nul
if exist "..\Documentation" if exist "docs\SageBridge - Documentation.pdf" copy /y "docs\SageBridge - Documentation.pdf" "..\Documentation\SageBridge - Documentation.pdf" >nul
echo OK %VER% >> "%LOG%"
echo.
echo Compilation reussie : SageBridge %VER%
echo   bin\SageBridge-Setup.exe  (installateur a partager)
timeout /t 8 >nul
exit /b 0

:echec
echo ECHEC >> "%LOG%"
echo.
echo La compilation a echoue : voir build.log
timeout /t 600 >nul
exit /b 1
