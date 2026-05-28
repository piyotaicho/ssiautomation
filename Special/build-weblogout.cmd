@echo off
setlocal

set "CSC_PATH=C:\Windows\Microsoft.NET\Framework64\v4.0.30319\csc.exe"
if not exist "%CSC_PATH%" (
    set "CSC_PATH=C:\Windows\Microsoft.NET\Framework\v4.0.30319\csc.exe"
)

set "REF_DIR=C:\Program Files (x86)\Reference Assemblies\Microsoft\Framework\.NETFramework\v4.0"
if not exist "%REF_DIR%\UIAutomationClient.dll" (
    set "REF_DIR=C:\Program Files (x86)\Reference Assemblies\Microsoft\Framework\.NETFramework\v4.8"
)

if not exist "%REF_DIR%\UIAutomationClient.dll" (
    echo Error: .NET Framework Reference Assemblies not found.
    echo Please install .NET Framework 4.x Developer Pack.
    exit /b 1
)

pushd "%~dp0"

"%CSC_PATH%" /target:exe /out:WebLogout.exe ^
 /reference:"%REF_DIR%\UIAutomationClient.dll" ^
 /reference:"%REF_DIR%\UIAutomationTypes.dll" ^
 WebLogout.cs

if errorlevel 1 (
    echo Build failed.
    popd
    exit /b 1
)

echo Build succeeded: WebLogout.exe
popd
exit /b 0