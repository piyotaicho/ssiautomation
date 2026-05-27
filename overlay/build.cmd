@echo off
setlocal

set "CSC_PATH=C:\Windows\Microsoft.NET\Framework64\v4.0.30319\csc.exe"
if not exist "%CSC_PATH%" (
    set "CSC_PATH=C:\Windows\Microsoft.NET\Framework\v4.0.30319\csc.exe"
)

set "REF_DIR=C:\Program Files (x86)\Reference Assemblies\Microsoft\Framework\.NETFramework\v4.0"
if not exist "%REF_DIR%\PresentationFramework.dll" (
    set "REF_DIR=C:\Program Files (x86)\Reference Assemblies\Microsoft\Framework\.NETFramework\v4.8"
)

if not exist "%REF_DIR%\PresentationFramework.dll" (
    echo エラー: WPF 参照アセンブリが見つかりません。
    echo Developer Pack ^(.NET Framework 4.x Targeting Pack^) をインストールしてください。
    exit /b 1
)

pushd "%~dp0"

"%CSC_PATH%" /target:winexe /out:Overlay.exe ^
 /reference:"%REF_DIR%\PresentationFramework.dll" ^
 /reference:"%REF_DIR%\PresentationCore.dll" ^
 /reference:"%REF_DIR%\WindowsBase.dll" ^
 /reference:"%REF_DIR%\System.Xaml.dll" ^
 Overlay.cs

if errorlevel 1 (
    echo ビルド失敗
    popd
    exit /b 1
)

echo ビルド完了: Overlay.exe
popd
exit /b 0
