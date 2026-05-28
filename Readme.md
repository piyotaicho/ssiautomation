# SSIの電子カルテNewtonsでRPA

SSIの電子カルテNewtonsに対して各種RPA操作を行うものです。

Windows10以降の環境そのままで動作を目指して以下の構成を取っています

- Windows UI automation
- Windows PowerShell 5.1
- .NET Framework 4 

が必須環境で通常これらが欠損していることはありません。

# 構成
## Tools-AutomationCore.ps1

オートメーションに必要な各種操作をラップする関数を提供します。

PowerShellからWindows UI automationを使用した際、automation proxyが適切に動作しないOS側のバグがあります。
オートメーションの最初で行うアプリケーションウインドウの取得をCSharpを使ったアセンブリを使うことで、UIAutomationClientProvidersをインポートして適切な動作ができるようにしています。

Reference: https://qiita.com/mima_ita/items/3f2aa49fceca7496c587

## 
