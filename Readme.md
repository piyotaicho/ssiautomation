# SSIの電子カルテNewtonsでRPA

## 事前環境整備
Windows UI automationのDLLをラップするために、.NETライブラリの [UIDeskAutomation](https://github.com/ddeltasolutions/UIDeskAutomation) を使用します。

- 利点
  - Automationを用意するのに必要なプロセス周りのユーティリティ関数が用意されている
  - 面倒なInvokeまわりをラップしてくれている
- 欠点
  - 完全無保証の3rd partyユーティリティです

開発環境(VSCode + PowerShell Pro Tools extension)でフォームエディタを動作させるために、PowerShell 5.1を利用する様に以下のワークスペース設定が必要です。
```
{
    "terminal.integrated.defaultProfile.windows": "Windows PowerShell",
    "powershell.powerShellDefaultVersion": "Windows Powershell (x64)"
}
```