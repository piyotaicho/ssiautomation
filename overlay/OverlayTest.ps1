# 2秒間だけ緑色で表示して自動終了させる例（PowerShellで待機）
Start-Process .\Overlay.exe -ArgumentList "正常終了しました", "#4400FF00"
Start-Sleep -Seconds 2
Get-Process Overlay | Stop-Process
