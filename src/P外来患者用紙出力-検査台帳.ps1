# A外来患者用紙出力-検査台帳.ps1 を起動する

# オーバーレイの起動
$overlay = Start-Process powershell.exe -PassThru -ArgumentList "-WindowStyle Hidden -PassThru -File $PSScriptRoot\Utility-Overlay.ps1 -Color '#55FF8000'"

Start-Sleep -Milliseconds 500 # オーバーレイが起動するのを少し待つ

# 処理を開始
try {
    & $PSScriptRoot\A外来患者用紙出力-検査台帳.ps1
} catch {
    Stop-Process -Id $overlay.Id -ErrorAction SilentlyContinue
    Write-Error $_

    & $PSScriptRoot\Utility-FullScreenDialog.ps1 -Message $_.Exception.Message -Title '検査台帳発行'
    exit 1
}

# 正常終了を「緑」で3秒間だけ通知する
$successPS = Start-Process powershell.exe -PassThru -ArgumentList "-WindowStyle Hidden -File $PSScriptRoot\Overlay.ps1 -Message '自動処理が正常に完了しました！' -Color '#4400FF00'"
Start-Sleep -Seconds 3
Stop-Process -Id $successPS.Id -ErrorAction SilentlyContinue
