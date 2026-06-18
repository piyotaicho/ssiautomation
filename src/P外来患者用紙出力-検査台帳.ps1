# A外来患者用紙出力-検査台帳.ps1 を起動する
. .\Tools-UserInteraction.ps1

# Excelをすべて閉じる
$excelProcs = Get-Process -Name excel -ErrorAction SilentlyContinue
if ($excelProcs.Count -gt 0) {
    $result = Confirm-UserOkCancel '現在動作しているExcelをすべて強制終了します、よろしければOK、処理を中止するにはキャンセルを押してください'
    if ($result) {
        $excelProcs | Stop-Process -ErrorAction SilentlyContinue | Out-Null
    } else {
        exit 0
    }
}

# オーバーレイの起動
$overlay = Start-Process powershell.exe -PassThru -ArgumentList "-ExecutionPolicy bypass -WindowStyle Hidden -File $PSScriptRoot\Utility-Overlay.ps1 -Color '#55FF8000'"
Start-Sleep -Milliseconds 500 # オーバーレイが起動するのを少し待つ

# 処理を開始
try {
    # 実際の処理スクリプトを起動
    & $PSScriptRoot\A外来患者用紙出力-検査台帳.ps1

    # 正常終了を「緑」で3秒間だけ通知する
    Stop-Process -Id $overlay.Id -ErrorAction SilentlyContinue
    $successPS = Start-Process powershell.exe -PassThru -ArgumentList "-ExecutionPolicy bypass -WindowStyle Hidden -File $PSScriptRoot\Utility-Overlay.ps1 -Message '自動処理が正常に完了しました！' -Color '#4400FF00'"
    Start-Sleep -Seconds 3
    Stop-Process $successPS.Id -ErrorAction SilentlyContinue
} catch {
    Stop-Process -Id $overlay.Id -ErrorAction SilentlyContinue
    Write-Error $_

    # エラー表示はクリックをまつ
    & $PSScriptRoot\Utility-FullScreenDialog.ps1 -Message $_.Exception.Message -Title '検査台帳発行'
}
