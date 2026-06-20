# 診療情報提供料(1)の算定CSVから推測される紹介先を割り当てる
param (
    [switch]$noOverlay = $false
)

# アセンブリのロード
. "$PSScriptRoot\Tools-UserInteraction.ps1"

# ダイアログを開いてCSVファイルを指定
$sourceCSV = Invoke-FileSelectDialog -title '診療データ検索から出力されたCSVを指定してください' -filter 'CSVファイル|*.csv|すべてのファイル|*.*'

# automation
$result = @()
$overlay = $null

if ($null -ne $sourceCSV) {
    if (-not $noOverlay) {
        # オーバーレイの起動
        $overlay = Start-Process powershell.exe -PassThru -ArgumentList "-ExecutionPolicy bypass -WindowStyle Hidden -File $PSScriptRoot\Utility-Overlay.ps1 -Color '#55FF8000'"
        Start-Sleep -Milliseconds 500 # オーバーレイが起動するのを少し待つ
    }

    try {
        $result = (. "$PSScriptRoot\A医事課-診療情報提供料CSVと突合.ps1" $sourceCSV)
        # 結果を保存
        $filename = "診療情報提供先突合結果-$(Get-Date -Format 'yyyy-MM-dd').csv"
        $result | Export-Csv -Path  "$([System.Environment]::GetFolderPath('Desktop'))\$filename" -Encoding Default -NoTypeInformation # ConvertTo-Csv | Out-File -Encoding default -FilePath $filename

        if (($null -ne $overlay) -or (-not $noOverlay)) {
            Stop-Process -Id $overlay.Id -ErrorAction SilentlyContinue
        }

        # 結果を表示
        if (Confirm-UserYesNo -text "結果をデスクトップに`n$($filename)`nとして保存しました.`n結果を表示しますか？") {
            & start ([System.Environment]::GetFolderPath('Desktop') + $filename)
        }
    } catch {
        if ($null -ne $ovrelay) {
            Stop-Process -Id $overlay.Id -ErrorAction SilentlyContinue
        }
        Write-Error $_

        # エラー表示はクリックをまつ
        & $PSScriptRoot\Utility-FullScreenDialog.ps1 -Message $_.Exception.Message -Title '検査台帳発行'
    } finally {
        if ($null -ne $overlay) {
            Stop-Process -Id $overlay.Id -ErrorAction SilentlyContinue
        }
    }
}
