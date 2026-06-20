# 指定した電算コードの算定リストを取得する
# 
param (
    
)

. "$PSScriptRoot\Tools-AutomationCore.ps1"

#
$appWindow = $null
$dateStart = '20260601'
$dateEnd = '20260610'
$densanCode = '180016110'

try {
    # & "$PSScriptRoot\Aエントランスから起動.ps1" '外来医事業務' '診療データ検索'
    & "$PSScriptRoot\Aエントランスから起動.ps1" '医事統計業務' '診療データ検索'
    [UIATools]::Sleep(3000)

    $appWindow = GetAppWindow -Name '診療データ検索'
} catch {
    Write-Error $_
    throw '診療データ検索を起動・確認できませんでした.'
}

try {
    # 検索条件を設定
    $mainPane = Get-UIAPane -Parent $appWindow -Id 'pnlMain'
    $conditionPane = Get-UIAPane -Parent $mainPane -Id 'pnlCondition1'

    # 期間開始日
    $startEdit = Get-UIAEdit -Parent (Get-UIAPane -Parent $conditionPane -Id 'dtxtKikanSt') -Id 'txtDate'
    Set-UIAValue $startEdit $dateStart | Out-Null

    # 期間終了日
    $endEdit = Get-UIAEdit -Parent (Get-UIAPane -Parent $conditionPane -Id 'dtxtKikanEn') -Id 'txtDate'
    Set-UIAValue $endEdit $dateEnd | Out-Null

    # 電算コードを指定
    $codeEditPane = Get-UIAPane -Parent $conditionPane -Id 'txtIjiTencd'
    Set-UIAValue $codeEditPane "$densanCode{enter}" -Force -OmitEscape | Out-Null

    [UIATools]::Sleep(1000)

    # 検索を開始
    Set-UIAValue $appWindow '%S' -Force -OmitEscape | Out-Null
} catch {
    Write-Error $_
    throw '検索条件の設定ができませんでした.'
}

# 検索結果を待つ 3分間まで ポーリングは5秒ごと
$ichiranWindow = $null
$stopwatch = [System.Diagnostics.Stopwatch]::StartNew()
while ($stopwatch.Elapsed.TotalSeconds -lt 180) {
    [UIATools]::Sleep(5000)
    $childWindows = Get-UIAChildWindows -Parent $appWindow -Name '検索一覧'
    if ($null -ne $childWindows -and $childWindows.Count -gt 0) {
        $ichiranWindow = $childWindows[0]
        break
    }
}
if ($null -eq $ichiranWindow) {
    throw '検索一覧の待機時間(180秒)がタイムアウトしました.'
}

try {
    # CSV保存へ
    Invoke-UIAElement (Get-UIAButton $ichiranWindow -Id 'btnCSV') | Out-Null

    [UIATools]::Sleep(500)

    # Csv保存設定
    $csvWindow = (Get-UIAChildWindows $ichiranWindow -Name 'CSV出力設定')[0]

    # CSV(2)を選択
    $csvtype2Radio = Get-UIARadio $csvWindow -Id 'optCSV2'
    if (-not (Get-UIAValue $csvtype2Radio)) {
        Invoke-UIAElement $csvtype2Radio | Out-Null
    }

    # 保存
    Invoke-UIAElement (Get-UIAButton $ichiranWindow -Id 'btnOK') | Out-Null
    [UIATools]::Sleep(500)

    # 名前をつけて保存ダイアログがでる
    $filesaveDialog = Get-UIAWindow $ichiranWindow -Name '名前を付けて保存'

    Set-UIAValue (Get-UIAEdit $filesaveDialog -Id '1001') 'C:\Users\ssiuser\Desktop\test.csv'
    Invoke-UIAElement (Get-UIAButton $filesaveDialog -Id '1')
    [UIATools]::Sleep(1000)
} catch {
    Write-Error $_
    throw '結果の保存に失敗しました.'
}

try {
    # 完了ダイアログがシステムモーダルで表示されるので閉じる
    $finishDialog = Get-UIAWindow ([UIATools]::RootElement) -Id 'SSIMessageDialog'
    Invoke-UIAElement (Get-UIAButton $finishDialog -Id 'Button1') | Out-Null

    # 検索一覧を閉じる
    Invoke-UIAElement (Get-UIAButton $ichiranWindow -Id 'btnCancel') | Out-Null
    [UIATools]::Sleep()

    # アプリケーションを閉じる
    Set-UIAValue $appWindow '{F4}' -Force -OmitEscape | Out-Null
} catch {
    Write-Error $_
}