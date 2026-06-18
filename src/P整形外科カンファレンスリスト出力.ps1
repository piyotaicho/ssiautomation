#
# 整形外科　術前術後カンファレンスのリストをつくる
# 術前：前の火曜日～月曜日
# 術後：こんどの火曜日～次の月曜日
#
$templatePath = "$PSScriptRoot\template-整形オペカンファレンス.xlsx"

# オーバーレイの起動
$overlay = Start-Process powershell.exe -PassThru -ArgumentList "-ExecutionPolicy bypass -WindowStyle Hidden -File $PSScriptRoot\Utility-Overlay.ps1 -Color '#55FF8000'"
Start-Sleep -Milliseconds 500 # オーバーレイが起動するのを少し待つ

# Step - 1 手術情報の取得

# 抽出する日付の計算
$startDate = Get-Date -Hour 0 -Minute 0 -Second 0
if (([int]$startDate.DayOfWeek) -gt 2) {
    $startDate = $startDate.AddDays(([int]$startDate.DayOfWeek)-8)
} else {
    $startDate = $startDate.AddDays(-(5+([int]$startDate.DayOfWeek)))
}
$endDate = $startDate.AddDays(13)

# 手術リストを取得
$opeList = @()
try {
    $opeList = ( & "$PSScriptRoot\A手術一覧-取得.ps1" -startDate $startDate.ToString('yyyy/MM/dd') -endDate $endDate.ToString('yyyy/MM/dd') -ka '11')
} catch {
    throw $_
}

if ($opeList.Count -eq 0) {
    # 対象症例なし
    Stop-Process -Id $overlay.Id -ErrorAction SilentlyContinue
    $successPS = Start-Process powershell.exe -PassThru -ArgumentList "-ExecutionPolicy bypass -WindowStyle Hidden -File $PSScriptRoot\Utility-Overlay.ps1 -Message '対象がないので印刷なしで終了しました' -Color '#4400FF00'"
    Start-Sleep -Seconds 3
    Stop-Process $successPS.Id -ErrorAction SilentlyContinue

    return 0
}

# Step - 2 リストの印刷

# カンファレンスの日のシリアル値で術前後を区分する
$meetingDate = $startDate.AddDays(7)
$thresholdSerial = ($meetingDate - (Get-Date '1900/1/1 0:0:0')).TotalDays + 1

# エクセルでの操作
$COMexcel = $null
$COMworkbook = $null
$COMsheet = $null

try {
    $COMexcel = New-Object -ComObject Excel.Application
    $COMworkbook = $COMexcel.Workbooks.Open($templatePath)
    if ($null -eq $COMworkbook) {
        throw 'テンプレートエクセルワークブックに接続できません'
    }
    $COMsheet = $COMworkbook.ActiveSheet

    # ワークシートに値を割り当ててゆく　1ページ20件　術前術後でわける
    $title = $COMsheet.Cells.Item(1, 8).Value2
    $title = $title -replace 'yyyy', $meetingDate.ToString('yyyy')
    $title = $title -replace 'mmm', $meetingDate.ToString('  M')
    $title = $title -replace 'ddd', $meetingDate.ToString('  d')
    $COMsheet.Cells.Item(1, 8).Value2 = $title

    # ループ
    $pageNumber = 1
    $pagebreak = $false
    $postope = $false
    $index = 0
    while ($index -lt $opeList.Count) {
        $row = 6
        while ($row -le 25) {
            if ($pagebreak) {
                $COMsheet.Cells.Item($row, 1).Value2 = ''
                $COMsheet.Cells.Item($row, 2).Value2 = ''
                $COMsheet.Cells.Item($row, 3).Value2 = ''
                $COMsheet.Cells.Item($row, 4).Value2 = ''
                $COMsheet.Cells.Item($row, 5).Value2 = ''
                $COMsheet.Cells.Item($row, 6).Value2 = ''
                #$COMsheet.Cells.Item($row, 7).Value2 = ''
                #$COMsheet.Cells.Item($row, 8).Value2 = ''
                #$COMsheet.Cells.Item($row, 10).Value2 = ''
                $row++
            } else {
                # 術前術後の切り替え
                if ((-not $postope) -and ($opeList[$index].'手術日シリアル' -gt $thresholdSerial)) {
                    $postope = $true
                    $pagebreak = $true
                    continue
                }
                # データーを埋める
                $COMsheet.Cells.Item($row, 1).Value2 = $opeList[$index].'手術日'
                $COMsheet.Cells.Item($row, 2).Value2 = $opeList[$index].'曜日'
                $COMsheet.Cells.Item($row, 3).Value2 = $opeList[$index].'患者姓'
                $COMsheet.Cells.Item($row, 4).Value2 = $opeList[$index].'患者名'
                $COMsheet.Cells.Item($row, 5).Value2 = $opeList[$index].'年齢'
                $COMsheet.Cells.Item($row, 6).Value2 = $opeList[$index].'性別'
                #$COMsheet.Cells.Item($row, 7).Value2 = ''
                #$COMsheet.Cells.Item($row, 8).Value2 = ''
                #$COMsheet.Cells.Item($row, 10).Value2 = ''

                $row++
                $index++

                # リスト最終行の動作切り替え
                if ($index -eq ($opeList.Count + 1)) {
                    $pagebreak = $true
                }
            }
        }

        # 印刷
        $COMsheet.PageSetup.RightFooter = "$($pageNumber++)"
        $COMsheet.PrintOut(1, 1)
        $pagebreak = $false
        [UIATools]::Sleep(500)
    }

    # COMを開放
    # excelを保存せずに終了
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($COMsheet) | Out-Null
    $COMworkbook.Close($false)
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($COMworkbook) | Out-Null
    $COMexcel.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($COMexcel) | Out-Null

    # 正常終了を「緑」で3秒間だけ通知する
    $pageNumber--
    Stop-Process -Id $overlay.Id -ErrorAction SilentlyContinue
    $successPS = Start-Process powershell.exe -PassThru -ArgumentList "-ExecutionPolicy bypass -WindowStyle Hidden -File $PSScriptRoot\Utility-Overlay.ps1 -Message '$($pageNumber)ページ出力しました.' -Color '#4400FF00'"
    Start-Sleep -Seconds 3
    Stop-Process $successPS.Id -ErrorAction SilentlyContinue

    return $pageNumber
} catch {
    Stop-Process -Id $overlay.Id -ErrorAction SilentlyContinue
    # エラー表示はクリックをまつ
    & $PSScriptRoot\Utility-FullScreenDialog.ps1 -Message $_.Exception.Message -Title '検査台帳発行'

    throw $_
} finally {
    # excelの残渣を片付ける
    try {
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($COMsheet) | Out-Null
    } catch {}
    try {
        $COMworkbook.Close($false)    
    } catch {}
    try {
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($COMworkbook) | Out-Null
    } catch {}
    try {
        $COMexel.Quit()
    } catch {}
    try {
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($COMexcel) | Out-Null
    } catch {}
}
