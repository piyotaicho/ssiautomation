#
# 指定期間の手術予定リストを取得 なにも指定しなかったら１週間後の月～金
#
# -startDate yyyy/mm/dd -endDate yyyy/mm/dd 省略した場合は実行日の次の週の月～金
# -Ka 診療科名もしくは診療科コード(数値) - 省略した場合は 23(産婦人科)
#
param (
    [string]$startDate, # yyyy/mm/dd
    [string]$endDate,   # yyyy/mm/dd
    [string]$ka = '23'
)

# 追加アセンブリをロード
. "$PSScriptRoot\Tools-AutomationCore.ps1"
. "$PSScriptRoot\Tools-ExcelWindow.ps1"
. "$PSScriptRoot\Tools-MasterTables.ps1"

# 定数
$exe             = 'C:\SSI\exe\SjtYtPrtEx.exe'
$procName        = 'SjtYtPrtEx'
$NAME_main_window = '手術室予定表 (SjtYtPrtEx)'

$AID_date_start_edit  = '2'
$AID_date_end_edit    = '2'
$AID_ka_string_list   = '20'
$AID_print_button     = '22'
$AID_preview_button   = '23'
$AID_close_button     = '21'

$NamePreviewAbortDialog = ''
$AIDPreviewAbortButton = '5'

# オプションの確認
$startDateObject = $null
$endDateObject = $null

if ($startDate -match '^(\d{4})/(\d?\d)/(\d?\d)$') {
    $startDateObject = Get-Date -Year $Matches[1] -Month $Matches[2] -Day $Matches[3]
} else {
    if (-not $startDate) {
        $startDateObject = Get-Date
        $startDateObject = $startDateObject.AddDays(8 - [int]$startDateObject.DayOfWeek)
    } else {
        throw '開始日は yyyy/mm/dd で指定してください'
    }
}

if ($endDate -match '^(\d{4})/(\d?\d)/(\d?\d)$') {
    $endDateObject = Get-Date -Year $Matches[1] -Month $Matches[2] -Day $Matches[3]
} else {
    if (-not $endDate) {
        $endDateObject = Get-Date
        $endDateObject = $endDateObject.AddDays(13 - [int]$endDateObject.DayOfWeek)
    } else {
        throw '終了日は yyyy/mm/dd で指定してください'
    }
}

#Write-Host "Start $($startDateObject.ToString())"
#Write-Host "End $($endDateObject.ToString())"

if ($endDateObject -lt $startDateObject) {
    throw '終了日は開始日よりもあとにしてください'
}

# 診療科チェック
if ($ka) {
    if ($ka -match '^\d+$') {
        if($null -eq $MasterKaCode[[Int]$ka]) {
            throw "診療科コード $ka は不正な指定です."
        }
        # 検索用に数値は２桁に正規化
        $ka = ([int]$ka).toString('00')
    }
    #Write-Host "診療科$ka"
}

# 指定用の和暦を生成
$cultureInfo = New-Object System.Globalization.CultureInfo('ja-JP', $true)
$cultureInfo.DateTimeFormat.Calendar = New-Object System.Globalization.JapaneseCalendar

$startDateString = $startDateObject.ToString('yyMMdd', $cultureInfo)
$endDateString = $endDateObject.ToString('yyMMdd', $cultureInfo)

# オートメーションの開始
$dataRows = @()
try {
    # 手術予定を起動して取得
    $appWindow = GetAppWindow -Name $NAME_main_window -ExecutablePath $exe
    if ($null -eq $appWindow) {
        throw '手術予定アプリケーションを起動できませんでした.'
    }

    # アプリのPane[1]（index 1 = 右側）配下の AutomationId "2" の Edit
    $panes = Get-UIAChildPanes -Parent $appWindow
    if ($panes.Count -ne 2) {
        throw 'アプリケーションに異常があります（ペイン数が想定外です）.'
    }

    # --- 開始日付を入力 ---
    $editStart = Get-UIAEdit -Parent $panes[1] -Id $AID_date_start_edit
    if ($null -eq $editStart) { throw 'アプリケーションに異常があります（開始日付フィールドが見つかりません）.' }
    Set-UIAValue $editStart $startDateString | Out-Null

    # --- 終了日付を入力 ---
    $editEnd = Get-UIAEdit -Parent $panes[0] -Id $AID_date_end_edit
    if ($null -eq $editEnd) { throw 'アプリケーションに異常があります（終了日付フィールドが見つかりません）.' }
    Set-UIAValue $editEnd $endDateString | Out-Null

    # --- 診療科指定があれば選択, 選択できなかったらデフォルトのまま ---
    if ($ka) {
        $kaList = Get-UIAList -Parent $appWindow -Id $AID_ka_string_list
        if ($null -eq $kaList) { throw 'アプリケーションに異常があります（科リストが見つかりません）.' }

        Set-UIAListSelection $kaList $ka -UseMatch
    }

    [UIATools]::Sleep(800)

    # プレビュー起動前のhWnd != 0 なエクセルのPIDを取得しておく
    $preservedExcelPIDs = (Get-Process -Name exel -ErrorAction SilentlyContinue | Where-Object MainWindowHandle -ne 0).Id

    # --- プレビューを起動(オートメーションでボタンを操作するとなにかコンフリクトするのでアプリケーションウインドウにキーストロークを送る) ---
    Set-UIAValue $appWindow '%V' -Force -OmitEscape

    [UIATools]::Sleep(2000)

    # エラーダイアログの有無を確認
    $checkWindows = Get-UIAChildWindows -Parent $appWindow
    if ($null -ne $checkWindows -and $checkWindows.Count -gt 0 -and $null -ne $checkWindows[0]) {
        $textMessasge = Get-UIAText -Parent $checkWindows[0]
        $btnOk = Get-UIAButton -Parent $checkWindows[0]
        if ($null -eq $textMessasge) {
            $errorMessage = 'アプリケーションにエラーが発生しました'
        } else {
            $errorMessage = Get-UIAValue $textMessasge
        }
        if ($null -eq $btnOk) {
            throw 'アプリケーションに異常があります（不正なダイアログです）.'
        }
        Invoke-UIAElement $btnOk | Out-Null

        if ($errorMessage -like '*ありません*') {
            # アプリケーションを終了する
            Set-UIAValue $appWindow '{F4}' -Force -OmitEscape | Out-Null

            # データーなしで正常終了
            return @()
        }
        throw $errorMessage
    }

    # Excelのオートメーション
    $COMexcel = $null
    $COMworkbook = $null
    $COMworksheet = $null

    try {
        $automationtimeout = [System.Diagnostics.Stopwatch]::StartNew()
        while ($automationtimeout.Elapsed.TotalSeconds -lt 300) {
            # エクセルのプロセスを取得
            $previewExcel = @()
            $excelWait = [System.Diagnostics.Stopwatch]::StartNew()
            while ($excelWait.Elapsed.TotalSeconds -lt 30) {
                $previewExcel = (Get-Process -Name excel -ErrorAction SilentlyContinue | Where-Object MainWindowHandle -NE 0 | Where-Object Id -NotIn $preservedExcelPIDs)
                if ($previewExcel.Count -eq 1) {
                    break
                }
                if ($previewExcel.Count -eq 0) {
                    if (($dataRows.Count % 20) -gt 0) {
                        # 新規エクセルが開かなくても既に取得したデーターがあったら処理は終了と見做す
                        # ただし20件ずつでちょうどの場合次があるかもしれないので待機する
                        break
                    }
                    [UIATools]::Sleep(2000)
                } else {
                    # ２個以上の新規プロセスがある状態は異常
                    throw '複数の新規エクセルが起動しています.アプリケーションの状態を確認してください.'
                }
            }
            
            # エクセルのオートメーションを開始
            if ($previewExcel.Count -eq 0) {
                break
            }
            $COMexcel = [ExcelWindowFactory]::GetExcelApplicationFromHwnd($previewExcel[0].MainWindowHandle)
            $COMworkbook = $COMexcel.ActiveWorkbook
            $COMworksheet = $COMworkbook.Worksheets.Item(2)

            # データーのコピー(1ファイルあたり20件までしかない）
            $row = 2
            while ($row -lt 25) {
                if ($COMworksheet.Cells.Item($row, 1).Text) {
                    $dataRows += ,@{
                        手術日 = $COMworksheet.Cells.Item($row, 1).Text
                        手術日シリアル  = $COMworksheet.Cells.Item($row, 1).Value2
                        曜日 = $COMworksheet.Cells.Item($row, 2).Value2
                        ID = ([Int]($COMworksheet.Cells.Item($row, 26).Value2)).toString('00000000')
                        患者姓 = $COMworksheet.Cells.Item($row, 3).Value2
                        患者名 = $COMworksheet.Cells.Item($row, 4).Value2
                        性別 = if ($COMworksheet.Cells.Item($row, 5).Value2 -eq '1') { '男' } else { '女' }
                        年齢 = $COMworksheet.Cells.Item($row, 6).Value2
                        診療科 = $COMworksheet.Cells.Item($row, 7).Value2
                        術式 = $COMworksheet.Cells.Item($row, 12).Value2
                    }
                    $row++
                } else {
                    break
                }
            }

            # Excelを閉じてCOMの開放
            [System.Runtime.InteropServices.Marshal]::ReleaseComObject($COMworksheet) | Out-Null
            $COMworkbook.Close($false)
            $COMexcel.Quit()
            [System.Runtime.InteropServices.Marshal]::ReleaseComObject($COMworkbook) | Out-Null
            [System.Runtime.InteropServices.Marshal]::ReleaseComObject($COMexcel) | Out-Null

            [UIATools]::Sleep(500)
        }
    } catch {
        #Write-Error $_
        throw $_
    } finally {
        # COMをダメ押しで開放する
        try {
            [System.Runtime.InteropServices.Marshal]::ReleaseComObject($COMworksheet) | Out-Null
        } catch {}
        try {
            [System.Runtime.InteropServices.Marshal]::ReleaseComObject($COMworkbook) | Out-Null
        } catch {}
        try {
            [System.Runtime.InteropServices.Marshal]::ReleaseComObject($COMexcel) | Out-Null
        } catch {}
    }

    # アプリケーションを終了する
    Set-UIAValue $appWindow '{F4}' -Force -OmitEscape | Out-Null

    # 結果を返す
    $dataRows
} catch {
    throw $_
} finally {
  if ($null -ne $appWindow) {
  }
}
