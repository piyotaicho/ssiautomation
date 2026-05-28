# 来週の産婦人科の手術予定を印刷する
#

. "$PSScriptRoot\Tools-AutomationCore.ps1"

Add-Type -AssemblyName PresentationFramework

# 定数
$exe             = 'C:\SSI\exe\SjtYtPrtEx.exe'
$procName        = 'SjtYtPrtEx'
$NAME_main_window = '手術室予定表 (SjtYtPrtEx)'

$AID_date_start_edit  = '2'
$AID_date_end_edit    = '2'
$AID_ka_string_list   = '20'
$AID_print_button     = '22'
$AID_close_button     = '21'

$NAME_print_window    = '手術室予定表'
$AID_start_print_button = '4'

$NAME_confirm_window  = '確認！'
$AID_confirm_ok_button = '2'

# --- 来週月曜〜土曜の日付文字列を算出（和暦 yyMMdd）---
$cultureInfo = New-Object System.Globalization.CultureInfo('ja-JP', $true)
$cultureInfo.DateTimeFormat.Calendar = New-Object System.Globalization.JapaneseCalendar

$today        = Get-Date
$nextMonday   = $today.AddDays(8 - [int]$today.DayOfWeek)
$start_date   = $nextMonday.ToString('yyMMdd', $cultureInfo)
$end_date     = $nextMonday.AddDays(5).ToString('yyMMdd', $cultureInfo)

# --- アプリ起動 ---
$appWindow = GetAppWindow -Name $NAME_main_window -ExecutablePath $exe
if ($null -eq $appWindow) {
    Write-Error 'アプリケーションを起動できませんでした.'
    exit 1
}

try {
    # --- 開始日付を入力 ---
    # アプリのPane[1]（index 1 = 右側）配下の AutomationId "2" の Edit
    $panes = Get-UIAChildPanes -Parent $appWindow
    if ($panes.Count -ne 2) {
        throw 'アプリケーションに異常があります（ペイン数が想定外です）.'
    }

    $editStart = Get-UIAEdit -Parent $panes[1] -Id $AID_date_start_edit
    if ($null -eq $editStart) { throw 'アプリケーションに異常があります（開始日付フィールドが見つかりません）.' }
    Set-UIAValue $editStart $start_date

    # --- 終了日付を入力 ---
    $editEnd = Get-UIAEdit -Parent $panes[0] -Id $AID_date_end_edit
    if ($null -eq $editEnd) { throw 'アプリケーションに異常があります（終了日付フィールドが見つかりません）.' }
    Set-UIAValue $editEnd $end_date

    # --- 診療科リストから「23 産婦人科」を選択 ---
    $kaList = Get-UIAList -Parent $appWindow -Id $AID_ka_string_list
    if ($null -eq $kaList) { throw 'アプリケーションに異常があります（科リストが見つかりません）.' }

    $kaItem = Get-UIAControl -Parent $kaList -Name '23 産婦人科' -Type ([Windows.Automation.ControlType]::ListItem)
    if ($null -eq $kaItem) { throw 'アプリケーションに異常があります（産婦人科の項目が見つかりません）.' }

    $selectionPattern = $null
    if ($kaItem.TryGetCurrentPattern([Windows.Automation.SelectionItemPattern]::Pattern, [ref]$selectionPattern)) {
        $selectionPattern.Select()
    } else {
        throw '診療科の選択に失敗しました.'
    }

    # --- 印刷ボタンを押す ---
    $printButton = Get-UIAButton -Parent $appWindow -Id $AID_print_button
    if ($null -eq $printButton) { throw 'アプリケーションに異常があります（印刷ボタンが見つかりません）.' }
    Invoke-UIAElement $printButton

    [UIATools]::sleep(1000)

    # --- 子ウィンドウの種類で分岐 ---
    $childWindows = Get-UIAChildWindows -Parent $appWindow
    if ($childWindows.Count -eq 0) {
        throw 'アプリケーションに異常があります（子ウィンドウが開きません）.'
    }
    $childWindow = $childWindows[0]
    $childTitle  = $childWindow.Current.Name

    if ($childTitle -eq $NAME_print_window) {
        # 症例あり：印刷ダイアログで印刷を実行
        $startPrintButton = Get-UIAButton -Parent $childWindow -Id $AID_start_print_button
        if ($null -eq $startPrintButton) { throw 'アプリケーションに異常があります（印刷開始ボタンが見つかりません）.' }
        Invoke-UIAElement $startPrintButton
        [UIATools]::sleep(5000) # 印刷処理待ち

    } elseif ($childTitle -eq $NAME_confirm_window) {
        # 症例なし：確認ダイアログを閉じる
        $okButton = Get-UIAButton -Parent $childWindow -Id $AID_confirm_ok_button
        if ($null -eq $okButton) { throw 'アプリケーションに異常があります（OKボタンが見つかりません）.' }
        Invoke-UIAElement $okButton
        [UIATools]::sleep(1000)

        [System.Windows.MessageBox]::Show('来週の手術予定がありませんでした.', '報告', 'OK') | Out-Null

    } else {
        throw "アプリケーションに異常があります（想定外のウィンドウ: $childTitle）."
    }

    # --- アプリを閉じる ---
    $closeButton = Get-UIAButton -Parent $appWindow -Id $AID_close_button
    if ($null -eq $closeButton) { throw 'アプリケーションに異常があります（閉じるボタンが見つかりません）.' }
    Invoke-UIAElement $closeButton

} catch {
    Write-Error $_
    exit 1
}

exit 0
