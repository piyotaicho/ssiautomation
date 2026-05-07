# 来週の産婦人科の手術予定を印刷する
#

# 設定
$no_entry = $false
$highlight = $false

# 定数
$dllpath = 'Y:\病院\医局\産婦人科\やまもと\解析\UIDeskAutomation.dll'
$exe = "C:\SSI\exe\SjtYtPrtEx.exe"
$procName = "SjtYtPrtEx"
$NAME_main_window = "手術室予定表 (SjtYtPrtEx)"

$ORD_date_start_pane = 1
$AID_date_start_edit = "2"
$ORD_date_end_pane = 0
$AID_date_end_edit = "2"
$AID_ka_string_list = "20"
$AID_print_button = "22"
$AID_close_button = "21"

$NAME_print_window = "手術室予定表"
$AID_start_print_button = "4"

$NAME_cofirm_window = "確認！"
$AID_confirm_ok_button = "2"

# 初期設定
Add-Type -Path $dllpath
$engine = [UIDeskAutomationLib.Engine]::new()
$engine.Timeout = 1000

Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName Microsoft.VisualBasic

$cultureInfo = New-Object cultureinfo('ja-jp',$true)
$cultureInfo.DateTimeFormat.Calendar = New-Object System.Globalization.JapaneseCalendar

# アプリの起動確認と確認
try {
    $procId = (Get-Process -Name $procName -ErrorAction SilentlyContinue).Id
} catch {
}
if ( $procId -eq $null ) {
    $procId = $engine.StartProcess( $exe )
    $engine.Sleep(3000)
}

$appWindow = $engine.GetTopLevelByProcId( $procId, $NAME_main_window )

# 値の計算
$today = Get-Date
$nextmonday = $today.AddDays(8-[int]$today.DayOfWeek)

$start_date_string = $nextmonday.ToString('yyMMdd', $cultureInfo)
$end_date_string = ($nextmonday.AddDays(5)).ToString('yyMMdd', $cultureInfo)

# 値の設定
$panes = $appWindow.Panes("")
if ( $panes.Count -ne 2 ) {
    write-host "アプリケーションに異常があります."
    exit 1
}

$wiget = $panes[$ORD_date_start_pane].FindFirstDescendant([UIDeskAutomationLib.UIDA_Property]::AutomationId, $AID_date_start_edit)
if ( $null -eq $wiget ) {
    write-host "アプリケーションに異常があります."
    exit 1
}
$highlight -and $wiget.Highlight()
$wiget.AsEdit().SetText($start_date_string)

$wiget = $panes[$ORD_date_end_pane].FindFirstDescendant([UIDeskAutomationLib.UIDA_Property]::AutomationId, $AID_date_end_edit)
if ( $null -eq $wiget ) {
    write-host "アプリケーションに異常があります."
    exit 1
}
$highlight -and $wiget.Highlight()
$wiget.AsEdit().SetText($end_date_string)

$wiget = $appWindow.FindFirstDescendant([UIDeskAutomationLib.UIDA_Property]::AutomationId, $AID_ka_string_list)
if ( $null -eq $wiget ) {
    write-host "アプリケーションに異常があります."
    exit 1
}
$highlight -and $wiget.Highlight()

$wiget = $wiget.AsList().ListItem("23 産婦人科")
if ( $null -eq $wiget ) {
    write-host "アプリケーションに異常があります."
    exit 1
} else {
    $highlight -and $wiget.Highlight()
    $wiget.Select()
}

$wiget = $appWindow.FindFirstDescendant([UIDeskAutomationLib.UIDA_Property]::AutomationId, $AID_print_button)
if ( $null -eq $wiget ) {
    write-host "アプリケーションに異常があります."
    exit 1
}
$highlight -and $wiget.Highlight()
$wiget.AsButton().Press()

# 印刷ウインドウや確認ウインドでの操作
$engine.Sleep(1000)

$targetWindow = $appWindow.WindowAt("*", 0)
$childWindowTitle = $targetWindow.GetWindowText()
if ( $childWindowTitle -eq $NAME_print_window) {
    # 症例があれば印刷ダイアログが開く
    $wiget = $targetWindow.FindFirstDescendant([UIDeskAutomationLib.UIDA_Property]::AutomationId, $AID_start_print_button)
    if ( $null -eq $wiget ) {
        write-host "アプリケーションに異常があります."
        exit 1
    }
    $highlight -and $wiget.Highlight()
    $wiget.AsButton().Press()
    $engine.Sleep(5000) # 処理中のダイアログが表示されているので少し長めにウエイトを取る
} else {
    # 症例がないと確認ダイアログが開く
    if ( $childWindowTitle -eq $NAME_cofirm_window ) {
        $no_entry = $true
        $wiget = $targetWindow.FindFirstDescendant([UIDeskAutomationLib.UIDA_Property]::AutomationId, $AID_confirm_ok_button)
        if ( $null -eq $wiget ) {
            write-host "アプリケーションに異常があります."
            exit 1
        }
        $wiget.AsButton().Press()
        $engine.Sleep(1000)
    } else {
        write-host "アプリケーションに異常があります."
        exit 1
    }
}

if ( $no_entry -eq $true ) {
    [System.Windows.MessageBox]::Show("来週の手術予定がありませんでした.", "報告", "OK")
}

# アプリケーションを閉じる
$wiget = $appWindow.FindFirstDescendant([UIDeskAutomationLib.UIDA_Property]::AutomationId, $AID_close_button)
if ( $null -eq $wiget ) {
    write-host "アプリケーションに異常があります."
    exit 1
}
$wiget.AsButton().Press()

exit 0
