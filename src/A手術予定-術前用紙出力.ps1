#
# 手術予定表からデーターを抽出して術前訪問用紙を出力
#
# アセンブリをロード
. "$PSScriptRoot\Tools-AutomationCore.ps1"
. "$PSScriptRoot\Tools-ExcelWindow.ps1"
# . "$PSScriptRoot\Tools-UserInteraction.ps1"

# 定数定義
$appName = '手術予定入力*'
$appPath = 'C:\SSI\EXE\srvSyujutuL.exe'
$excelTimeoutSec = 60
$templatePath = "$PSScriptRoot\template-術前訪問用紙.xlsx"
# 変数
$appWindow = $null
$excelCOM = $null
$excelWorkbookCOM = $null
$excelSheetCOM = $null

# automation
try {
    # 値のチェック
    if (-not $appName -or -not (Test-Path -Path $templatePath)) {
        throw '設定を確認してください'
    }

    # 手術予定の実行確認
    $appProcess = (Get-Process -Name (($appPath -split '\\')[-1] -split '\.')[0] -ErrorAction SilentlyContinue)
    if ($null -eq $appProcess -or $appProcess.Count -eq 0) {
        throw '手術予定一覧をひらいて適当な日付を選択してください'
    }

    $appWindow = GetAppWindow -Name $appName

    # 手術予定プレビューの表示
    # 事前にエクセルのプロセスを取得しておく
    $prevExcelPIDs = (Get-Process -Name excel -ErrorAction SilentlyContinue).Id

    # メニューから ファイル - プレビューをinvoke
    Set-UIAWindowActive $appWindow
    $menuItemFile = Get-UIAMenuItem -Parent $appWindow -Name 'ﾌｧｲﾙ(F)'
    if ($null -eq $menuItemFile) {
        throw '表示に問題があります'
    }
    Set-UIAElementExpanded $menuItemFile
    [UIATools]::Sleep()

    $menuItemPreview= Get-UIAMenuItem -Parent $appWindow -Name '予約一覧プレビュー'
    if ($null -eq $menuItemPreview) {
        throw '表示に問題があります'
    }
    Invoke-UIAElement $menuItemPreview
    [UIATools]::Sleep()

    # プレビュー確認ウインドウが開くがこれはautomationできないの
    # アプリケーションウインドウがアクティブになっているのでSendKeysで ALT + P を送る
    [System.Windows.Forms.SendKeys]::SendWait('%P')

    # 症例なしダイアログが出ていないか確認 - 出ていたら終了
    $alertWindows = Get-UIAChildWindows -Parent $appWindow
    if ($null -ne $alertWindows -and $alertWindows.Count -gt 0) {
        $errorText = '手術予定入力のエラーです'

        # ダイアログのメッセージを取得
        $errorTextElement = Get-UIAText $alertWindows[0]
        if ($null -ne $errorTextElement) {
            $errorText = Get-UIAValue $errorTextElement
        }
        # ボタン操作で閉じる
        $closeBtn = Get-UIAButton -Parent $alertWindows[0] -Id '2'
        if ($null -ne $closeBtn) {
            Invoke-UIAElement $closeBtn
        }
        throw $errorText
    }

    # プレビューExcelの起動待ち
    $newExcelhWnd = $null
    $stopwatch = [System.Diagnostics.Stopwatch]::StartNew()
    while ($stopwatch.Elapsed.TotalSeconds -lt $excelTimeoutSec) {
        $newExcelProcesses = (Get-Process -Name excel -ErrorAction SilentlyContinue | Where-Object { $_.Id -notin $excelPIDs -and $_.MainWindowHandle -ne 0})
        if ($newExcelProcesses.Count -ge 1) {
            $newExcelhWnd = $newExcelProcesses.MainWindowHandle
            break;
        }
        [UIATools]::Sleep(1000)
    }
    if ($null -eq $newExcelhWnd) {
        throw 'Excelの起動待ちがタイムアウトしました'
    }

    # ExceのCOMオブジェクトを接続
    $excelCOM = [ExcelWindowFactory]::GetExcelApplicationFromHwnd($newExcelHwnd)
    [UIATools]::Sleep()

    # シート3 - データーに接続
    $excelWorkbookCOM = $excelCOM.ActiveWorkbook
    if ($null -eq $excelWorkbookCOM) {
        throw '既存エクセルワークブックに接続できません'
    }
    $excelSheetCOM = $excelWorkbookCOM.Sheets.Item(3)

    # データーを読み出す
    $dataDate = $excelSheetCOM.Range('CL2').Value2 # 日付シリアル 文字列としてP1

    $data = @()
    $row = 2
    while ($excelSheetCOM.Range("A$row").Value2) {
        $data += (, @{
            id = $excelSheetCOM.Range("D$row").Text # B2
            name = $excelSheetCOM.Range("E$row").Text # F2
            age = $excelSheetCOM.Range("G$row").Text # J2
            sex = $excelSheetCOM.Range("H$row").Text # N2
            bloodtype = $excelSheetCOM.Range("BU$row").Text # Q2
            height = $excelSheetCOM.Range("BS$row").Text # A3
            weight = $excelSheetCOM.Range("BT$row").Text # F3
            ka = $excelSheetCOM.Range("B$row").Text # N3
            room = $excelSheetCOM.Range("AO$row").Text # K3
            diagnosis = $excelSheetCOM.Range("I$row").Text -replace ' / ',"`n" # C5
            procedures = $excelSheetCOM.Range("K$row").Text -replace ' / ',"`n" # L5
            ishi = ($excelSheetCOM.Range("M$row").Text -split '\s+')[0] # Q3
            kinkyu = $excelSheetCOM.Range("CC$row").Text -eq '●' # チェック2
        })
        $row++
    }

    # プレビューのexcelを閉じる
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excelSheetCOM) | Out-Null
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excelWorkbookCOM) | Out-Null
    $excelCOM.ActiveWorkbook.Close($false)
    $excelCOM.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excelCOM) | Out-Null
    [UIATools]::Sleep(1000)

    # 帳票出力用のエクセルを開く
    $excelCOM = New-Object -ComObject Excel.Application
    $excelWorkbookCOM = $excelCOM.Workbooks.Open($templatePath)
    if ($null -eq $excelWorkbookCOM) {
        throw 'テンプレートエクセルワークブックに接続できません'
    }
    $excelSheetCOM = $excelWorkbookCOM.ActiveSheet

    # 日付を入力 yyyy / mm / dd (weekday) 
    $datestr = ''
    if ($dataDate.GetType() -ne 'String') {
        $datestr = [datetime]::FromOADate($dataDate).ToString('yyyy / M / d (ddd)')
    } else {
        try {
            $datestr = Get-Date $dataDate -Format 'yyyy / M / d (ddd)'
        } catch{
            $datestr = $dataDate
        }
    }
    Write-Host $datestr
    $excelSheetCOM.Range('P1').Value = $datestr
    $excelSheetCOM.Range('O4').Value = $datestr.Substring(0,4) + '　/　 　/ 　 （　）'

    # データー毎に値を入れてゆく
    foreach ($case in $data)  {
        $case | ConvertTo-Json -Depth 2 | Write-Host
        $excelSheetCOM.Range('B2').Value = $case.id 
        $excelSheetCOM.Range('F2').Value = $case.name
        $excelSheetCOM.Range('J2').Value = $case.age
        $excelSheetCOM.Range('N2').Value = $case.sex
        if ($case.bloodtype -notlike '*`?*') {
            $excelSheetCOM.Range('Q2').Value = $case.bloodtype
        } else {
            $excelSheetCOM.Range('Q2').Value = ''
        }
        if ($case.height) {
            $excelSheetCOM.Range('A3').Value = $case.height.toString()
        } else {
            $excelSheetCOM.Range('A3').Value = ''
        }
        if ($case.weight) {
            $excelSheetCOM.Range('F3').Value = $case.weight.toString()
        } else {
            $excelSheetCOM.Range('F3').Value = ''
        }
        if ($case.height -and $case.weight) {
            $excelSheetCOM.Range('I3').Value = (10000.0 * ([Double]$case.weight) / ([Double]$case.height) / ([Double]$case.height)).ToString('0.0')
        }
        $excelSheetCOM.Range('N3').Value = $case.ka
        $excelSheetCOM.Range('C5').Value = $case.diagnosis
        $excelSheetCOM.Range('L5').Value = $case.procedures
        $excelSheetCOM.Range('Q3').Value = $case.ishi
        # 部屋番号があるときは括弧をつける
        if ($case.room) {
            $excelSheetCOM.Range('K3').Value = "$case.room (　)"
        } else {
            $excelSheetCOM.Range('K3').Value = ''
        }

        # 緊急チェックを選択
        if ($case.kinkyu) {
            $excelSheetCOM.CheckBoxes('チェック 1').Value = 0
            $excelSheetCOM.CheckBoxes('チェック 2').Value = 1
        } else {
            $excelSheetCOM.CheckBoxes('チェック 1').Value = 1
            $excelSheetCOM.CheckBoxes('チェック 2').Value = 0
        }
        # 1ページ目だけを印刷
        $excelSheetCOM.PrintOut(1, 1)
        [UIATools]::Sleep(300)
    }

    # COMを開放
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excelSheetCOM) | Out-Null
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excelWorkbookCOM) | Out-Null
    # excelを保存せずに終了
    $excelCOM.ActiveWorkbook.Close($false)
    $excelCOM.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excelCOM) | Out-Null
}

catch {
    throw $_
}
finally {
    try {
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excelSheetCOM) | Out-Null
    } catch {}
    try {
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excelWorkbookCOM) | Out-Null
    } catch {}
    try {
        $excelCOM.ActiveWorkbook.Close($false)
    } catch {}
    try {
        $excelCOM.Quit()
    } catch {}
    try {
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excelCOM) | Out-Null
    } catch {}
}
