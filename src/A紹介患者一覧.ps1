#
# 指定期間の当院からの紹介患者一覧(日付, ID, 氏名, 診療科, 紹介先名称)を取得する
#
# -Start 開始日日付(指定がなければ当日) -End 終了日日付(なければ当日)
Param (
    [string]$Start,
    [string]$End
)

. "$PSScriptRoot\Tools-AutomationCore.ps1"

# 引数の確認
if (-not $Start) {
    $Start = (Get-Date).ToString('yyyy/MM/dd')
}

if (-not $End) {
    $End = (Get-Date).ToString('yyyy/MM/dd')
}

if ((Get-Date $Start) -gt (Get-Date $End)) {
    throw "開始日 $Start が 終了日 $End よりも後に指定されています."
}

#
function Wait-AppWindowStateReady {
    Param (
        [Parameter(Mandatory)][Windows.Automation.AutomationElement]$appWindow
    )
    # アプリケーションウインドウが操作を受け付けるようになるのを30秒まで待つ
    $stopwatch = [System.Diagnostics.Stopwatch]::StartNew()
    $windowPattern = $appWindow.GetCurrentPattern([Windows.Automation.WindowPattern]::Pattern)
    while ($stopwatch.Elapsed.TotalSeconds -lt 30) {
        $windowState = $windowPattern.Current.WindowInteractionState
        if ($windowState -eq ([Windows.Automation.WindowInteractionState]::Running)) {
            break
        }
        [UIATools]::Sleep(1000)
    }
    # タイムアウトの確認
    if ($windowState -ne ([Windows.Automation.WindowInteractionState]::Running)) {
        throw '表示画面の更新確認に失敗しました.アプリケーションの状態を確認してください.'
    }
}

# オートメーション
$appWindow = GetAppWindow -Name '紹介患者一覧_AplinSyokai*' -ExecutablePath 'C:\ssi\exe\AplinSyokai.exe'
if ($null -eq $appWindow) {
    throw '紹介患者一覧を開始できませんでした.'
}

try {
    # 表示設定から条件を設定して表示する
    $controlPane = Get-UIAChildPane -Parent $appWindow -Id 'pnlSsfListContainer'
    $settingButton = Get-UIAButton -Parent $controlPane -Name '表示条件'
    Invoke-UIAElement $settingButton | Out-Null

    [UIATools]::Sleep(1500)
} catch {
    Write-Error $_
    throw 'ボタン操作に失敗しました.'
}

# 設定画面を取得
$setteiWindow = Get-UIAWindow -Parent $appWindow -Name '設定'
if ($null -eq $setteiWindow) {
    throw '表示設定への移行ができませんでした.'
}

try {
    # 開始日の設定
    $startPane = Get-UIAPane -Parent $setteiWindow -Id 'dtxtYmd_000'
    $startEdit = Get-UIAEdit -Parent $startPane -Id 'txtDate'
    Set-UIAValue -Element $startEdit -Value $Start -Force
    Set-UIAValue -Element $startEdit -Value '{Enter}' -Force -OmitEscape

    # 終了日の設定
    $endPane = Get-UIAPane -Parent $setteiWindow -Id 'dtxtYmd_001'
    $endEdit = Get-UIAEdit -Parent $endPane -Id 'txtDate'
    Set-UIAValue -Element $endEdit -Value $End -Force
    Set-UIAValue -Element $endEdit -Value '{Enter}' -Force -OmitEscape

    # 抽出条件の設定
    # 他施設から
    $incomingCheck = Get-UIACheckBox -Parent $setteiWindow -Id 'chkBunrui_000'
    $incomingCheckValue = Get-UIAValue $incomingCheck
    if (-not $incomingCheckValue) {
        Invoke-UIAElement $incomingCheck | Out-Null
    }
    
    # 他施設へ
    $outgoingCheck = Get-UIACheckBox -Parent $setteiWindow -Id 'chkBunrui_001'
    $outgoingCheckValue = Get-UIAValue $outgoingCheck
    if (-not $outgoingCheckValue) {
        Invoke-UIAElement $outgoingCheck | Out-Null
    }

    # 院内紹介
    $referrerCheck = Get-UIACheckBox -Parent $setteiWindow -Id 'chkBunrui_002'
    $referrerCheckValue = Get-UIAValue $referrerCheck
    if ($referrerCheckValue) {
        Invoke-UIAElement $referrerCheck | Out-Null
    }

    # 返書済み解除
    $responseStateRadio = Get-UIARadio -Parent $setteiWindow -Id 'OptHensyo_002'
    Invoke-UIAElement $responseStateRadio | Out-Null

    # 設定ボタンで検索
    $kakuteiButton = Get-UIAButton -Parent $setteiWindow -Id 'cmdButton_000'
    Invoke-UIAElement $kakuteiButton | Out-Null
} catch {
    Write-Error $_
    throw '設定に失敗しました.'
}

# データーの取得を待つ

Wait-AppWindowStateReady $appWindow

# CSVデーターを取得
$CSVfilename = ([System.IO.Path]::GetTempFileName() -replace '\..*$', '.csv')

try {
    # メニュー操作
    $menubar = Get-UIAChildMenuBar -Parent $appWindow -Id 'mnuMain'
    $exportMenuItem = Get-UIAMenuItem -Parent $menubar -Name 'CSV出力'
    Invoke-UIAElement $exportMenuItem | Out-Null

    [UIATools]::Sleep()

    # ファイルダイアログを指定して保存
    $filesaveDialog = Get-UIAWindow -Parent $appWindow -Name '名前を付けて保存'
    $filenameEdit = Get-UIAEdit -Parent $filesaveDialog -Id '1001'
    Set-UIAValue $filenameEdit $CSVfilename | Out-Null
    $filesaveButton = Get-UIAButton -Parent $filesaveDialog -Id '1'
    Invoke-UIAElement $filesaveButton | Out-Null

    [UIATools]::Sleep(500)

    # 確認ダイアログが出るので閉じる
    $confirmDialog = Get-UIAWindow -Parent $appWindow -Name 'SSIBaseFrm'
    $confirmButton = Get-UIAButton -Parent $confirmDialog -Name '了解(S)'
    Invoke-UIAElement $confirmButton | Out-Null

    [UIATools]::Sleep()
} catch {
    Write-Error $_
    throw 'CSV出力の操作に失敗しました.'
} finally {
    Remove-Item $CSVfilename -ErrorAction SilentlyContinue
}

# アプリケーションを閉じる
try {
    $exitMenuItem = Get-UIAMenuItem -Parent $menubar -Name '終了(X)'
    Invoke-UIAElement $exitMenuItem | Out-Null
} catch {
    Write-Error $_
}

# CSVファイルから情報を取得する
# 最初の2行は不要なのでスキップする
# CSVファイルは終わったら削除
try {
    $CSVfiledata = Get-Content -Path $CSVfilename -Encoding Default

    $CSVfiledata = $CSVfiledata | Select-Object -Skip 2
    $members = $CSVfiledata[0] -split ','

    $CSVfiledata = $CSVfiledata | Select-Object -Skip 1
    # ヘッダの重複と空白の確認
    $notuniqueHeaderMembers = ($members | Group-Object | Where-Object Count -gt 1).Name
    $i = 0
    while($i -lt $members.Count) {
        if ($members[$i] -in $notuniqueHeaderMembers) {
            $members[$i] = "$($members[$i])$i"
        }
        if ('' -eq $members[$i]) {
            $members[$i] = "$i"
        }
        $i++
    }
    $referralList = ConvertFrom-Csv -InputObject $CSVfiledata -Header @($members)| `
        Select-Object @( `
            @{Name='日付'; Expression={$_.'受付日'}}, `
            @{Name='ID'; Expression={$_.'ｶﾙﾃID'}}, `
            '患者氏名', `
            '診療科', `
            '入外', `
            @{Name='担当医'; Expression={$_.'当院担当医'}}, `
            @{Name='紹介先医療機関名'; Expression={$_.'紹介元医療機関名'}} `
            )
} catch {
    Write-Error $_
    throw 'CSVファイルに異常があります.'
} finally {
    Remove-Item $CSVfilename -ErrorAction SilentlyContinue
}

$referralList
