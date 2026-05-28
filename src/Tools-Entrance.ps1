# エントランスの操作をする
if (-not ([System.Management.Automation.PSTypeName]'UIATools').Type) {
    throw 'Tools-AutomationCode.ps1を先にロードしてください'
}

# エントランスが起動しているか確認
function Test-Entrance {
    # エントランスの起動確認(プロセスでチェック)
    return  (get-process -Name 'Entrance' -ErrorAction SilentlyContinue).Id
}

# 表示されている エントランス - 外来 からCSVで出力される情報を取得する
# returns [string]$_.Date - 表示されている日付, $_.List - CSVファイルの内容
function Get-GairaiList {
    # エントランスの起動確認(プロセスでチェック)
    $pidAplin = Test-Entrance

    if ($null -eq $pidAplin) {
        throw 'エントランスが起動していません'
    }

    # アプリケーションウインドウを取得
    # Name = 外来 だけだとexplorerなどノイズが多いのでカスタムする
    $appAplin = $null
    $conditions = @()
    $conditions += New-Object Windows.Automation.PropertyCondition([Windows.Automation.AutomationElement]::ControlTypeProperty, ([Windows.Automation.ControlType]::Window))
    $conditions += New-Object Windows.Automation.PropertyCondition([Windows.Automation.AutomationElement]::ProcessIdProperty, $pidAplin)
    # $conditions += New-Object Windows.Automation.PropertyCondition([Windows.Automation.AutomationElement]::AutomationIdProperty, 'OrdAplin')
    $conditions += New-Object Windows.Automation.PropertyCondition([Windows.Automation.AutomationElement]::NameProperty, '外来')
    try {
        $appAplin = Get-UIAControl -Parent ([UIATools]::RootElement) -Scope ([Windows.Automation.TreeScope]::Children) -Condition (New-Object Windows.Automation.AndCondition(,[Windows.Automation.PropertyCondition[]]$conditions)) -TimeoutSec 1
    } catch {}
    if ($null -eq $appAplin) {
        throw 'エントランスから外来を開いてください'
    }

    # エントランスの患者リストはネストが深いので探索に相当時間がかかる
    # Childernでpaneを列挙して必要な階層だけ検索する
    $paneMenu = $null
    $paneHeader = $null

    # level　1 - メニューのあるPaneを取得
    # $panes = Get-UIAControls -Parent $appAplin -Scope ([Windows.Automation.TreeScope]::Children) -Condition (New-Object Windows.Automation.PropertyCondition([Windows.Automation.AutomationElement]::ControlTypeProperty, ([Windows.Automation.ControlType]::Pane))) -TimeoutSec 1
    $panes = Get-UIAChildPanes -Parent $appAplin
    $paneMenu = $panes[2] # ($panes | Where-Object { $_.Current.AutomationId -eq 'pnlMainMenu' })[0]

    # level　2 to 3 - カレンダーを表示している部分の親Paneを取得
    # $panes = Get-UIAControls -Parent $panes[0] -Scope ([Windows.Automation.TreeScope]::Children) -Condition (New-Object Windows.Automation.PropertyCondition([Windows.Automation.AutomationElement]::ControlTypeProperty, ([Windows.Automation.ControlType]::Pane))) -TimeoutSec 1
    # $panes = Get-UIAControls -Parent $panes[0] -Scope ([Windows.Automation.TreeScope]::Children) -Condition (New-Object Windows.Automation.PropertyCondition([Windows.Automation.AutomationElement]::ControlTypeProperty, ([Windows.Automation.ControlType]::Pane))) -TimeoutSec 1
    $panes = Get-UIAChildPanes -Parent $panes[0]
    $panes = Get-UIAChildPanes -Parent $panes[0]
    $paneHeader = $panes[1] # ($panes | Where-Object { $_.Current.AutomationId -eq 'pnlHeader2' })[0]

    if ($null -eq $paneMenu -or $null -eq $paneHeader) {
        throw 'エントランスの構造が想定外です'
    }

    # 表示している日付を取得
    $editAplinDate = Get-UIAEdit -Parent $paneHeader -Id 'txtDate'
    if ($null -eq $editAplinDate) {
        throw 'エントランスの構造が想定外です'
    }
    [string]$displayedDateString = Get-UIAValue $editAplinDate

    if ($displayedDateString -eq '月') {
        $displayedDateString = ''
    }

    # メニュー操作でCSV保存して一覧に表示されている情報を取得
    $menuCSV = Get-UIAMenuItem -Parent $paneMenu -Name 'CSV出力...'
    if ($null -eq $menuCSV) {
        throw 'エントランスのメニュー操作ができません'
    }
    Invoke-UIAElement $menuCSV | Out-Null
    [UIATools]::Sleep(300)

    #ファイル保存ダイアログで一時ファイル名でCSV保存
    $CSVfilename = ([System.IO.Path]::GetTempFileName() -replace '\..*$', '.csv')

    $windowFileSave = Get-UIAWindow -Parent $appAplin -Name '名前を付けて保存'
    
    Set-UIAValue -Element (Get-UIAComboBox -Parent $windowFileSave -Id 'FileNameControlHost') $CSVfilename
    Invoke-UIAElement (Get-UIAButton -Parent $windowFileSave -id '1') | Out-Null

    [UIATools]::Sleep(300)

    #CSVファイルの内容を取得
    $gairaiListData = (Get-Content -Path $CSVfilename -Encoding Default | ConvertFrom-Csv)

    #CSVファイルを削除
    Remove-Item -Path $CSVfilename -ErrorAction SilentlyContinue | Out-Null

    # 返り値はカスタムオブジェクト
    return @{
        Date = $displayedDateString
        List = $gairaiListData
    }
}

# . でライブラリとして利用していない場合のデフォルトアクション
if ($MyInvocation.InvocationName -ne '.') {
    if (Test-Entrance) {
        'エントランスは起動しています' | Out-Default
    } else {
        'エントランスは起動していません' | Out-Default
    }
}
