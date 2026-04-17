# アセンブリのロード
Add-Type -AssemblyName "UIAutomationClient"
Add-Type -AssemblyName "UIAutomationTypes"

# --- [0] 版用ユーティリティ関数 ---
function Start-UIASleep {
    <#
    .SYNOPSIS
        UI操作の合間に待機を挟みます。デフォルトは 200ms です。
    #>
    param(
        [Parameter(ParameterSetName = "Milli", Position = 0)]
        [int]$Milliseconds,

        [Parameter(ParameterSetName = "Sec")]
        [int]$Seconds
    )

    if ($PSBoundParameters.ContainsKey('Seconds')) {
        # 秒が指定された場合
        Start-Sleep -Seconds $Seconds
    }
    elseif ($PSBoundParameters.ContainsKey('Milliseconds')) {
        # ミリ秒が指定された場合
        Start-Sleep -Milliseconds $Milliseconds
    }
    else {
        # 何も指定がない場合はデフォルトの 200ms
        Start-Sleep -Milliseconds 200
    }
}

# --- [1] 汎用検索エンジン ---
function Get-UIAControls {
    <#
    .SYNOPSIS
        条件に一致するすべてのUI要素を配列で取得します。Nameにワイルドカード(*)を使用可能です。
    #>
    param(
        [Windows.Automation.AutomationElement]$Parent,
        [string]$Id,
        [string]$Name,
        [Windows.Automation.ControlType]$Type,
        [Windows.Automation.TreeScope]$Scope = ([Windows.Automation.TreeScope]::Descendants),
        [Windows.Automation.Condition]$Condition # 直接Conditionを渡す場合に使用
    )

    Write-Host 'Get-UIAControls:'

    if ($null -eq $Parent) {
        throw "Parent要素は必須です"
    }

    $hasWildcardName = $Name.Contains('*')
    
    $searchCond = $Condition
    if ($null -eq $searchCond) {
        write-host "Building search condition from parameters: Id='$Id', Name='$Name', Type='$($Type.LocalizedControlType)'"
        $conditions = @()
        if ($Id) { $conditions += New-Object Windows.Automation.PropertyCondition([Windows.Automation.AutomationElement]::AutomationIdProperty, $Id) }
        if ($Name -and -not $hasWildcardName) { write-host 'Get-UIAControls: Adding non wildcard Name condition'; $conditions += New-Object Windows.Automation.PropertyCondition([Windows.Automation.AutomationElement]::NameProperty, $Name) }
        if ($Type) { $conditions += New-Object Windows.Automation.PropertyCondition([Windows.Automation.AutomationElement]::ControlTypeProperty, $Type) }

        $searchCond = if ($conditions.Count -eq 0) { [Windows.Automation.Condition]::TrueCondition }
                     elseif ($conditions.Count -gt 1) { New-Object Windows.Automation.AndCondition(,[Windows.Automation.PropertyCondition[]]$conditions) }
                     else { $conditions[0] }
    }

    $elements = $Parent.FindAll($Scope, $searchCond)

    write-host "Found $($elements.Count) elements matching initial conditions"

    # ワイルドカードのName指定がある場合は、Nameプロパティをフィルタリングする
    if ($hasWildcardName) {
        $elements = $elements | Where-Object {
            write-host "Checking element: Name='$($_.Current.Name)', Id='$($_.Current.AutomationId)' against pattern '$Name'"
            $_.Current.Name -like $Name
        }
    }
    return @($elements)
}

function Get-UIAControl {
    param(
        [Windows.Automation.AutomationElement]$Parent, 
        [string]$Id, 
        [string]$Name, 
        [Windows.Automation.ControlType]$Type, 
        [Windows.Automation.TreeScope]$Scope = ([Windows.Automation.TreeScope]::Descendants),
        [Windows.Automation.Condition]$Condition,
        [int]$TimeoutSec = 10
    )
    Write-Host 'Get-UIAControl:'

    if ($null -eq $Parent) {
        throw "Parent要素は必須です"
    }

    # ワイルドカードがある場合は、Get-UIAControlsの結果から最初の要素を取得する。なければFindFirstで高速に取得する。
    $hasWildcardName = ($null -ne $Name) -and $Name.Contains('*')
    if ($hasWildcardName) {
        return Get-UIAControls -Parent $Parent -Id $Id -Name $Name -Type $Type -Condition $Condition -Scope $Scope | Select-Object -First 1
    }

    # Conditionが未指定の場合のみ、Id, NameとTypeから組み立てる
    $searchCond = $Condition
    if ($null -eq $searchCond) {
        write-host "Building search condition from parameters: Id='$Id', Name='$Name', Type='$($Type.LocalizedControlType)'"
        $conditions = @()
        if ($Id) { $conditions += New-Object Windows.Automation.PropertyCondition([Windows.Automation.AutomationElement]::AutomationIdProperty, $Id) }
        if ($Name) { $conditions += New-Object Windows.Automation.PropertyCondition([Windows.Automation.AutomationElement]::NameProperty, $Name) }
        if ($Type) { $conditions += New-Object Windows.Automation.PropertyCondition([Windows.Automation.AutomationElement]::ControlTypeProperty, $Type) }
        
        if ($conditions.Count -gt 1) {
            # AndCondition は配列を 1 引数で渡す必要がある（先頭のカンマで配列を保持）
            $searchCond = New-Object Windows.Automation.AndCondition(,[Windows.Automation.PropertyCondition[]]$conditions)
        } elseif ($conditions.Count -eq 1) {
            $searchCond = $conditions[0]
        } else {
            $searchCond = [Windows.Automation.Condition]::TrueCondition
        }
    }

    # debug findall
    $allelements = $Parent.FindAll($Scope, [Windows.Automation.Condition]::TrueCondition)
    write-host "Dump all elements under parent: $($Parent.Current.Name)"
    foreach ($el in $allelements) {
        write-host "  Element: Name='$($el.Current.Name)', Id='$($el.Current.AutomationId)', Type='$($el.Current.ControlType.LocalizedControlType)'"
    }
    write-host "--"

    $stopwatch = [System.Diagnostics.Stopwatch]::StartNew()
    while ($stopwatch.Elapsed.TotalSeconds -lt $TimeoutSec) {
        $element = $Parent.FindFirst($Scope, $searchCond)
        if ($null -ne $element) { return $element }
        Start-Sleep -Milliseconds 500
    }
    throw "要素取得タイムアウト: Name=$Name, Id=$Id"
}

# --- [2] コントロール別 取得関数 ---

# デスクトップ直下からアプリウィンドウを探す（名前またはPID指定）
function Get-UIAAppWindow { 
    param(
        [string]$Name,
        [int]$ProcessId = 0
    ) 

    $aeType = [Windows.Automation.AutomationElement]
    $root = $aeType::RootElement

    # 1. 確実に判定できる「型(Window)」と「PID」だけで先に検索条件を作る
    $conditions = @()
    $conditions += New-Object Windows.Automation.PropertyCondition($aeType::ControlTypeProperty, [Windows.Automation.ControlType]::Window)

    if ($ProcessId -gt 0) { 
        $conditions += New-Object Windows.Automation.PropertyCondition($aeType::ProcessIdProperty, $ProcessId) 
    }
 
    # Nameにワイルドカードが含まれている場合は、PropertyConditionには入れず、後段のフィルタリング（Get-UIAControl内の処理）に任せる
    $hasWildcardName = $Name.Contains('*')
     if ($null -ne $Name -and -not $hasWildcardName) { 
        $conditions += New-Object Windows.Automation.PropertyCondition($aeType::NameProperty, $Name) 
    }

    # PIDもNameも指定がない場合は、$conditions[0]（Window条件）だけで検索する
    if ($conditions.Count -gt 1) {
        $finalCond = New-Object Windows.Automation.AndCondition([Windows.Automation.Condition[]]$conditions)
    } else {
        $finalCond = $conditions[0]
    }

    # 2. Get-UIAControls を呼び出す。
    return Get-UIAControls -Parent $root -Name $Name -Condition $finalCond -Scope ([Windows.Automation.TreeScope]::Children) | Select-Object -First 1
}

# 任意の親要素（Windowなど）の下にある子要素を探す
function Get-UIAWindow   { param($Parent, $Id, $Name) Get-UIAControl -Parent $Parent -Id $Id -Name $Name -Type ([Windows.Automation.ControlType]::Window) }
function Get-UIAWindows  { param($Parent, $Id, $Name) Get-UIAControls -Parent $Parent -Id $Id -Name $Name -Type ([Windows.Automation.ControlType]::Window) }

function Get-UIAPane     { param($Parent, $Id, $Name) Get-UIAControl -Parent $Parent -Id $Id -Name $Name -Type ([Windows.Automation.ControlType]::Pane) }
function Get-UIAPanes    { param($Parent, $Id, $Name) Get-UIAControls -Parent $Parent -Id $Id -Name $Name -Type ([Windows.Automation.ControlType]::Pane) }

function Get-UIAButton   { param($Parent, $Id, $Name) Get-UIAControl -Parent $Parent -Id $Id -Name $Name -Type ([Windows.Automation.ControlType]::Button) }
function Get-UIAText     { param($Parent, $Id, $Name) Get-UIAControl -Parent $Parent -Id $Id -Name $Name -Type ([Windows.Automation.ControlType]::Text) }
function Get-UIAEdit     { param($Parent, $Id, $Name) Get-UIAControl -Parent $Parent -Id $Id -Name $Name -Type ([Windows.Automation.ControlType]::Edit) }
function Get-UIACheckBox { param($Parent, $Id)        Get-UIAControl -Parent $Parent -Id $Id -Type ([Windows.Automation.ControlType]::CheckBox) }
function Get-UIAComboBox { param($Parent, $Id)        Get-UIAControl -Parent $Parent -Id $Id -Type ([Windows.Automation.ControlType]::ComboBox) }
function Get-UIARadio    { param($Parent, $Id)        Get-UIAControl -Parent $Parent -Id $Id -Type ([Windows.Automation.ControlType]::RadioButton) }
function Get-UIAMenuBar  { param($Parent, $Name)      Get-UIAControl -Parent $Parent -Name $Name -Type ([Windows.Automation.ControlType]::MenuBar) }
function Get-UIAMenuItem { param($Parent, $Name)      Get-UIAControl -Parent $Parent -Name $Name -Type ([Windows.Automation.ControlType]::MenuItem) }

# --- [3] 操作・ユーティリティ関数 ---

function Invoke-UIAElement {
    param([Parameter(Mandatory)] $Element)
    if ($Element.TryGetCurrentPattern([Windows.Automation.InvokePattern]::Pattern, [ref]$p)) {
        $p.Invoke()
    } elseif ($Element.TryGetCurrentPattern([Windows.Automation.TogglePattern]::Pattern, [ref]$p)) {
        $p.Toggle()
    } else { throw "Invoke非対応の要素です" }
}

function Set-UIAElementExpanded {
    param([Parameter(Mandatory)] $Element)
    if ($Element.TryGetCurrentPattern([Windows.Automation.ExpandCollapsePattern]::Pattern, [ref]$p)) {
        $p.Expand()
    } else {
        throw "展開非対応の要素です"
    }
}

function Set-UIAElementCollapsed {
    param([Parameter(Mandatory)] $Element)
    if ($Element.TryGetCurrentPattern([Windows.Automation.ExpandCollapsePattern]::Pattern, [ref]$p)) {
        $p.Collapse()
    } else {
        throw "折りたたみ非対応の要素です"
    }
}

function Close-UIAWindow {
    param([Parameter(Mandatory)] $Element)
    if ($Element.Current.ControlType -ne [Windows.Automation.ControlType]::Window) {
        throw "ウィンドウ以外の要素を閉じることはできません"
    }

    if ($Element.TryGetCurrentPattern([Windows.Automation.WindowPattern]::Pattern, [ref]$p)) {
        $p.Close()
    } else {
        throw "ウィンドウクローズ非対応の要素です"
    }
}

function Get-UIAName {
    <#
    .SYNOPSIS
        要素の Name プロパティ（Windowのタイトル、Buttonのテキスト、Textラベルなど）を取得します。
    #>
    param(
        [Parameter(Mandatory)]
        [Windows.Automation.AutomationElement]$Element
    )

    return $Element.Current.Name
}

function Set-UIAValue {
    param([Parameter(Mandatory)] $Element, [Parameter(Mandatory)] [string]$Value)
    if ($Element.TryGetCurrentPattern([Windows.Automation.ValuePattern]::Pattern, [ref]$p)) {
        $p.SetValue($Value)
    } else {
        throw "値入力非対応の要素です"
    }
}

function Get-UIAValue {
    <#
    .SYNOPSIS
        Text, Edit, ListItem などの要素から表示されている文字列を取得します。
    #>
    param(
        [Parameter(Mandatory)]
        [Windows.Automation.AutomationElement]$Element
    )

    # 1. ValuePattern を持っている場合 (Edit, ComboBoxなど)
    if ($Element.TryGetCurrentPattern([Windows.Automation.ValuePattern]::Pattern, [ref]$vp)) {
        return $vp.Current.Value
    }
    
    # 2. ValuePattern がない場合 (Text, Button, ListItemなど)
    # 多くの場合は Name プロパティに文字列が入っている
    return $Element.Current.Name
}

function Get-UIAComboBoxItems {
    <#
    .SYNOPSIS
        コンボボックスを展開し、内部のアイテム一覧を String[] として取得します。
    #>
    param([Parameter(Mandatory)] [Windows.Automation.AutomationElement]$ComboBox)
    $expandPattern = $true
    try {
        Set-UIAElementExpanded -Element $ComboBox
        Start-UIASleep -Milliseconds 200 # 展開待ち
    } catch {
        $expandPattern = $false
    }

    $listItems = Get-UIAControls -Parent $ComboBox -Type ([Windows.Automation.ControlType]::ListItem) -Scope Descendants
    $itemNames = foreach ($item in $listItems) { $item.Current.Name }

    if ($expandPattern) {
        Set-UIAElementCollapsed -Element $ComboBox
    }
    return [string[]]$itemNames
}

# --- [4] UIオートメーション複合関数 ---
function GetAppWindow {
    param(
        [Parameter(Mandatory)] [string]$Name,
        [string]$ExecutablePath
    )

    $window = $null
    # 既に起動していたらウィンドウを取得して返す
    try {
        $window = Get-UIAAppWindow -Name $Name
    } catch {}

    if ($null -ne $window) {
        return $window
    }

    # 見つからない場合は新規起動する
    if ($null -ne $ExecutablePath -and (Test-Path $ExecutablePath)) {
        write-host "アプリを起動します: $ExecutablePath"
        $process = Start-Process -FilePath $ExecutablePath
        Start-UIASleep -Sec 5 # 起動待ち
        try {
            $window = Get-UIAAppWindow -Name $Name -ProcessId $process.Id
        } catch {}
    }
    return $window
}

# --- [5] テストコード --- TeraTermを起動して接続ダイアログを閉じて、ヘルプのAbout TeraTermを開いて、閉じてExitする ---
$exePath = 'C:\Program Files\VideoLAN\VLC\vlc.exe'
$name_app_window = 'VLC*'

try {
    [Windows.Automation.AutomationElement]$appWindow = GetAppWindow -Name $name_app_window -ExecutablePath $exePath
    if ($null -eq $appWindow) {
        throw "アプリウィンドウが見つかりませんでした: Name=$name_app_window"
    }

    # ヘルプ > Tera Termについて を開く
    $menuBar = Get-UIAMenuBar -Parent $appWindow
    if ($null -eq $menuBar) {
        throw "メニューバーが見つかりませんでした"
    }
    $menuHelp = Get-UIAMenuItem -Parent $menuBar -Name 'ヘルプ*'
    if ($null -eq $menuHelp) {
        throw "ヘルプメニューが見つかりませんでした"
    }
    Expand-UIAElement -Element $menuHelp

    $menuAbout = Get-UIAMenuItem -Parent $menuHelp -Name 'Video*'
    if ($null -eq $menuAbout) {
        throw "Aboutメニューが見つかりませんでした"
    }
    Invoke-UIAElement -Element $menuAbout
    Start-UIASleep -Seconds 5

    # Aboutダイアログを閉じる
    $aboutDialog = Get-UIAWindow -Parent $appWindow -Name 'VideoLAN*'
    if ($null -ne $aboutDialog) {
        Start-UIASleep -Seconds 2
        Close-UIAWindow -Element $aboutDialog
    }

    # アプリを終了する
    Close-UIAWindow -Element $appWindow
} catch {
    Write-Error $_
}
