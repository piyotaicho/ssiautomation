# アセンブリのロード
Add-Type -AssemblyName "UIAutomationClient"
Add-Type -AssemblyName "UIAutomationTypes"

# --- [0] 版用ユーティリティ関数 ---
function Start-UiaSleep {
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
function Get-UiaControls {
    <#
    .SYNOPSIS
        条件に一致するすべてのUI要素を配列で取得します。Nameにワイルドカード(*)を使用可能です。
    #>
    param(
        [Windows.Automation.AutomationElement]$Parent = [Windows.Automation.AutomationElement]::RootElement,
        [string]$Id,
        [string]$Name,
        [Windows.Automation.ControlType]$Type,
        [Windows.Automation.TreeScope]$Scope = [Windows.Automation.TreeScope]::Descendants,
        [Windows.Automation.Condition]$Condition # 直接Conditionを渡す場合に使用
    )
    
    $finalCond = $Condition
    if ($null -eq $finalCond) {
        $conditions = @()
        if ($Id) { $conditions += New-Object Windows.Automation.PropertyCondition([Windows.Automation.AutomationElement]::AutomationIdProperty, $Id) }
        if ($Type) { $conditions += New-Object Windows.Automation.PropertyCondition([Windows.Automation.AutomationElement]::ControlTypeProperty, $Type) }

        $finalCond = if ($conditions.Count -eq 0) { [Windows.Automation.Condition]::TrueCondition }
                     elseif ($conditions.Count -gt 1) { New-Object Windows.Automation.AndCondition(,[Windows.Automation.PropertyCondition[]]$conditions) }
                     else { $conditions[0] }
    }

    $elements = $Parent.FindAll($Scope, $finalCond)
    
    # Name プロパティに対してワイルドカード（-like）でフィルタリング
    if ($null -ne $Name -and $Name -ne "") {
        $elements = $elements | Where-Object { $_.Current.Name -like $Name }
    }
    return @($elements)
}

function Get-UiaControl {
    param(
        $Parent, 
        [string]$Id, 
        [string]$Name, 
        [Windows.Automation.ControlType]$Type, 
        [Windows.Automation.Condition]$Condition,
        [Windows.Automation.TreeScope]$Scope = [Windows.Automation.TreeScope]::Descendants,
        [int]$TimeoutSec = 10
    )
    
    $stopwatch = [System.Diagnostics.Stopwatch]::StartNew()
    
    # Conditionが未指定の場合のみ、IdとTypeから組み立てる
    $searchCond = $Condition
    if ($null -eq $searchCond) {
        $conds = @()
        if ($Id) { $conds += New-Object Windows.Automation.PropertyCondition([Windows.Automation.AutomationElement]::AutomationIdProperty, $Id) }
        if ($Type) { $conds += New-Object Windows.Automation.PropertyCondition([Windows.Automation.AutomationElement]::ControlTypeProperty, $Type) }
        
        if ($conds.Count -gt 1) { 
            $searchCond = New-Object Windows.Automation.AndCondition(,[Windows.Automation.Condition[]]$conds) 
        } elseif ($conds.Count -eq 1) {
            $searchCond = $conds[0]
        } else {
            $searchCond = [Windows.Automation.Condition]::TrueCondition
        }
    }

    while ($stopwatch.Elapsed.TotalSeconds -lt $TimeoutSec) {
        # FindAllの結果がnullにならないよう空配列でキャスト
        $elements = @($Parent.FindAll($Scope, $searchCond))
        
        # Nameのワイルドカード処理 (Nameが指定されている場合のみ)
        if ($null -ne $Name -and $Name -ne "") {
            $elements = $elements | Where-Object { $_.Current.Name -like $Name }
        }
        
        if ($elements.Count -gt 0) { return $elements[0] }
        Start-Sleep -Milliseconds 500
    }
    throw "要素取得タイムアウト: Name=$Name, Id=$Id"
}

# --- [2] コントロール別 取得関数 ---

# デスクトップ直下からアプリウィンドウを探す（名前またはPID指定）
function Get-UiaAppWindow { 
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

    # 名前(Name)は PropertyCondition に入れず、後段のフィルタリング（Get-UiaControl内の-like）に任せる
    # ただし、PIDもNameも指定がない場合は、TrueConditionにしておく
    $finalCond = if ($conditions.Count -gt 1) {
        New-Object Windows.Automation.AndCondition(,[Windows.Automation.Condition[]]$conditions)
    } elseif ($conditions.Count -eq 1) {
        $conditions[0]
    } else {
        [Windows.Automation.Condition]::TrueCondition
    }

    # 2. Get-UiaControl を呼び出す。ここで $Name を渡せば、内部の Where-Object でワイルドカードが処理される
    return Get-UiaControl -Parent $root -Condition $finalCond -Name $Name -Scope Children
}

# 任意の親要素（Windowなど）の下にある子要素を探す
function Get-UiaWindow   { param($Parent, $Id, $Name) Get-UiaControl -Parent $Parent -Id $Id -Name $Name -Type ([Windows.Automation.ControlType]::Window) }
function Get-UiaWindows  { param($Parent, $Id, $Name) Get-UiaControls -Parent $Parent -Id $Id -Name $Name -Type ([Windows.Automation.ControlType]::Window) }

function Get-UiaPane     { param($Parent, $Id, $Name) Get-UiaControl -Parent $Parent -Id $Id -Name $Name -Type ([Windows.Automation.ControlType]::Pane) }
function Get-UiaPanes    { param($Parent, $Id, $Name) Get-UiaControls -Parent $Parent -Id $Id -Name $Name -Type ([Windows.Automation.ControlType]::Pane) }

function Get-UiaButton   { param($Parent, $Id, $Name) Get-UiaControl -Parent $Parent -Id $Id -Name $Name -Type ([Windows.Automation.ControlType]::Button) }
function Get-UiaText     { param($Parent, $Id, $Name) Get-UiaControl -Parent $Parent -Id $Id -Name $Name -Type ([Windows.Automation.ControlType]::Text) }
function Get-UiaEdit     { param($Parent, $Id, $Name) Get-UiaControl -Parent $Parent -Id $Id -Name $Name -Type ([Windows.Automation.ControlType]::Edit) }
function Get-UiaCheckBox { param($Parent, $Id)        Get-UiaControl -Parent $Parent -Id $Id -Type ([Windows.Automation.ControlType]::CheckBox) }
function Get-UiaComboBox { param($Parent, $Id)        Get-UiaControl -Parent $Parent -Id $Id -Type ([Windows.Automation.ControlType]::ComboBox) }
function Get-UiaRadio    { param($Parent, $Id)        Get-UiaControl -Parent $Parent -Id $Id -Type ([Windows.Automation.ControlType]::RadioButton) }
function Get-UiaMenuBar  { param($Parent, $Name)      Get-UiaControl -Parent $Parent -Name $Name -Type ([Windows.Automation.ControlType]::MenuBar) }
function Get-UiaMenuItem { param($Parent, $Name)      Get-UiaControl -Parent $Parent -Name $Name -Type ([Windows.Automation.ControlType]::MenuItem) }

# --- [3] 操作・ユーティリティ関数 ---

function Invoke-UiaElement {
    param([Parameter(Mandatory)] $Element)
    if ($Element.TryGetCurrentPattern([Windows.Automation.InvokePattern]::Pattern, [ref]$p)) { $p.Invoke() }
    elseif ($Element.TryGetCurrentPattern([Windows.Automation.TogglePattern]::Pattern, [ref]$p)) { $p.Toggle() }
    else { throw "Invoke非対応の要素です" }
}

function Expand-UiaElement {
    param([Parameter(Mandatory)] $Element)
    if ($Element.TryGetCurrentPattern([Windows.Automation.ExpandCollapsePattern]::Pattern, [ref]$p)) { $p.Expand() }
    else { throw "展開非対応の要素です" }
}

function Collapse-UiaElement {
    param([Parameter(Mandatory)] $Element)
    if ($Element.TryGetCurrentPattern([Windows.Automation.ExpandCollapsePattern]::Pattern, [ref]$p)) { $p.Collapse() }
    else { throw "折りたたみ非対応の要素です" }
}

function Close-UiaWindow {
    param([Parameter(Mandatory)] $Element)
    if ($Element.Current.ControlType -ne [Windows.Automation.ControlType]::Window) {
        throw "ウィンドウ以外の要素を閉じることはできません"
    }

    if ($Element.TryGetCurrentPattern([Windows.Automation.WindowPattern]::Pattern, [ref]$p)) { $p.Close() }
    else { throw "ウィンドウクローズ非対応の要素です" }
}

function Get-UiaName {
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

function Set-UiaValue {
    param([Parameter(Mandatory)] $Element, [Parameter(Mandatory)] [string]$Value)
    if ($Element.TryGetCurrentPattern([Windows.Automation.ValuePattern]::Pattern, [ref]$p)) { $p.SetValue($Value) }
    else { throw "値入力非対応の要素です" }
}

function Get-UiaValue {
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

function Get-UiaComboBoxItems {
    <#
    .SYNOPSIS
        コンボボックスを展開し、内部のアイテム一覧を String[] として取得します。
    #>
    param([Parameter(Mandatory)] [Windows.Automation.AutomationElement]$ComboBox)
    $expandPattern = $null
    if ($ComboBox.TryGetCurrentPattern([Windows.Automation.ExpandCollapsePattern]::Pattern, [ref]$expandPattern)) {
        $expandPattern.Expand()
        Start-Sleep -Milliseconds 200
    }
    $listItems = Get-UiaControls -Parent $ComboBox -Type ([Windows.Automation.ControlType]::ListItem) -Scope Descendants
    $itemNames = foreach ($item in $listItems) { $item.Current.Name }
    if ($null -ne $expandPattern) { $expandPattern.Collapse() }
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
    $window = Get-UiaAppWindow -Name $Name -ErrorAction SilentlyContinue

    if ($null -ne $window) {
        return $window
    }

    # 見つからない場合は新規起動する
    if ($null -ne $ExecutablePath -and (Test-Path $ExecutablePath)) {
        write-host "アプリを起動します: $ExecutablePath"
        $process = Start-Process -FilePath $ExecutablePath
        Start-UiaSleep -Sec 5 # 起動待ち
        $window = Get-UiaAppWindow -Name $Name -ProcessId $process.Id -ErrorAction SilentlyContinue
    }
    return $window
}

# --- [5] テストコード --- TeraTermを起動して接続ダイアログを閉じて、ヘルプのAbout TeraTermを開いて、閉じてExitする ---
$exePath = 'C:\Program Files\teraterm5\ttermpro.exe'
$name_app_window = '*Tera Term*'

try {
    [Windows.Automation.AutomationElement]$appWindow = GetAppWindow -Name $name_app_window -ExecutablePath $exePath
    if ($null -eq $appWindow) {
        throw "アプリウィンドウが見つかりませんでした: Name=$name_app_window"
    }

    # 接続ダイアログの「キャンセル」ボタンを押す
    $connectDialog = Get-UiaWindow -Parent $appWindow -Name '*New connection*'
    if ($null -ne $connectDialog) {
        $cancelButton = Get-UiaButton -Parent $connectDialog -Name 'Cancel'
        if ($null -ne $cancelButton) {
            Invoke-UiaElement -Element $cancelButton
            Start-UiaSleep
        }
    }

    # ヘルプ > Tera Termについて を開く
    $menuBar = Get-UiaMenuBar -Parent $appWindow
    if ($null -eq $menuBar) {
        throw "メニューバーが見つかりませんでした"
    }
    $menuHelp = Get-UiaMenuItem -Parent $menuBar -Name 'Help'
    if ($null -eq $menuHelp) {
        throw "ヘルプメニューが見つかりませんでした"
    }
    Expand-UiaElement -Element $menuHelp

    $menuAbout = Get-UiaMenuItem -Parent $menuHelp -Name 'About Tera Term*'
    if ($null -eq $menuAbout) {
        throw "About Tera Term メニューが見つかりませんでした"
    }
    Invoke-UiaElement -Element $menuAbout
    Start-UiaSleep -Seconds 5

    # Aboutダイアログを閉じる
    $aboutDialog = Get-UiaWindow -Parent $appWindow -Name 'About*'
    if ($null -ne $aboutDialog) {
        $okButton = Get-UiaButton -Parent $aboutDialog -Name 'OK'
        if ($null -ne $okButton) {
            Invoke-UiaElement -Element $okButton
            Start-UiaSleep
        }
    }

    # アプリを終了する
    Close-UiaWindow -Element $appWindow
} catch {
    Write-Error $_
}
