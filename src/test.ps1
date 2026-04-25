# アセンブリのロード
Add-Type -AssemblyName "UIAutomationClient"
Add-Type -AssemblyName "UIAutomationTypes"
# UIAutomationClientProvidersをロードする必要があるが、PowerShell単体ではロードできないので
# C#コードをAdd-Typeでコンパイルして、UIAutomationClientの型を参照してUIAutomationClientProvidersの機能を呼び出す
$sourceGetMainWindow = @"
using System;
using System.Text.RegularExpressions;
using System.Windows.Automation;
public class UIATools
{
        public static AutomationElement RootElement
        {
            get
            {
                return AutomationElement.RootElement;
            }
        }
        public static AutomationElement GetMainWindowByName(string name) {
            if (string.IsNullOrEmpty(name)) {
                return null;
            }

            string escaped = Regex.Escape(name).Replace("\\*", ".*");
            Regex wildcardRegex = new Regex("^" + escaped + "$", RegexOptions.IgnoreCase);
            Condition cond = new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Window);
            AutomationElementCollection windows = RootElement.FindAll(TreeScope.Element | TreeScope.Children, cond);
            foreach (AutomationElement window in windows) {
                if (wildcardRegex.IsMatch(window.Current.Name ?? string.Empty)) {
                    return window;
                }
            }
            return null;
        }

        public static AutomationElement GetMainWindowByProcessID(int processId) {
            if (processId <= 0) {
                return null;
            }

            Condition cond = new AndCondition(
                new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Window),
                new PropertyCondition(AutomationElement.ProcessIdProperty, processId)
            );
            return RootElement.FindFirst(TreeScope.Element | TreeScope.Children, cond);
    }
}
"@
if (-not ([System.Management.Automation.PSTypeName]'UIATools').Type) {
    Add-Type -TypeDefinition $sourceGetMainWindow -Language CSharp -ReferencedAssemblies("UIAutomationClient", "UIAutomationTypes")
}

# --- [0] 版用ユーティリティ関数 ---
function Start-UIASleep {
    <#
    .SYNOPSIS
        UI操作の合間に待機を挟みます。デフォルトは 200ms です。
    #>
    param(
        [Parameter(Position = 0)]
        [int]$Milliseconds,

        [int]$Seconds
    )

    if ($PSBoundParameters.ContainsKey('Seconds') -and $PSBoundParameters.ContainsKey('Milliseconds')) {
        throw "-Seconds と -Milliseconds は同時に指定できません"
    }

    if ($PSBoundParameters.ContainsKey('Seconds')) {
        # 秒が指定された場合
        Start-Sleep -Seconds $Seconds
    } elseif ($PSBoundParameters.ContainsKey('Milliseconds')) {
        # ミリ秒が指定された場合
        Start-Sleep -Milliseconds $Milliseconds
    } else {
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

    $hasWildcardName = (-not [string]::IsNullOrEmpty($Name)) -and $Name.Contains('*')
    
    $searchCond = $Condition
    if ($null -eq $searchCond) {
        $conditions = @()
        if ($Id) { 
            write-host "Building search condition from parameters: Id='$Id'"
            $conditions += New-Object Windows.Automation.PropertyCondition([Windows.Automation.AutomationElement]::AutomationIdProperty, $Id)
        }
        if ($Name -and -not $hasWildcardName) {
            write-host "Building search condition from parameters: Name='$Name'"
            $conditions += New-Object Windows.Automation.PropertyCondition([Windows.Automation.AutomationElement]::NameProperty, $Name)
        }
        if ($Type) {
            Write-Host "Building search condition from parameters: Type='$($Type.LocalizedControlType)'"
            $conditions += New-Object Windows.Automation.PropertyCondition([Windows.Automation.AutomationElement]::ControlTypeProperty, $Type)
        }

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
        [Windows.Automation.TreeScope]$Scope = [Windows.Automation.TreeScope]::Descendants,
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
        $conditions = @()
        if ($Id) {
            write-host "Building search condition from parameters: Id='$Id'"
            $conditions += New-Object Windows.Automation.PropertyCondition([Windows.Automation.AutomationElement]::AutomationIdProperty, $Id)
        }
        if ($Name) {
            write-host "Building search condition from parameters: Name='$Name'"
            $conditions += New-Object Windows.Automation.PropertyCondition([Windows.Automation.AutomationElement]::NameProperty, $Name)
        }
        if ($Type) {
            write-host "Building search condition from parameters: Type='$($Type.LocalizedControlType)'"
            $conditions += New-Object Windows.Automation.PropertyCondition([Windows.Automation.AutomationElement]::ControlTypeProperty, $Type)
        }
        
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
    # $allelements = $Parent.FindAll([Windows.Automation.TreeScope]::Descendants, [Windows.Automation.Condition]::TrueCondition)
    # write-host "Dump all elements under parent: $($Parent.Current.Name)"
    # foreach ($el in $allelements) {
    #     write-host "  Element: Name='$($el.Current.Name)', Id='$($el.Current.AutomationId)', Type='$($el.Current.ControlType.LocalizedControlType)'"
    # }
    # write-host "--"

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

    # 1. 確実に判定できるPIDがある場合はそれで対応
    if ($ProcessId -gt 0) { 
        return [UIATools]::GetMainWindowByProcessId($ProcessId)
    }

    # 2. Nameで検索
    if ($null -ne $Name) {
        return [UIATools]::GetMainWindowByName($Name)
    }

    return $null
}

# 任意の親要素（Windowなど）の下にある子要素を探す
function Get-UIAWindow   { param($Parent, $Id, $Name) Get-UIAControl -Parent $Parent -Id $Id -Name $Name -Type ([Windows.Automation.ControlType]::Window) }
function Get-UIAWindows  { param($Parent, $Id, $Name) Get-UIAControls -Parent $Parent -Id $Id -Name $Name -Type ([Windows.Automation.ControlType]::Window) }

function Get-UIAPane     { param($Parent, $Id, $Name) Get-UIAControl -Parent $Parent -Id $Id -Name $Name -Type ([Windows.Automation.ControlType]::Pane) }
function Get-UIAPanes    { param($Parent, $Id, $Name) Get-UIAControls -Parent $Parent -Id $Id -Name $Name -Type ([Windows.Automation.ControlType]::Pane) }

function Get-UIAButton   { param($Parent, $Id, $Name) Get-UIAControl -Parent $Parent -Id $Id -Name $Name -Type ([Windows.Automation.ControlType]::Button) }
function Get-UIAText     { param($Parent, $Id, $Name) Get-UIAControl -Parent $Parent -Id $Id -Name $Name -Type ([Windows.Automation.ControlType]::Text) }
function Get-UIAEdit     { param($Parent, $Id, $Name) Get-UIAControl -Parent $Parent -Id $Id -Name $Name -Type ([Windows.Automation.ControlType]::Edit) }
function Get-UIACheckBox { param($Parent, $Id, $Name) Get-UIAControl -Parent $Parent -Id $Id -Name $Name -Type ([Windows.Automation.ControlType]::CheckBox) }
function Get-UIAComboBox { param($Parent, $Id, $Name) Get-UIAControl -Parent $Parent -Id $Id -Name $Name -Type ([Windows.Automation.ControlType]::ComboBox) }
function Get-UIAList     { param($Parent, $Id, $Name) Get-UIAControl -Parent $Parent -Id $Id -Name $Name -Type ([Windows.Automation.ControlType]::List) }
function Get-UIARadio    { param($Parent, $Id, $Name) Get-UIAControl -Parent $Parent -Id $Id -Name $Name -Type ([Windows.Automation.ControlType]::RadioButton) }
function Get-UIAMenuBar  { param($Parent, $Id, $Name) Get-UIAControl -Parent $Parent -Id $Id -Name $Name -Type ([Windows.Automation.ControlType]::MenuBar) }
function Get-UIAMenuItem { param($Parent, $Id, $Name) Get-UIAControl -Parent $Parent -Id $Id -Name $Name -Type ([Windows.Automation.ControlType]::MenuItem) }
function Get-UIADataGrid { param($Parent, $Id, $Name) Get-UIAControl -Parent $Parent -Id $Id -Name $Name -Type ([Windows.Automation.ControlType]::DataGrid) }
function Get-UIATable    { param($Parent, $Id, $Name) Get-UIAControl -Parent $Parent -Id $Id -Name $Name -Type ([Windows.Automation.ControlType]::Table) }

# --- [3] 操作・ユーティリティ関数 ---

function Invoke-UIAElement {
    param([Parameter(Mandatory, Position = 0)] $Element)
    $invokePattern = $null
    $togglePattern = $null
    if ($Element.TryGetCurrentPattern([Windows.Automation.InvokePattern]::Pattern, [ref]$invokePattern)) {
        $invokePattern.Invoke()
    } elseif ($Element.TryGetCurrentPattern([Windows.Automation.TogglePattern]::Pattern, [ref]$togglePattern)) {
        $togglePattern.Toggle()
    } else {
        throw "Invoke非対応の要素です"
    }
}

function Set-UIAElementExpanded {
    param([Parameter(Mandatory, Position = 0)] $Element)
    $expandCollapsePattern = $null
    if ($Element.TryGetCurrentPattern([Windows.Automation.ExpandCollapsePattern]::Pattern, [ref]$expandCollapsePattern)) {
        $expandCollapsePattern.Expand()
    } else {
        throw "展開非対応の要素です"
    }
}

function Set-UIAElementCollapsed {
    param([Parameter(Mandatory, Position = 0)] $Element)
    $expandCollapsePattern = $null
    if ($Element.TryGetCurrentPattern([Windows.Automation.ExpandCollapsePattern]::Pattern, [ref]$expandCollapsePattern)) {
        $expandCollapsePattern.Collapse()
    } else {
        throw "折りたたみ非対応の要素です"
    }
}

function Close-UIAWindow {
    param([Parameter(Mandatory, Position = 0)] $Element)
    if ($Element.Current.ControlType -ne [Windows.Automation.ControlType]::Window) {
        throw "ウィンドウ以外の要素を閉じることはできません"
    }

    $windowPattern = $null
    if ($Element.TryGetCurrentPattern([Windows.Automation.WindowPattern]::Pattern, [ref]$windowPattern)) {
        $windowPattern.Close()
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
        [Parameter(Mandatory, Position = 0)]
        [Windows.Automation.AutomationElement]$Element
    )

    return $Element.Current.Name
}

function Set-UIAValue {
    param([Parameter(Mandatory, Position = 0)] $Element, [Parameter(Mandatory, Position = 1)] [string]$Value)
    $valuePattern = $null
    if ($Element.TryGetCurrentPattern([Windows.Automation.ValuePattern]::Pattern, [ref]$valuePattern)) {
        $valuePattern.SetValue($Value)
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
        [Parameter(Mandatory, Position = 0)]
        [Windows.Automation.AutomationElement]$Element
    )

    # 1. ValuePattern を持っている場合 (Edit, ComboBoxなど)
    $valuePattern = $null
    if ($Element.TryGetCurrentPattern([Windows.Automation.ValuePattern]::Pattern, [ref]$valuePattern)) {
        return $valuePattern.Current.Value
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
    param([Parameter(Mandatory, Position = 0)] [Windows.Automation.AutomationElement]$ComboBox)
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

function Get-UIAListItems {
    <#
    .SYNOPSIS
        リストコントロールのアイテム一覧を String[] として取得します。
    #>
    param([Parameter(Mandatory, Position = 0)] [Windows.Automation.AutomationElement]$List)
    $listItems = Get-UIAControls -Parent $List -Type ([Windows.Automation.ControlType]::ListItem) -Scope Descendants
    return [string[]]($listItems | ForEach-Object { $_.Current.Name })
}

function Get-UIATableContents {
    <#
    .SYNOPSIS
        TableやDataGridの内容を headerとvaluesで取得します。
        ヘッダーが無い場合、headerは空配列になります。
    #>
    param(
        [Parameter(Mandatory, Position = 0)]
        [Alias('Table', 'DataGrid')]
        [Windows.Automation.AutomationElement]$Element
    )

    $supportedControlTypes = @(
        [Windows.Automation.ControlType]::Table,
        [Windows.Automation.ControlType]::DataGrid
    )
    if ($supportedControlTypes -notcontains $Element.Current.ControlType) {
        throw "Get-UIATableContents は Table または DataGrid 要素のみ対応しています"
    }

    $headers = Get-UIAControls -Parent $Element -Type ([Windows.Automation.ControlType]::Header) -Scope Descendants
    $headerNames = foreach ($header in $headers) { $header.Current.Name }
    $rows = Get-UIAControls -Parent $Element -Type ([Windows.Automation.ControlType]::DataItem) -Scope Descendants
    $values = foreach ($row in $rows) {
        # 多くの実装では DataItem の直下がセルだが、実装差異を考慮して子要素優先 + 子孫要素へフォールバックする
        $cells = Get-UIAControls -Parent $row -Scope Children -Condition ([Windows.Automation.Condition]::TrueCondition)
        if ($cells.Count -eq 0) {
            $cells = Get-UIAControls -Parent $row -Scope Descendants -Condition ([Windows.Automation.Condition]::TrueCondition)
        }

        $cellValues = foreach ($cell in $cells) {
            Get-UIAValue -Element $cell
        }
        ,$cellValues # 2次元配列にするため、行ごとにカンマで配列化
    }
    return @{
        Header = [string[]]$headerNames
        Values = [string[][]]$values
    }
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
        $process = Start-Process -FilePath $ExecutablePath -PassThru
        Start-UIASleep -Sec 5 # 起動待ち
        try {
            $window = Get-UIAAppWindow -ProcessId $process.Id
        } catch {}
    }
    return $window
}

#
# --- テストコード ---
#
$exePath = 'C:\Program Files (x86)\TeraPad\TeraPad.exe'
$name_app_window = '* TeraPad'

try {
    [Windows.Automation.AutomationElement]$appWindow = GetAppWindow -Name $name_app_window -ExecutablePath $exePath
    if ($null -eq $appWindow) {
        throw "アプリウィンドウが見つかりませんでした: Name=$name_app_window"
    }
    Start-UIASleep -Milliseconds 300

    # ヘルプを開く
    $menuBar = Get-UIAMenuBar -Parent $appWindow -Id 'MenuBar'
    if ($null -eq $menuBar) {
        throw "メニューバーが見つかりませんでした"
    }
    $menuHelp = Get-UIAMenuItem -Parent $menuBar -Name 'ヘルプ*'
    if ($null -eq $menuHelp) {
        throw "ヘルプメニューが見つかりませんでした"
    }
    Set-UIAElementExpanded -Element $menuHelp
    Start-UIASleep -Milliseconds 500

    $menuAbout = Get-UIAMenuItem -Parent $menuHelp -Name '*情報*'
    if ($null -eq $menuAbout) {
        throw "Aboutメニューが見つかりませんでした"
    }
    Invoke-UIAElement -Element $menuAbout
    Start-UIASleep -Seconds 5

    # Aboutダイアログを閉じる
    $aboutDialog = Get-UIAWindow -Parent $appWindow -Name 'バージョン情報'
    if ($null -eq $aboutDialog) {
        throw "Aboutダイアログが見つかりませんでした"
    }
    Start-UIASleep -Seconds 2

    $aboutCloseButton = Get-UIAButton -Parent $aboutDialog -Name 'OK'
    if ($null -eq $aboutCloseButton) {
        throw "Aboutダイアログの閉じるボタンが見つかりませんでした"
    }
    Invoke-UIAElement -Element $aboutCloseButton
    Start-UIASleep -Milliseconds 300

    # アプリを終了する
    Close-UIAWindow -Element $appWindow
} catch {
    Write-Error $_
}
