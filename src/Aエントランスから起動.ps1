#
# エントランスから希望のアプリを起動する
#
# エントランスのカテゴリ $Category の $ItemName のものを起動する、ユーザーの権限で表示されないものは実行できないので注意
#
Param (
    [Parameter(Mandatory = $true)] [string]$Category,
    [Parameter(Mandatory = $true)] [string]$ItemName
)

# アセンブリをロード
. .\Tools-AutomationCore.ps1
. .\Tools-CheckLogin.ps1

if (-not $Category -or -not $ItemName) {
    throw '起動オプション -Category カテゴリ名 -ItemName 対象アイテム名 が不足しています.'
}

# ログイン状態を取得
if (-not [SSICheckLogin]::check()) {
    throw '電子カルテシステムにログインしてください.'
}

# オートメーションの開始
$appWindow = $null
try {
    $appWindow = GetAppWindow -Name 'エントランス -*'
} catch {
    Write-Error $_
}
if ($null -eq $appWindow) {
    throw 'システムの状態が不正です. エントランスがみつかりません.'
}

$PaneContents = Get-UIAPane -Parent $appWindow -Id 'contents'
if ($null -eq $PaneContents) {
    throw 'アプリケーションの構成が不正です.'
}
$PaneContents = Get-UIAPane -Parent $PaneContents -Id 'tlpContents'
if ($null -eq $PaneContents) {
    throw 'アプリケーションの構成が不正です.'
}

# tlpContentsの下にあるPanesから必要なものを取得する - 速度重視でChildrenを使用する
$PaneCategory = $null
$PaneCategory = Get-UIAControl -Parent $PaneContents -Id 'flpBusinessItems' -Type ([Windows.Automation.ControlType]::Pane) -Scope ([Windows.Automation.TreeScope]::Children)
if ($null -eq $PaneCategory) {
    throw 'アプリケーションの構成が不正です.'
}

$ButtonCategory = $null
try {
    $ButtonCategory = Get-UIAButton -Parent $PaneCategory -Name $Category
    Invoke-UIAElement $ButtonCategory | Out-Null
} catch {
    Write-Error $_
}
if ($null -eq $ButtonCategory) {
    throw "アプリのカテゴリー $Category がみつかりませんでした."
}

[UIATools]::Sleep(500)

# tlpContentsの下にあるPanesから必要なものを取得する - 速度重視でChildren
$PaneItems = $null
$PaneItems = Get-UIAControl -Parent $PaneContents -Id 'flpLauncherItems' -Type ([Windows.Automation.ControlType]::Pane) -Scope ([Windows.Automation.TreeScope]::Children)
if ($null -eq $PaneItems) {
    throw 'アプリケーションの構成が不正です.'
}

$ButtonItem = $null
try {
    $ButtonItem = Get-UIAButton -Parent $PaneItems -Name $ItemName
    Invoke-UIAElement $ButtonItem | Out-Null
} catch {
    Write-Error $_
}
if ($null -eq $ButtonItem) {
    throw "アプリのアイテム $ItemName がみつかりませんでした."
}

# END
