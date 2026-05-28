#
# 外来 - コメントにスメアと書かれたひとのシールを出す
#
# 追加アセンブリのロード
Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName Microsoft.VisualBasic
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

. "$PSScriptRoot\Tools-AutomationCore.ps1"
. "$PSScriptRoot\Tools-Entrance.ps1"

# アプリケーションウインドウ
$appKanLbl = $null

try {
    # エントランス 外来からデーターを取得
    $gairaiList = Get-GairaiList

    $csvdata = $gairaiList.List


    # 患者ラベルを起動
    $appKanLbl = GetAppWindow -Name '患者ラベル発行' -ExecutablePath 'C:\ssi\exe\srvKanLbl.exe'
    if ($null -eq $appKanLbl) {
        throw 'ラベル発行が起動しません.'
    }

    try {
        $appKanLbl.SetFocus()
    } catch {}

    # 患者ラベルのログイン操作が必要なら3分まで待つ
    $childWindow = $null

    try {
        $childWindow = Get-UIAControl -Parent $appKanLbl -Type ([Windows.Automation.ControlType]::Window) -Scope ([Windows.Automation.TreeScope]::Children) -TimeoutSec 1
    } catch {}
    if ($null -ne $childWindow) {
        Write-Host 'Waiting for user actions'
        $stopwatch = [System.Diagnostics.Stopwatch]::StartNew()
        while ($stopwatch.Elapsed.TotalSeconds -lt 180) {
            if ([UIATools]::IsAlive($childWindow)) {
                [UIATools]::Sleep(250)
            } else {
                break
            }
        }
        if ([UIATools]::IsAlive($childWindow)) {
            throw '操作(3分)がタイムアウトしました.'
        }
    }

    # ループ
    foreach($entry in $csvdata) {
        $valueId = $entry.'患者コード'
        $valueComment = @($entry.'診療予約コメント', $entry.'患者一覧用コメント') -join ' '

        write-host "$($valueId) コメントのチェック $($valueComment)"

        if ($valueComment -match '(スメア|すめあ)(?!結果)') {
            $btnChange = Get-UIAButton -Parent $appKanLbl -Id '4'
            $btnCommit = Get-UIAButton -Parent $appKanLbl -Id '7'
            $editId = Get-UIAEdit -Parent $appKanLbl -Id '8'

            if ($null -eq $btnChange -or $null -eq $btnCommit -or $null -eq $editId) {
                throw 'コントロールが取得できませんでした'
            }

            # 患者Idを設定
            Invoke-UIAElement $btnChange | Out-Null

            Set-UIAValue $editId $valueId | Out-Null
            $btnChange.SetFocus() | Out-Null
            [UIATools]::Sleep(300)
            
            # 印刷ボタンを操作
            Invoke-UIAElement $btnCommit | Out-Null
            [UIATools]::Sleep(200)
        }
    }

    # 患者ラベルは終了させずに放置する
} catch {
    throw $_
}

