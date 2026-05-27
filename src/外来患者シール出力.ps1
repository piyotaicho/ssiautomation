#
# 外来 - シールを出す
#
# 追加アセンブリのロード
Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName Microsoft.VisualBasic
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

. "$PSScriptRoot\automation-core.ps1"

# アプリケーションウインドウ
$appAplin = $null
$appKanLbl = $null

try {
    $pidAplin = (get-process -Name 'Entrance' -ErrorAction SilentlyContinue).Id

    if ($null -eq $pidAplin) {
        throw 'エントランスが起動していません'
    }

    # PIDからウインドウを取得
    $appAplin = GetAppWindow -Name '外来'
    if ($null -eq $appAplin) {
        throw 'エントランスから外来を開いてください'
    }

    # CSV保存して一覧を取得
    $menuCSV = Get-UIAMenuItem -Parent $appAplin -Name 'CSV出力...'
    Invoke-UIAElement $menuCSV
    [UIATools]::Sleep(300)

    #ファイル保存ダイアログで一時ファイル名でCSV保存
    $tempCSVfilename = ([System.IO.Path]::GetTempFileName() -replace '\..*$', '.csv')

    $windowFileSave = Get-UIAWindow -Parent $appAplin -Name '名前を付けて保存'
    
    Set-UIAValue -Element (Get-UIAComboBox -Parent $windowFileSave -Id 'FileNameControlHost') $tempCSVfilename
    Invoke-UIAElement (Get-UIAButton -Parent $windowFileSave -id '1')

    [UIATools]::Sleep(300)

    #CSVファイルを取得
    $csvdata = (Get-Content -Path $tempCSVfilename -Encoding Default | ConvertFrom-Csv)

    #CSVファイルを削除
    Remove-Item -Path $tempCSVfilename -ErrorAction SilentlyContinue | Out-Null

    if($csvdata.Count -eq 0) {
        throw '対象がありません'
    }

    write-host "Got $($csvdata.Count) entries"

    # 患者ラベルを起動
    $appKanLbl = GetAppWindow -Name '患者ラベル発行' -ExecutablePath 'C:\ssi\exe\srvKanLbl.exe'
    if ($null -eq $appKanLbl) {
        throw 'ラベル発行が起動しません.'
    }

    try {
        $appKanLbl.SetFocus()
    } catch {}

    # ログイン操作が必要なら待つ
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

    write-host 'Proceed automation...'

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

            write-host 'Print a label'

            write-host "Id $($valueId)"
            Invoke-UIAElement $btnChange | Out-Null

            Set-UIAValue $editId $valueId
            $btnChange.SetFocus()
            [UIATools]::Sleep(300)

            
            Invoke-UIAElement $btnCommit
            [UIATools]::Sleep(200)
        }
    }
} catch {
    throw $_
}

