#
# 文書管理にスキャンして画像を登録する
#
# 追加アセンブリのロード
Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName Microsoft.VisualBasic
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

. "$PSScriptRoot\automation-core.ps1"

######################################################################　スクリプト変数
$appWindow = $null

######################################################################　サブルーチン
# 必要に応じて文書リスト画面などから設定画面にfallbackして設定ウインドウを取得する
function Get-SetteiWindow {
    param(
        [Parameter(Mandatory)][Windows.Automation.AutomationElement]$appWindow
    )
    if ( $null -eq $AppWindow ) {
        return $null
    }

    # 子ウインドウを取得 最大10回のループで抜けられなかったら例外
    $watchDog = 10
    do {
        $childWindow = $null
        try {
            [UIATools]::Sleep()
            $childWindow = Get-UIAControl -Parent $AppWindow -Type ([Windows.Automation.ControlType]::Window) -Scope ([Windows.Automation.TreeScope]::Children) -TimeoutSec 1
        }
        catch {}

        if ( $null -eq $childWindow ) {
            # 子ウインドウが開いていないのでリスト画面
            # リスト画面からのfallbackはメニュー操作
            Invoke-ListNewMenu -AppWindow $AppWindow
        }
        else {
            $windowTitle = $childWindow.Current.Name
            # 設定画面なので処理は不要でウインドウを返す
            if ( $windowTitle -eq '設定') {
                return $childWindow
            }

            # 当該ウインドウを閉じて戻る～スキャナのコントロールなどで深かったらエラーになる
            $buttons = Get-UIAControls -Parent $childWindow -Name '*閉じる*' -Type ([Windows.Automation.ControlType]::Button)
            if ( $null -eq $buttons -or $buttons.Count -eq 0 ) {
                $buttons = Get-UIAControls -Parent $childWindow -Name '*Close*' -Type ([Windows.Automation.ControlType]::Button)
            }
            if ( $null -ne $buttons -and $buttons.Count -gt 0 ) {
                // 一番最後にあらわれた閉じるボタンを操作する
                $button = $buttons[-1]
                Invoke-UIAElement $button
            }
        }
        $watchdog--
    } while ( $watchdog -gt 0 )
    throw '子ウインドウに閉じるボタンがありません'
}

# リスト画面でメニューから新規を操作する
function Invoke-ListNewMenu {
    param(
        [Parameter(Mandatory)] [Windows.Automation.AutomationElement]$appWindow
    )
    $menubar = Get-UIAMenuBar -Parent $AppWindow -Name ''
    # $menuitem = Get-UIAMenuItem -Parent $menubar -Name 'ファイル(F)'
    # if ($null -eq $menuitem ) {
    #     throw 'メニュー取得(ファイル)に問題が発生しました'
    # }
                
    # $menuitem = Get-UIAMenuItem -Parent $menuitem -Name '新規(N)...'
    $menuitem = Get-UIAMenuItem -Parent $menubar -Name '新規(N)...'
    if ($null -eq $menuitem ) {
        throw 'メニュー取得(新規)に問題が発生しました'
    }
    Invoke-UIAElement $menuitem
}

# 設定画面の内容を入力して次に進む
function Set-SetteiWindow {
    param(
        [Parameter(Mandatory)] [Windows.Automation.AutomationElement]$setteiWindow,
        [Parameter(Mandatory)] [string]$kanjyaId,
        [string]$ishiCode,
        [string]$kaCode,
        [string]$nyuugai
    )

    # 患者ID
    if ($null -eq $kanjyaId -or $kanjyaId -notmatch '^(0?\d{2,7}|9\d{7})$') {
        throw '患者IDの指定が正しくありません'
    }
    $editKanjyaId = Get-UIAEdit -Parent $setteiWindow -Id 'TxtCode'
    Set-UIAValue $editKanjyaId $kanjyaId

    # ボタンにフォーカスしてエラーダイアログをチェック
    $button = Get-UIAButton -Parent $setteiWindow -Id 'Cmd_000'
    $button.SetFocus()
    [UIATools]::Sleep()
    $errorDialogs = Get-UIAWindows -Parent $setteiWindow -Id 'SSIMessageDialog'
    if ($errorDialogs.Count -gt 0) {
        throw '不正な患者IDです'
    }

    # 医師ID
    if ($ishiCode.Trim()) {
        $editIshiCode = Get-UIAEdit -Parent $setteiWindow -Id 'TxtDrcd'
        Set-UIAValue $editIshiCode $ishiCode

        # ボタンにフォーカスしてエラーダイアログをチェック
        $button.SetFocus()
        [UIATools]::Sleep()
        $errorDialogs = Get-UIAWindows -Parent $setteiWindow -Id 'SSIMessageDialog'
        if ($errorDialogs.Count -gt 0) {
            throw '不正な医師IDです'
        }
    }

    # 入外区分
    if ($nyuugai.Trim()) {
        $comboNyuugai = Get-UIAComboBox -Parent $setteiWindow -Id 'CboNgkb'
        if ($nyuugai -match '入') {
            Set-UIAComboBoxValue $comboNyuugai '3：入院'
        }
        else {
            Set-UIAComboBoxValue $comboNyuugai '1：外来'
        }
    }

    # 診療科
    if ($kaCode.Trim()) {
        $editShinryoka = Get-UIAComboBox -Parent $setteiWindow -Id 'CboSnk'
        Set-UIAComboBoxValue $editShinryoka $kaCode
    }

    # ボタンを押して続行、設定ウインドウが閉じていれば問題なし、閉じていなくてエラーダイアログが出ていればエラー
    Invoke-UIAElement $button
    [UIATools]::Sleep(300)
    $isError = $false
    try {
        if ($setteiWindow.Current.Name) {
            $errorDialogs = Get-UIAWindows -Parent $setteiWindow -Id 'SSIMessageDialog'
            if ($errorDialogs.Count -gt 0) {
                $isError = $true
            }
        }   
    }
    catch {}
    if ($isError) {
        throw '設定に誤りがあります'
    }
}

# 新規画像ウインドウの内容を設定
function Set-NewImageWindowContent {
    param(
        [parameter(Mandatory)][Windows.Automation.AutomationElement]$NewImageWindow,
        [string]$DocumentName,
        [string]$Description,
        [string]$Date
    )
    if ($null -eq $NewImageWindow) {
        throw 'Windowが指定されていません'
    }
    # フォーカス退避用に戻るボタンを取得しておく
    $cmdBackBtn = Get-UIAButton -Parent $NewImageWindow -Id 'cmdBack'

    # 文書名
    if ($DocumentName.Trim()) {
        # Write-Host '名称の入力'
        $titleCmb = Get-UIAComboBox -Parent $NewImageWindow -Id 'txtFile'

        # SetValueなど一切をうけつけないコントロールでIMEもコントロールできないので文字列をクリップボード経由で貼り付ける
        Set-Clipboard $DocumentName
        Set-UIAValue -Element $titleCmb -Value '^a{DEL}^v' -Force -OmitEscape
        $cmdBackBtn.SetFocus()
    }

    # 補足
    if ($Description.Trim()) {
        # Write-Host '補足の入力'
        $descriptionEdit = Get-UIAControl -Parent $NewImageWindow -Id 'txtBikou'

        # SetValueなど一切をうけつけないコントロールでIMEもコントロールできないので文字列をクリップボード経由で貼り付ける
        Set-Clipboard $Description.Trim()
        Set-UIAValue -Element $descriptionEdit -Value '^a{DEL}^v' -Force -OmitEscape
        $cmdBackBtn.SetFocus()
    }

    # カレンダー ThreeCalendarDialogを操作
    if ($Date -match '^20\d\d[/-]((0?[13578]|1[02])[/-](0?[1-9]|[12][0-9]|3[01])|(0?2[/-](0?[1-9]|1[0-9]|2[0-9]))|((0?[469]|11)[/-](0?[1-9]|[12][0-9]|30)))$') {
        # Write-Host '日付の設定'
        # 日付を分解
        $targetDateStr = $Date.Split('/-')
        $targetInMonth = [int]$targetDateStr[0] * 12 + [int]$targetDateStr[1]
        $targetDay = [int]$targetDateStr[2]

        # 文書管理のカレンダーを開く(力業しかない)
        $calenderButtonPane = Get-UIAPane -Parent (Get-UIAPane -Parent $NewImageWindow -Id 'pnlDate')

        # 現在の日付を取得して変更が必要ならばカレンダーを操作する
        $calenderDateStr = ($calenderButtonPane.Current.Name).Split('年月日')
        if ([int]$targetDateStr[0] -ne [int]$calenderDateStr[0] -or [int]$targetDateStr[1] -ne [int]$calenderDateStr[1] -or [int]$targetDateStr[2] -ne [int]$calenderDateStr[2]) {
            [UIATools]::ForceClick($calenderButtonPane)
            [UIATools]::Sleep()
            
            # カレンダーダイアログを取得
            $calendarWindow = Get-UIAWindow -Parent $NewImageWindow -Id 'ThreeCalendarDialog'
            if ($null -eq $calendarWindow) {
                throw 'カレンダーが開きません'
            }

            # 文書管理のカレンダーを操作して日付を選択する
            # 日付操作コントロールの取得
            $btnDec = Get-UIAButton -Parent $calendarWindow -Id 'btnPrevMonthRight'
            $btnInc = Get-UIAButton -Parent $calendarWindow -Id 'btnNextMonthRight'
            # 中央のカレンダーを利用する
            $calendarPane = Get-UIAPane -Parent $calendarWindow -Id 'ctlCalendar2'
            $txtYear = Get-UIAText -Parent $calendarPane -Id 'lblYear'
            $txtMonth = Get-UIAText -Parent $calendarPane -Id 'lblMonth'

            if ($null -eq $btnDec -or $null -eq $btnInc -or $null -eq $calendarPane -or $null -eq $txtYear -or $null -eq $txtMonth) {
                throw 'カレンダーに異常があります'
            }

            # 月を動かす(120ヶ月まで)
            $moveCountLimit = 120
            do {
                $currentInMonth = [int]((Get-UIAName $txtYear).Substring(0, 4)) * 12 + [int]((Get-UIAName $txtMonth).Substring(0, 2))
                if ($currentInMonth -eq $targetInMonth) {
                    break
                }
                elseif ($currentInMonth -gt $targetInMonth) {
                    Invoke-UIAElement $btnDec
                }
                else {
                    Invoke-UIAElement $btnInc
                }
                $moveCountLimit--
            } while ($moveCountLimit -gt 0)
            if ($moveCountLimit -eq 0) {
                throw '10年以上の移動はサポートしていません'
            }

            # 日を選択 (～すると文書管理のカレンダーウインドウは閉じる)
            $datePane = Get-UIAPane -Parent (Get-UIAPane -Parent $calendarPane)
            $dayBtn = Get-UIAButton -Parent $datePane -Name ($targetDay.ToString())
            if ($null -eq $dayBtn) {
                throw 'カレンダーに日付がありません'
            }
            Invoke-UIAElement $dayBtn
        }
    }
    # 念のためフォーカスを移動しておく
    $cmdBackBtn.SetFocus()
}

# スキャンして画像を登録する
function Register-ScanImage {
    param(
        [parameter(Mandatory)][Windows.Automation.AutomationElement]$NewImageWindow,
        [boolean]$Auto = $false
    )

    if ($null -eq $NewImageWindow) {
        throw 'ウインドウが指定されていません'
    }

    # スキャンボタンを押す
    $cmdScanBtn = Get-UIAButton -Parent $NewImageWindow -Id 'cmdScan'
    if ($null -eq $cmdScanBtn) {
        throw 'ウインドウが異なるようです'
    }
    Invoke-UIAElement $cmdScanBtn

    # オートモードの時はスキャンボタンを押す
    if ($Auto) {
        # スキャナのUIウインドウの出現を待つ(最大30秒)
        $scanWindow = $null
        try {
            $scanWindow = Get-UIAControl -Parent $Element -Type ([Windows.Automation.ControlType]::Window) -Scope ([Windows.Automation.TreeScope]::Children) -TimeoutSec 30
        } catch {
        }
        if ($null -ne $scanWindow) {
            write-host 'Got window immidiately'
            write-host "Window name is $($scanWindow.Current.Name)($($scanWindow.Current.NativeWindowHandle.ToString('X')))"
        } else {
            write-host 'Got NO window immidiately'
        }   

        # スキャナウインドウが表示されるまで待機
        [UIATools]::Sleep(5000)

        # スキャナのUIウインドウの出現を待つ(最大30秒)
        $scanWindow = $null
        try {
            $scanWindow = Get-UIAControl -Parent $Element -Type ([Windows.Automation.ControlType]::Window) -Scope ([Windows.Automation.TreeScope]::Children) -TimeoutSec 30
        } catch {
        }
        if ($null -ne $scanWindow) {
            write-host 'Got window'
            write-host "Window name is $($scanWindow.Current.Name)($($scanWindow.Current.NativeWindowHandle.ToString('X')))"
        } else {
            write-host 'Got NO window'
        }   

        if ($null -eq $scanWindow) {
            throw 'スキャナの接続に問題があります'
        }


        $btnScan = Get-UIAButtons -Parent $scanWindow -Name 'スキャン*'
        if ($null -eq $btnScan -or $btnScan.Count -eq 0) {
            throw 'スキャナウインドウにスキャンボタンがありません'
        }
        Write-Host "Got $($btnScan.Count) buttons!"

        Invoke-UIAElement $btnScan[0]

        # スキャナウインドウの消失を最大3分待つ 1秒に2回ウインドウの存在確認
        # autoでない場合はこの間に複数のドキュメントをスキャンしておく
        $stopwatch = [System.Diagnostics.Stopwatch]::StartNew()
        while ($stopwatch.Elapsed.TotalSeconds -lt 180) {
            try {
                $hwnd = $scanWindow.Current.NativeWindowHandle
                if ($hwnd -eq [IntPtr]::Zero) {
                    break
                }
                if ([UIATools]::IsWindow($hwnd) -eq $false) {
                    break
                }
            } catch {
                break
            }
            [UIATools]::Sleep(500)
        }
        if ($null -ne $scanWindow.Current.Name) {
            throw 'スキャナの反応に問題があります'
        }
    
        # 画像の登録ボタンを押す
        Invoke-UIAElement (Get-UIAButton -Parent $Element -Id 'cmdSave')
        [UIATools]::Sleep()

        # 確認ダイアログのボタンを操作(もしスキャンができていなくてもダイアログは閉じる)
        Invoke-UIAElement (Get-UIAButton -Parent (Get-UIAWindow -Parent $Element -Id 'SSIMessageDialog') -Id 'Button1')
        [UIATools]::Sleep()

        # この操作で新規画像ウインドウは閉じている
    }

    # オートモード非対応などの場合は強制的に半自動モードになる
    # 半自動モードでは新規画像ウインドウの操作を待機する
    if ([UIATools]::IsAlive($NewImageWindow)) {
        # 新規画像ウインドウでのスキャナ操作と登録を最大10分待つ 1秒に4回ウインドウの存在確認
        $stopwatch = [System.Diagnostics.Stopwatch]::StartNew()
        while ($stopwatch.Elapsed.TotalSeconds -lt 600) {
            if ([UIATools]::IsAlive($NewImageWindow)) {
                [UIATools]::Sleep(250)
            } else {
                break
            }
            #try {
            #    $hwnd = $Element.Current.NativeWindowHandle
            #    if ($hwnd -eq [IntPtr]::Zero) {
            #        break
            #    }
            #    if ([UIATools]::IsWindow($hwnd) -eq $false) {
            #        break
            #    }
            #} catch {
            #    break
            #}
            #[UIATools]::Sleep(250)
        }
        if ([UIATools]::IsAlive($NewImageWindow)) {
            throw '手動スキャンの操作(10分)がタイムアウトしました.'
        }
    }
}

######################################################################　スキャナ別取り込みルーチン
# スキャナ設定画面の操作 (EPSON Scan)
function Set-EPSONScan {
    param (
        [parameter(Mandatory)][Windows.Automation.AutomationElement]$ScanWindow
    )


}

# スキャナ設定画面の操作 (EPSON Scan2)
function Set-EPSONScan {
    param (
        [parameter(Mandatory)][Windows.Automation.AutomationElement]$ScanWindow
    )
}

######################################################################　フォームルーチン
function Show-ControlForm {
    param (
        [string[]]$kaList
    )
    # チェック
    if ($null -eq $appWindow) {
        throw 'ウインドウの指定がありません'
    }

    # フォーム設定
    $form = New-Object System.Windows.Forms.Form
    $form.Text = '文書スキャン取り込み'
    $form.Size = New-Object System.Drawing.Size(390,480)
    $form.StartPosition = 'CenterScreen'
    $form.FormBorderStyle = 'FixedDialog'
    
    $fontLabel = New-Object System.Drawing.Font('Yu Gothic UI', 11)

    $currentY = 20
    function Add-Label ($text, $posY) {
        $label = New-Object System.Windows.Forms.Label
        $label.Text = $text
        $label.Location = New-Object System.Drawing.Point(20, ($posY + 3))
        $label.Size = New-Object System.Drawing.Size(90, 20)
        $label.Font = $fontLabel
        $form.Controls.Add($label)
    }

    # コントロールの配置
    Add-Label 'ID(必須):' $currentY
    $txtId = New-Object System.Windows.Forms.TextBox
    $txtId.Location = New-Object System.Drawing.Point(110, $currentY)
    $txtId.Size = New-Object System.Drawing.Size(250, 25)
    $txtId.Font = $fontLabel
    $txtId.ImeMode = [System.Windows.Forms.ImeMode]::Disable
    $form.Controls.Add($txtId)
    $currentY += 40

    Add-Label '名称(必須):' $currentY
    $txtTitle = New-Object System.Windows.Forms.TextBox
    $txtTitle.Location = New-Object System.Drawing.Point(110, $currentY)
    $txtTitle.Size = New-Object System.Drawing.Size(250, 25)
    $txtTitle.Font = $fontLabel
    $form.Controls.Add($txtTitle)
    $currentY += 40

    Add-Label '補足 :' $currentY
    $txtDescription = New-Object System.Windows.Forms.TextBox
    $txtDescription.Location = New-Object System.Drawing.Point(110, $currentY)
    $txtDescription.Size = New-Object System.Drawing.Size(250, 70)
    $txtDescription.Font = $fontLabel
    $txtDescription.Multiline = $true
    $txtDescription.AcceptsReturn = $true
    $txtDescription.ScrollBars = [System.Windows.Forms.ScrollBars]::Vertical
    $txtDescription.WordWrap = $true
    $form.Controls.Add($txtDescription)
    $currentY += 80

    Add-Label '診療科 :' $currentY
    $cmbKa = New-Object System.Windows.Forms.ComboBox
    $cmbKa.Location = New-Object System.Drawing.Point(110, $currentY)
    $cmbKa.Size = New-Object System.Drawing.Size(250, 25)
    $cmbKa.Items.AddRange(@(''; $kaList))
    $cmbKa.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
    $cmbKa.Font = $fontLabel
    $form.Controls.Add($cmbKa)
    $currentY += 40

    Add-Label '入外 :' $currentY
    $cmbNyuugai = New-Object System.Windows.Forms.ComboBox
    $cmbNyuugai.Location = New-Object System.Drawing.Point(110, $currentY)
    $cmbNyuugai.Size = New-Object System.Drawing.Size(250, 25)
    $cmbNyuugai.Items.AddRange(@('', '入院', '外来'))
    $cmbNyuugai.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
    $cmbNyuugai.Font = $fontLabel
    $form.Controls.Add($cmbNyuugai)
    $currentY += 40

    Add-Label '日付 :' $currentY
    $dtpDate = New-Object System.Windows.Forms.DateTimePicker
    $dtpDate.Location = New-Object System.Drawing.Point(110, $currentY)
    $dtpDate.Size = New-Object System.Drawing.Size(250, 25)
    $dtpDate.Format = [System.Windows.Forms.DateTimePickerFormat]::Custom
    $dtpDate.CustomFormat = 'yyyy/MM/dd'
    $dtpDate.Font = $fontLabel
    $form.Controls.Add($dtpDate)
    $currentY += 40

    Add-Label '医師コード :' $currentY
    $txtIshiCd = New-Object System.Windows.Forms.TextBox
    $txtIshiCd.Location = New-Object System.Drawing.Point(110, $currentY)
    $txtIshiCd.Size = New-Object System.Drawing.Size(250, 25)
    $txtIshiCd.Font = $fontLabel
    $txtIshiCd.ImeMode = [System.Windows.Forms.ImeMode]::Disable
    $form.Controls.Add($txtIshiCd)
    $currentY += 40

    Add-Label '自動スキャン :' $currentY
    $chkAuto = New-Object System.Windows.Forms.CheckBox
    $chkAuto.Location = New-Object System.Drawing.Point(130, ($currentY + 3))
    $chkAuto.Size = New-Object System.Drawing.Size(250, 25)
    $chkAuto.Font = $fontLabel
    $chkAuto.Text = '自動で1枚文書を登録します'
    $form.Controls.Add($chkAuto)
    $currentY += 50

    # ボタン
    $btnScan = New-Object System.Windows.Forms.Button
    $btnScan.Text = 'スキャン'
    $btnScan.Location = New-Object System.Drawing.Point(70, $currentY)
    $btnScan.Size = New-Object System.Drawing.Size(110, 40)
    $form.Controls.Add($btnScan)

    $btnClose = New-Object System.Windows.Forms.Button
    $btnClose.Text = '閉じる'
    $btnClose.Location = New-Object System.Drawing.Point(230, $currentY)
    $btnClose.Size = New-Object System.Drawing.Size(110, 40)
    $form.Controls.Add($btnClose)

    # ボタンロジック
    $submitAction = {
        # 値のvalidation
        if (-not $txtId.Text.Trim()) {
            [System.Windows.Forms.MessageBox]::Show('IDは必須入力です', 'エラー')
            return
        } elseif ($txtId.Text.Trim() -notmatch '^(0?\d{2,7}|9\d{7})$') {
            [System.Windows.Forms.MessageBox]::Show('IDが明らかに不正です', 'エラー')
            return
        }

        if (-not $txtTitle.Text.Trim()) {
            [System.Windows.Forms.MessageBox]::Show('名称は必須入力です', 'エラー')
            return
        }

        if ($txtIshiCd.Text.Trim()) {
            if ($txtIshiCd.Text.Trim() -notmatch '^\d+$') {
                [System.Windows.Forms.MessageBox]::Show('医師コードは数値のみです', 'エラー')
                return
            }
        }

        try {
            # 設定画面を取得して設定
            $setteiWindow = Get-SetteiWindow -AppWindow $script:appWindow
            [void](Set-SetteiWindow -setteiWindow $setteiWindow -kanjyaId $txtId.Text.Trim() -kaCode $cmbKa.SelectedItem -nyuugai $cmbNyuugai.SelectedItem -ishiCode $txtIshiCd.Text.Trim())

            # 新規画像へ推移
            $shinkiButton = Get-UIAButton -Parent $appWindow -id 'btnNewPic'
            [void](Invoke-UIAElement $shinkiButton)
            [UIATools]::Sleep()

            $shinkiWindow = Get-UIAWindow -Parent $appWindow -id 'frmSave'
            if ( $null -eq $shinkiWindow) {
                throw '新規画像登録ウインドウが開いていません'
            }

            # 新規画像の設定を入力
            [void](Set-NewImageWindowContent -NewImageWindow $shinkiWindow -DocumentName $txtTitle.Text -Description $txtDescription.Text -Date $dtpDate.Value.ToString('yyyy/MM/dd'))
            # スキャンボタンを操作してスキャンを実施
            [void](Register-ScanImage -NewImageWindow $shinkiWindow -Auto $chkAuto.Checked)

            $txtId.Clear()

        } catch {
            Write-Error $_
            $form.Activate()
            [System.Windows.Forms.MessageBox]::Show($form, $_.Exception.Message, 'エラー')
        }

        # 次のループへ
        $form.Activate()
        $txtId.Focus()
    }

    # イベント登録
    $txtId.Add_KeyDown({
        param($sender, $event)
        if ($event.KeyCode -eq [System.Windows.Forms.Keys]::Enter) {
            $event.SuppressKeyPress = $true
            $btnScan.Focus()
        }
    })
    
    $btnScan.Add_Click({ &$submitAction })
    $btnClose.Add_Click({ $form.Close() })
    $form.Add_Shown({
        $form.Activate()
        $txtId.Focus()
    })

    # フォームを開く
    $form.ShowDialog() | Out-Null
}

###################################################################### メインルーチン
try {
    # メインウインドウの存在確認
    $appWindow = GetAppWindow -Name '患者別文書管理*' -ExecutablePath 'c:\ssi\exe\srvDocumentManager.exe'
    if ($null -eq $appWindow) {
        throw '文書管理が確認できませんでした'
    }

    # 設定ウインドウから診療科コードを取得するためのキャッシュ変数
    $kaCodeList = $null

    # 設定画面を取得する
    $setteiWindow = Get-SetteiWindow -AppWindow $appWindow
    if ($null -eq $setteiWindow) {
        throw '文書管理の状態が異常です'
    }

    # 診療科コードを取得
    try {
        $kaCodeList = Get-UIAComboBoxItems (Get-UIAComboBox -Parent $setteiWindow -Id 'CboSnk')
    } catch {
        throw '診療科コード一覧の取得に失敗しました'
    }

    # ダイアログの操作へうつる
    Show-ControlForm $kaCodeList

    # 文書管理を閉じる
    try {
        $btnClose = Get-UIAButton -Parent $appWindow -Id 'btnClose'
        if ($null -ne $btnClose) {
            Invoke-UIAElement $btnClose
        }
    } catch {}
} catch {
    Write-Error $_
    [System.Windows.Forms.MessageBox]::Show($_.Exception.Message, '異常終了')
    exit 1
}

