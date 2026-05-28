# 文書管理の操作をするツールキット
if (-not ([System.Management.Automation.PSTypeName]'UIATools').Type) {
    throw 'Tools-AutomationCode.ps1を先にロードしてください'
}

# 文書管理を起動する
function Get-DocumentManager {
    $appWindow = GetAppWindow -Name '患者別文書管理*' -ExecutablePath 'c:\ssi\exe\srvDocumentManager.exe'
    if ($null -eq $appWindow) {
        throw '文書管理が正常に起動できませんでした'
    }
    return $appWindow
}

# 文書管理を終了する
function Close-DocumentManager {
    param(
        [Parameter(Mandatory)][Windows.Automation.AutomationElement]$appWindow
    )
    if ($null -eq $appWindow) {
        throw '文書管理が指定されていません'
    }

    $setteiWindow = Get-DMSetteiWindow $appWindow

    # 閉じるボタンを操作
    $btnClose = Get-UIAButton -Parent $setteiWindow -Id 'Cmd_001'
    Invoke-UIAElement $btnClose
    [UIATools]::Sleep()

    # メニュー操作で閉じる
    $menubar = Get-UIAMenuBar -Parent $AppWindow -Name ''
    $menuitem = Get-UIAMenuItem -Parent $menubar -Name '終了(X)'
    if ($null -eq $menuitem ) {
        throw 'メニュー取得(終了)に問題が発生しました'
    }
    Invoke-UIAElement $menuitem
}

# 必要に応じて文書リスト画面などから設定画面にfallbackして設定ウインドウを取得する
function Get-DMSetteiWindow {
    param(
        [Parameter(Mandatory)][Windows.Automation.AutomationElement]$appWindow
    )
    if ( $null -eq $AppWindow ) {
        return $null
    }

    # 子ウインドウを取得
    $watchDog = 10
    do {
        $childWindow = $null
        try {
            [UIATools]::Sleep()
            $childWindow = Get-UIAControl -Parent $AppWindow -Type ([Windows.Automation.ControlType]::Window) -Scope ([Windows.Automation.TreeScope]::Children) -TimeoutSec 1
        }
        catch {}

        if ( $null -eq $childWindow ) {
            # 子ウインドウが開いていないのでリスト画面かまっさらなウインドウ
            # fallbackはメニュー操作
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
                $button = $buttons[-1]
                Invoke-UIAElement $button
            }
        }
        $watchdog--
    } while ( $watchdog -gt 0 )
    throw '子ウインドウに閉じるボタンがありません'
}

# 設定画面の内容を入力して次に進む
function Set-DMSetteiWindow {
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

# 新規文書作成 - 指定された文書名のexcelを開くまでの操作をする
function New-DMDocument {
    param(
        [Parameter(Mandatory)][Windows.Automation.AutomationElement]$appWindow,
        [Parameter(Mandatory)][string]$documentName
    )

    if ($null -eq $appWindow) {
        throw '正しいアプリケーションウインドウが指定されていません' 
    }

    if (-not $documentName.Trim()) {
        throw '文書名が指定されていません'
    }

    # 新規文書を起動
    $shinkiDocBtn = Get-UIAButton -Parent $appWindow -Id 'btnNewDoc'
    Invoke-UIAElement $shinkiDocBtn

    [UIATools]::Sleep(300)

    # 新規文書ウインドウを取得して操作
    $shinkiDocWindow = Get-UIAWindow -Parent $appWindow -Id 'frmNewDocs'
    if ($null -eq $shinkiDocWindow) {
        throw '文書管理の状態が異常です'
    }

    # 検索
    $shinkiDocEdit = Get-UIAEdit -Parent $shinkiDocWindow -Id 'txtSearch'
    Set-UIAValue -Element $shinkiDocEdit -Value $documentName
    Set-UIAValue -Element $shinkiDocEdit -Value '~' -OmitEscape -Force # Enterキー
    [UIATools]::Sleep(300)

    # 検索結果を取得して最初の候補を選択
    $shinkiKensakuLst = Get-UIAList -Parent $shinkiDocWindow -Id 'lstSearch'
    if ($null -eq $shinkiKensakuLst) {
        throw "${$documentName}に該当する文書がありません"
    }
    $shinkiKensakuItem = (Get-UIAListItems -Parent $shinkiKensakuLst)[0]
    Invoke-UIAForceClick -ParentWindow $shinkiDocWindow -Element $shinkiKensakuItem # 強制クリック

    # Excelの起動を待つ
    [UIATools]::Sleep(3500)
}

# 新規画像ウインドウの内容を設定
function Set-DMNewImageWindowProperties {
    param(
        [parameter(Mandatory)][Windows.Automation.AutomationElement]$Element,
        [string]$DocumentName,
        [string]$Description,
        [string]$Date
    )
    if ($null -eq $Element) {
        throw 'Windowが指定されていません'
    }

    # 新規画像ウインドウであることを確認
    $windowName = Get-UIAName $Element
    $shinkiWindow = $null

    # 新規画像ウインドウでなければリスト画面から遷移する
    if ($windowName -ne '新規画像') {
        if ($windowName -like '患者別文書管理*') {
            # リスト画面なので新規画像ウインドウに遷移する
            $shinkiButton = Get-UIAButton -Parent $Element -id 'btnNewPic'
            if ($null -eq $shinkiButton) {
                throw 'アプリケーションウインドウが不正です'
            }
            [void](Invoke-UIAElement $shinkiButton)
            [UIATools]::Sleep(300)

            # 新規画像ウインドウを取得
            $shinkiWindow = Get-UIAWindow -Parent $Element -id 'frmSave'
            if ( $null -eq $shinkiWindow) {
                throw '新規画像登録ウインドウが確認できません'
            }
        } else {
            throw 'アプリケーションウインドウが不正です'
        }
    } else {
        $shinkiWindow = $Element
    }

    # フォーカス退避用に戻るボタンを取得しておく
    $cmdBackBtn = Get-UIAButton -Parent $shinkiWindow -Id 'cmdBack'

    # 文書名
    if ($DocumentName.Trim()) {
        # Write-Host '名称の入力'
        $titleCmb = Get-UIAComboBox -Parent $shinkiWindow -Id 'txtFile'

        # SetValueなど一切をうけつけないコントロールでIMEもコントロールできないので文字列をクリップボード経由で貼り付ける
        Set-Clipboard $DocumentName
        Set-UIAValue -Element $titleCmb -Value '^a{DEL}^v' -Force -OmitEscape
        $cmdBackBtn.SetFocus()
    }

    # 補足
    if ($Description.Trim()) {
        # Write-Host '補足の入力'
        $descriptionEdit = Get-UIAControl -Parent $shinkiWindow -Id 'txtBikou'

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
        $calenderButtonPane = Get-UIAPane -Parent (Get-UIAPane -Parent $shinkiWindow -Id 'pnlDate')

        # 現在の日付を取得して変更が必要ならばカレンダーを操作する
        $calenderDateStr = ($calenderButtonPane.Current.Name).Split('年月日')
        if ([int]$targetDateStr[0] -ne [int]$calenderDateStr[0] -or [int]$targetDateStr[1] -ne [int]$calenderDateStr[1] -or [int]$targetDateStr[2] -ne [int]$calenderDateStr[2]) {
            [UIATools]::ForceClick($calenderButtonPane)
            [UIATools]::Sleep()
            
            # カレンダーダイアログを取得
            $calendarWindow = Get-UIAWindow -Parent $shinkiWindow -Id 'ThreeCalendarDialog'
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

    # 他のルーチンで使うので新規画像ウインドウを返す
    return $shinkiWindow
}

# 画像を選択して登録する
function Register-DMNewImage {
    param(
        [parameter(Mandatory)][Windows.Automation.AutomationElement]$Element,
        [string]$Path
    )
    if ($null -eq $Element) {
        throw 'ウインドウが指定されていません'
    }

    # パスを確認
    if (-not (Test-Path -PathType Leaf -Path $Path)) {
        throw '無効なファイル指定です'
    }
    $filename = [System.IO.Path]::GetFileName($Path)

    # 画像取り込みボタンを押す
    $cmdSelectBtn = Get-UIAButton -Parent $shinkiWindow -Id 'cmdSelect'
    if ($null -eq $cmdSelectBtn) {
        throw 'ウインドウが異なるようです'
    }
    Invoke-UIAElement $cmdSelectBtn

    [UIATools]::Sleep(300)

    # 画像選択を開始
    # ディレクトリを指定する-メニュー操作
    $subWindow = Get-UIAWindow -Parent $Element
    if ($null -eq $subWindow) {
        throw '画像選択ウインドウが開きません.'
    }
    else {
        $menubar = Get-UIAMenuBar -Parent $subWindow -Id 'mnuMain'
        $menuitem = Get-UIAMenuItem -Parent $menubar -Name 'ファイル(F)'
        $menuitem = Get-UIAMenuItem -Parent $menuitem -Name '開く'
        if ($null -eq $menuitem) {
            throw '画像選択ウインドウに異常があります'
        }
        Invoke-UIAElement $menuitem
        [UIATools]::Sleep(500)

        # ファイル選択ダイアログを操作してディレクトリを指定する（パスをいれてフォルダが選ばれる謎仕様）
        $fileselectDialog = Get-UIAWindow -Parent $subWindow
        if ($null -eq $fileselectDialog) {
            throw 'ファイル選択ダイアログが開きません.'
        }
        else {
            Set-UIAValue (Get-UIAComboBox -Parent $fileselectDialog -Id '1148') $Path
            Invoke-UIAElement (Get-UIAButton -Parent $fileselectDialog -id '1')
            [UIATools]::Sleep(300)
        }

        # ファイルの選択
        # ファイル選択Paneを操作してファイルを探して指定する
        # tlpPictures - Panes(区切りの数だけPaneがある) - Pane(プレビュー), Text(Name = ファイル名)
        # 次へボタンで送ってゆく
        $picturePane = Get-UIAPane -Parent $subWindow -Id 'tlpPictures'
        $nextBtn = Get-UIAButton -Parent $subWindow -Id 'btnSmallIncrement'

        $checkBox = $null
        # Write-Host "Looking for $($filename)"
        do {
            # 画面タイルに表示された画像のキャプション（ファイル名）を取得してチェック
            $listPanes = Get-UIAControls -Parent $picturePane -Scope ([Windows.Automation.TreeScope]::Children) -Type ([Windows.Automation.ControlType]::Pane)
            if ($listPanes.Count -eq 0) {
                break
            }
            $listPanes | ForEach-Object {
                $paneLabel = Get-UIAText -Parent $_
                # Write-Host "Label is $($paneLabel.Current.Name)"
                if ($paneLabel.Current.Name -eq $filename) {
                    $checkBox = Get-UIACheckBox -Parent $_
                    break
                }
            }
            # 次へボタンが有効ならばまだ画像がある
            if ($nextBtn.Current.IsEnabled -eq $true) {
                Invoke-UIAElement $nextBtn
            }
            else {
                break;
            }
        } while ($true)

        if ($null -eq $checkBox) {
            throw 'ファイルが見つかりませんでした.'
        }
        # 対象の画像を選択
        Invoke-UIAElement $checkBox

        # 選択確定
        Invoke-UIAElement (Get-UIAButton -Parent $subWindow -Id 'btnRegist')
    }

    # 登録
    Invoke-UIAElement (Get-UIAButton -Parent $Element -Id 'cmdSave')
    [UIATools]::Sleep()

    # 確認ウインドウのボタンを操作
    Invoke-UIAElement (Get-UIAButton -Parent (Get-UIAWindow -Parent $Element -Id 'SSIMessageDialog') -Id 'Button1')
    [UIATools]::Sleep(300)
}

# スキャンして画像を登録する
function Register-DMScanImage {
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
#function Set-EPSONScan {
#    param (
#        [parameter(Mandatory)][Windows.Automation.AutomationElement]$ScanWindow
#    )
#
#
#}

# スキャナ設定画面の操作 (EPSON Scan2)
#function Set-EPSONScan {
#    param (
#        [parameter(Mandatory)][Windows.Automation.AutomationElement]$ScanWindow
#    )
#}
