#
# 文書管理に画像を登録する
#
# -- パラメーター --
# Id: 患者ID (必須)
# Path: 登録する画像のファイルフルパス (必須)
# Date: 文書の日付 (任意、例: 2024/12/31)
# Title: 文書名 (任意)
# Description: 補足 (任意)
# IshiCode: 医師コード (任意)
# KaCode: 診療科コード (任意)
# Nyuugai: 入外区分 (任意、入院なら「入」、外来なら「外」など)
[CmdletBinding()]
Param (
    [Parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true)]
    [string]$Id,
    [Parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true)]
    [string]$Path,
    [Parameter(ValueFromPipelineByPropertyName = $true)]
    [string]$Date,
    [Parameter(ValueFromPipelineByPropertyName = $true)]
    [string]$Title,
    [Parameter(ValueFromPipelineByPropertyName = $true)]
    [string]$Description,
    [Parameter(ValueFromPipelineByPropertyName = $true)]
    [string]$IshiCode,
    [Parameter(ValueFromPipelineByPropertyName = $true)]
    [string]$KaCode,
    [Parameter(ValueFromPipelineByPropertyName = $true)]
    [string]$Nyuugai
)

begin {
    # 追加アセンブリのロード
    # Add-Type -AssemblyName PresentationFramework
    # Add-Type -AssemblyName Microsoft.VisualBasic
    # Add-Type -AssemblyName Windows.Forms

    . "$PSScriptRoot\automation-core.ps1"

    ######################################################################
    # サブルーチン
    # 必要に応じて文書リスト画面などから設定画面にfallbackする
    function Get-SetteiWindow {
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
                $childWindow = Get-UIAControl -Parent $AppWindow -Type ([Windows.Automation.ControlType]::Window) -Scope ([Windows.Automation.TreeScope]::Children) -TimeoutSec 1
            }
            catch {}

            if ( $null -eq $childWindow ) {
                # write-host 'Fallback from list view'
                # 子ウインドウが開いていないのでリスト画面
                # リスト画面からのfallback
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
                # write-host 'Fallback from window'
                $windowTitle = $childWindow.Current.Name
                # write-host "Check window title `"$($windowTitle)`""
                # 設定画面なので処理は不要でウインドウを返す
                if ( $windowTitle -eq '設定') {
                    return $childWindow
                }

                # 当該ウインドウを閉じて戻る～スキャナのコントロールなどで深かったらエラーになる
                $button = Get-UIAButton -Parent $childWindow -Name '*閉じる*'
                if ( $null -eq $button ) {
                    $button = Get-UIAButton -Parent $childWindow -Name '*Close*'
                    if ( $null -eq $button ) {
                        throw '子ウインドウに閉じるボタンがありません'
                    }
                }
                Invoke-UIAElement $button
            }
            $watchdog--
        } while ( $watchdog -gt 0 )
        return $null
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
        if ($null -eq $kanjyaId -or $kanjyaId -match '^\s*$' -or $kanjyaId -notmatch '^0?\d{2,7}$') {
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
            [parameter(Mandatory)][Windows.Automation.AutomationElement]$Element,
            [string]$DocumentName,
            [string]$Description,
            [string]$Date
        )
        if ($null -eq $Element) {
            throw 'Windowが指定されていません'
        }

        # 文書名
        if ($DocumentName.Trim()) {
            # Write-Host '名称の入力'
            $titleCmb = Get-UIAComboBox -Parent $Element -Id 'txtFile'
            Set-UIAComboBoxValue $titleCmb $DocumentName.Trim()
        }

        # 補足
        if ($Description.Trim()) {
            # Write-Host '補足の入力'
            $descriptionEdit = Get-UIAControl -Parent $Element -Id 'txtBikou'

            # SetValue非対応になるので文字列を送信する
            $descriptionEdit.SetFocus()
            [UIATools]::Sleep(50)

            [System.Windows.Forms.SendKeys]::SendWait($Description.Trim())
        }

        # カレンダー ThreeCalendarDialogを操作
        if ($Date -match '^20\d\d[/-]((0?[13578]|1[02])[/-](0?[1-9]|[12][0-9]|3[01])|(0?2[/-](0?[1-9]|1[0-9]|2[0-9]))|((0?[469]|11)[/-](0?[1-9]|[12][0-9]|30)))$') {
            # Write-Host '日付の設定'
            # 日付を分解
            $targetDateStr = $Date.Split('/-')
            $targetInMonth = [int]$targetDateStr[0] * 12 + [int]$targetDateStr[1]
            $targetDay = [int]$targetDateStr[2]

            # 文書管理のカレンダーを開く(力業しかない)
            $calenderButtonPane = Get-UIAPane -Parent (Get-UIAPane -Parent $Element -Id 'pnlDate')
            [UIATools]::ForceClick($calenderButtonPane)

            [UIATools]::Sleep()
            
            # カレンダーダイアログを取得
            $calendarWindow = Get-UIAWindow -Parent $Element -Id 'ThreeCalendarDialog'
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

    function Register-NewImage {
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

            # ファイル選択ダイアログを操作してディレクトリを指定する
            $fileselectDialog = Get-UIAWindow -Parent $subWindow
            if ($null -eq $fileselectDialog) {
                throw 'ファイル選択ダイアログが開きません.'
            }
            else {
                Set-UIAValue (Get-UIAComboBox -Parent $fileselectDialog -Id '1148') $Path
                Invoke-UIAElement (Get-UIAButton -Parent $fileselectDialog -id '1')
                [UIATools]::Sleep(500)
            }

            # ファイル選択Paneを操作してファイルを探して指定する
            # tlpPictures - Panes(区切りの数だけPaneがある) - Pane(プレビュー), Text(Name = ファイル名)
            $picturePane = Get-UIAPane -Parent $subWindow -Id 'tlpPictures'
            $nextBtn = Get-UIAButton -Parent $subWindow -Id 'btnSmallIncrement'

            $checkBox = $null
            # Write-Host "Looking for $($filename)"
            do {
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
            Invoke-UIAElement $checkBox

            # 選択確定
            Invoke-UIAElement (Get-UIAButton -Parent $subWindow -Id 'btnRegist')
        }

        # 登録
        Invoke-UIAElement (Get-UIAButton -Parent $Element -Id 'cmdSave')
        [UIATools]::Sleep(500)

        # 確認ウインドウのボタンを操作
        Invoke-UIAElement (Get-UIAButton -Parent (Get-UIAWindow -Parent $Element -Id 'SSIMessageDialog') -Id 'Button1')
        [UIATools]::Sleep(1000)
    }

    # メインウインドウの存在確認
    # write-host 'アプリケーションを確認'
    $appWindow = GetAppWindow -Name '患者別文書管理*' -ExecutablePath 'c:\ssi\exe\srvDocumentManager.exe'
    if ($null -eq $appWindow) {
        throw '文書管理が確認できませんでした'
    }

    # 設定ウインドウから診療科コードを取得するためのキャッシュ変数
    $kaCodeList = $null

    # 致命的でない例外が生じた場合のエラー保存用配列
    $errorList = @()
}

process {
    # 必須パラメーター確認
    try {
        if ($Id.Trim() -eq '') {
            throw '患者IDは必須です'
        }
        if ($Id -notmatch '^0?\d{2,7}$') {
            throw '患者IDの形式が正しくありません'
        }

        if ($Path.Trim() -eq '') {
            throw 'ファイルパスは必須です'
        }
        if (-not (Test-Path -PathType Leaf -Path $Path)) {
            throw '無効なファイル指定です'
        }
        if ($Path.Trim() -notlike '*.jpg' -and $Path.Trim() -notlike '*.pdf') {
            throw 'ファイルはjpgかpdfのみ対応しています'
        }
    } catch {
        Write-Debug $_
        $errorList += "登録失敗 患者ID: ${Id} ファイル名: ${Path} - $_"
        return
    }

    # 設定画面を取得する
    # write-host '設定ウインドウへ遷移'
    $setteiWindow = Get-SetteiWindow -AppWindow $appWindow
    if ($null -eq $setteiWindow) {
        throw '文書管理の状態が異常です'
    }

    # 診療科コードを取得してリストから選択肢にあわせる
    try {
        if ($null -eq $kaCodeList) {
            $kaCodeList = Get-UIAComboBoxItems (Get-UIAComboBox -Parent $setteiWindow -Id 'CboSnk')
        }

        if ($kaCode.Trim()) {
            $kaCode = $kaCodeList | Where-Object { $_ -like $kaCode } | Select-Object -First 1
        }
        else {
            $kaCode = ''
        }
    } catch {
        Write-Error $_
        throw '診療科コードの取得に失敗しました'
    }

    try {
        Set-SetteiWindow -setteiWindow $setteiWindow -kanjyaId $Id -kaCode $kaCode -ishiCode $IshiCode -nyuugai $Nyuugai
    } catch {
        Write-Debug $_
        $errorList += "登録失敗 患者ID: ${Id} ファイル名: ${Path} - $_"
        return
    }

    # リスト画面での操作新規画像登録へ
    # 新規画像を登録へ
    # 新規画像登録画面に遷移
    $shinkiButton = Get-UIAButton -Parent $appWindow -id 'btnNewPic'
    Invoke-UIAElement $shinkiButton

    [UIATools]::Sleep(300)

    $shinkiWindow = Get-UIAWindow -Parent $appWindow -id 'frmSave'
    if ( $null -eq $shinkiWindow) {
        throw '新規画像登録ウインドウが開いていません'
    }

    # 新規画像諸元入力
    try {
        Set-NewImageWindowContent -Element $shinkiWindow -DocumentName $Title -Description $Description -Date $Date
        # 画像を選択して保存
        Register-NewImage -Element $shinkiWindow -Path $Path
    } catch {
        Write-Debug $_
        $errorList += "登録失敗 患者ID: ${Id} ファイル名: ${Path} - $_"
        return
    }
}

end {
    if ($errorList.Count -gt 0) {
        Write-Error "登録に失敗した項目が ${errorList.Count} 件あります。詳細は以下の通りです。"
        $errorList | ForEach-Object { Write-Error $_ }
    }
}
