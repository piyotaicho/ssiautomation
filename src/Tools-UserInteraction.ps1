# GUIベースでユーザーの入力を促すユーティリティ
# Last update: 2026-06-13 p
Add-Type -AssemblyName Microsoft.VisualBasic
Add-Type -AssemblyName System.Windows.Forms

# ユーザーから文字列の入力を受ける
function Confirm-UserInput {
    param(
        [string]$text,
        [string]$title = 'ユーザー入力'
    )
    $form = New-Object System.Windows.Forms.Form
    $form.TopMost = $true
    [Microsoft.VisualBasic.Interaction]::InputBox($form, $text, $title, '')
    $form.Dispose()
}

# YesNoの選択 Yesで$true
function Confirm-UserYesNo {
    param(
        [string]$text,
        [string]$title = 'ユーザー確認'
    )
    $form = New-Object System.Windows.Forms.Form
    $form.TopMost = $true
    $result = [System.Windows.Forms.MessageBox]::Show($form, $text, $title, 'YesNo', 'Question')
    $form.Dispose()
    return ($result -eq 'Yes')
}

# YesCancelの選択 Yesで$true
function Confirm-UserOkCancel {
    param(
        [string]$text,
        [string]$title = 'ユーザー確認'
    )
    $form = New-Object System.Windows.Forms.Form
    $form.TopMost = $true
    $result = [System.Windows.Forms.MessageBox]::Show($form, $text, $title, 'OkCancel', 'Question')
    $form.Dispose()
    return ($result -eq 'Ok')
}

# 注意ダイアログを表示
function Invoke-InformationDialog {
    param(
        [string]$text,
        [string]$title = '情報'
    )
    $form = New-Object System.Windows.Forms.Form
    $form.TopMost = $true
    [System.Windows.Forms.MessageBox]::Show($form, $text, $title, 'Ok', 'Information') | Out-Null
    $form.Dispose()
}

# 注意ダイアログを表示
function Invoke-AlertDialog {
    param(
        [string]$text,
        [string]$title = '情報'
    )
    $form = New-Object System.Windows.Forms.Form
    $form.TopMost = $true
    [System.Windows.Forms.MessageBox]::Show($form, $text, $title, 'Ok', 'Warning') | Out-Null
    $form.Dispose()
}

# エラーダイアログを表示
function Invoke-ErrorDialog {
    param(
        [string]$text,
        [string]$title = '情報'
    )
    $form = New-Object System.Windows.Forms.Form
    $form.TopMost = $true
    [System.Windows.Forms.MessageBox]::Show($form, $text, $title, 'Ok', 'Error') | Out-Null
    $form.Dispose()
}

# ファイル選択ダイアログを表示
function Invoke-FileSelectDialog {
    param(
        [string]$title = 'ファイルを選択してください',
        [string]$filter = 'すべてのファイル|*.*',
        [string]$initialDirectory
    )

    # パラメーターの確認
    $filters = $filter -split '|'
    if (($filters.Count % 2) -ne 0) {
        throw "指定されたフィルタ $filter が不正です"
    }

    if ($initialDirectory -and (-not (Test-Path -PathType Container -Path $initialDirectory))) {
        throw "指定された開始フォルダ $initialDirectory が不正です"
    }

    $fileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $form = New-Object System.Windows.Forms.Form
    $form.TopMost = $true

    $fileDialog.Title = $title
    $fileDialog.Filter = $filter
    if ($initialDirectory) {
        $fileDialog.InitialDirectory = $initialDirectory
    }

    $resultPath = $null
    if ($fileDialog.ShowDialog($form) -eq [System.Windows.Forms.DialogResult]::OK) {
        $resultPath = $fileDialog.FileName
    }
    $form.Dispose()
    $fileDialog.Dispose()

    return $resultPath
}

# デフォルトアクションはなし
