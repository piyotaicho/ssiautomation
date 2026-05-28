# GUIベースでユーザーの入力を促すユーティリティ
# Last update: 2026-05-27 p
Add-Type -AssemblyName Microsoft.VisualBasic

# ユーザーから文字列の入力を受ける
function Confirm-UserInput {
    param(
        [string]$text,
        [string]$title = 'ユーザー入力'
    )
    [Microsoft.VisualBasic.Interaction]::InputBox($text, $title, '')
}

# YesNoの選択 Yesで$true
function Confirm-UserYesNo {
    param(
        [string]$text,
        [string]$title = 'ユーザー確認'
    )
    $result = [System.Windows.Forms.MessageBox]::Show($text, $title, 'YesNo', 'Question')
    return ($result -eq 'Yes')
}

# YesCancelの選択 Yesで$true
function Confirm-UserOkCancel {
    param(
        [string]$text,
        [string]$title = 'ユーザー確認'
    )
    $result = [System.Windows.Forms.MessageBox]::Show($text, $title, 'OkCancel', 'Question')
    return ($result -eq 'Ok')
}

# 注意ダイアログを表示
function Invoke-InformationDialog {
    param(
        [string]$text,
        [string]$title = '情報'
    )
    [System.Windows.Forms.MessageBox]::Show($text, $title, 'Ok', 'Information') | Out-Null
}

# 注意ダイアログを表示
function Invoke-AlertDialog {
    param(
        [string]$text,
        [string]$title = '情報'
    )
    [System.Windows.Forms.MessageBox]::Show($text, $title, 'Ok', 'Warning') | Out-Null
}

# エラーダイアログを表示
function Invoke-ErrorDialog {
    param(
        [string]$text,
        [string]$title = '情報'
    )
    [System.Windows.Forms.MessageBox]::Show($text, $title, 'Ok', 'Error') | Out-Null
}

# デフォルトアクションはなし
