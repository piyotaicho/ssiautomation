# スキャン基本設定のフォームを表示する
#
# 診療科・入院外来・文書種別
# 文書の回転 そのまま・右90度・左90度・180度
#
# スキャン解像度は300dpi固定 一応グレースケールも許容
# スキャンサイズは A4 固定
Add-Type -AssemblyName System.Windows.Forms

# 定数リスト
$NyuuinGairaiList = @( "入院", "外来")
$ShinryoukaList = @()

# Formのコントロール定義
$FormInitialLoading = New-Object -TypeName System.Windows.Forms.Form
[System.Windows.Forms.Label]$LabelInitialLoading = $null
$script:IsInitialLoadingClosed = $false

$FormSetting = New-Object -TypeName System.Windows.Forms.Form
[System.Windows.Forms.Label]$Label1 = $null
[System.Windows.Forms.Label]$Label2 = $null
[System.Windows.Forms.Label]$Label3 = $null
[System.Windows.Forms.Label]$Label4 = $null
[System.Windows.Forms.Label]$Label5 = $null
[System.Windows.Forms.Label]$Label6 = $null
[System.Windows.Forms.ComboBox]$ComboShinryoka = $null
[System.Windows.Forms.ComboBox]$ComboNyuuinGairai = $null
[System.Windows.Forms.ComboBox]$ComboDocumentType = $null
[System.Windows.Forms.CheckBox]$CheckGrayscale = $null
[System.Windows.Forms.RadioButton]$RadioNormal = $null
[System.Windows.Forms.RadioButton]$RadioRotate180 = $null
[System.Windows.Forms.RadioButton]$RadioCCW90 = $null
[System.Windows.Forms.RadioButton]$RadioCW90 = $null
[System.Windows.Forms.TextBox]$TextID = $null
[System.Windows.Forms.Button]$ButtonProcess = $null

function ShowInitialLoading {
    $script:IsInitialLoadingClosed = $false
    $LabelInitialLoading = New-Object -TypeName System.Windows.Forms.Label -Property @{
        Font     = "Meiryo UI, 11.25pt"
        Location = "12,9"
        Size     = "380,23"
        Text     = "起動"
    }
    #
    #FormInitialLoading
    #
    $FormInitialLoading.SuspendLayout()
    $FormInitialLoading.ClientSize = '404,50'
    $FormInitialLoading.Controls.Add($LabelInitialLoading)
    $FormInitialLoading.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
    $FormInitialLoading.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen
    $FormInitialLoading.Text = '連続スキャン設定'
    $FormInitialLoading.add_FormClosed({
            $script:IsInitialLoadingClosed = $true
        })
    $FormInitialLoading.ResumeLayout($true)
    $FormInitialLoading.Show()
    [System.Windows.Forms.Application]::DoEvents()
}

function CloseInitialLoading {
    if ( $null -ne $FormInitialLoading -and -not $FormInitialLoading.IsDisposed ) {
        $FormInitialLoading.Close()
    }

    $script:IsInitialLoadingClosed = $true
}

function InitializeAppEnvilonWithLoading {
    # $initializeCommand = Get-Command -Name InitializeAppEnvilon -CommandType Function -ErrorAction SilentlyContinue
    #if ( $null -eq $initializeCommand ) {
    #    return
    #}

    ShowInitialLoading
    try {
        [System.Threading.Thread]::Sleep(1000) # 初期化処理の代わりに1秒待機
        $LabelInitialLoading.Text = "初期化処理中..."
        [System.Threading.Thread]::Sleep(5000) # 初期化処理の代わりに5秒待機
    }
    finally {
        if ( -not $script:IsInitialLoadingClosed ) {
            CloseInitialLoading
        }
    }
}

function InitializeSettingForm {
    $Label1 = New-Object -TypeName System.Windows.Forms.Label -Property @{
        Font     = "Meiryo UI, 11.25pt"
        Location = "12,9"
        Size     = "93,23"
        Text     = "診療科：" 
    }
    $ComboShinryoka = New-Object -TypeName System.Windows.Forms.ComboBox -Property @{
        Font          = "Meiryo UI, 11.25pt"
        Location      = "111,6"
        Size          = "165,27"
        DropDownStyle = "DropDownList"
    }
    $Label2 = New-Object -TypeName System.Windows.Forms.Label -Property @{
        Font     = "Meiryo UI, 11.25pt"
        Location = "12,42"
        Size     = "93,23"
        Text     = "入外："
    }
    $ComboNyuuinGairai = New-Object -TypeName System.Windows.Forms.ComboBox -Property @{
        Font          = "Meiryo UI, 11.25pt"
        Location      = "111,39"
        Size          = "121,27"
        DropDownStyle = "DropDownList"
    }
    $ComboNyuuinGairai.Items.AddRange($NyuuinGairaiList)
    $ComboNyuuinGairai.SelectedIndex = 0
    $Label3 = New-Object -TypeName System.Windows.Forms.Label -Property @{
        Font     = "Meiryo UI, 11.25pt"
        Location = "12,75"
        Size     = "93,23"
        Text     = "文書種別："
    }
    $ComboDocumentType = New-Object -TypeName System.Windows.Forms.ComboBox -Property @{
        Font     = "Meiryo UI, 11.25pt"
        Location = "111,72"
        Size     = "281,27"
        ImeMode  = "On" 
    }
    $CheckGrayscale = New-Object -TypeName System.Windows.Forms.CheckBox -Property @{
        Font     = "Meiryo UI, 11.25pt"
        Location = "111,106"
        Size     = "241,24"
        Text     = "カラー情報無し（グレー）"
    }
    $Label4 = New-Object -TypeName System.Windows.Forms.Label -Property @{
        Font     = "Meiryo UI, 11.25pt"
        Location = "12,138"
        Size     = "109,23"
        Text     = "文書の回転："
    }
    $RadioNormal = New-Object -TypeName System.Windows.Forms.RadioButton -Property @{
        Font     = "Meiryo UI, 11.25pt"
        Location = "111,136"
        Size     = "104,24"
        Text     = "なし"
        Checked  = $true
        TabStop  = $true
    }
    $RadioRotate180 = New-Object -TypeName System.Windows.Forms.RadioButton -Property @{
        Font     = "Meiryo UI, 11.25pt"
        Location = "248,136"
        Size     = "104,24"
        Text     = "180度"
    }
    $RadioCCW90 = New-Object -TypeName System.Windows.Forms.RadioButton -Property @{
        Font     = "Meiryo UI, 11.25pt"
        Location = "111,166"
        Size     = "104,24"
        Text     = "左90度"
    }
    $RadioCW90 = New-Object -TypeName System.Windows.Forms.RadioButton -Property @{
        Font     = "Meiryo UI, 11.25pt"
        Location = "248,166"
        Size     = "104,24"
        Text     = "右90度"
    }
    $Label5 = New-Object -TypeName System.Windows.Forms.Label -Property @{
        Location    = "12,202"
        Size        = "380,2"
        BorderStyle = "Fixed3D"
        Text        = ""
    }
    $Label6 = New-Object -TypeName System.Windows.Forms.Label -Property @{
        Font     = "Meiryo UI, 11.25pt"
        Location = "12,233"
        Size     = "93,23"
        Text     = "患者ID："
    }
    $TextID = New-Object -TypeName System.Windows.Forms.TextBox -Property @{
        Font     = "Meiryo UI, 11.25pt"
        Location = "111,230"
        Size     = "281,27"
        ImeMode  = "Disable"
    }
    $ButtonProcess = New-Object -TypeName System.Windows.Forms.Button -Property @{
        Font     = "Meiryo UI, 11.25pt"
        Location = "147,284"
        Size     = "110,35"
        Text     = "スキャン開始"
    }
    #
    #Form1
    #
    $FormSetting.SuspendLayout()
    $FormSetting.ClientSize = '404,330'
    $FormSetting.Controls.Add($ComboShinryoka)
    $FormSetting.Controls.Add($Label1)
    $FormSetting.Controls.Add($ComboNyuuinGairai)
    $FormSetting.Controls.Add($Label2)
    $FormSetting.Controls.Add($ComboDocumentType)
    $FormSetting.Controls.Add($Label3)
    $FormSetting.Controls.Add($CheckGrayscale)
    $FormSetting.Controls.Add($RadioNormal)
    $FormSetting.Controls.Add($RadioRotate180)
    $FormSetting.Controls.Add($RadioCCW90)
    $FormSetting.Controls.Add($RadioCW90)
    $FormSetting.Controls.Add($Label4)
    $FormSetting.Controls.Add($Label5)
    $FormSetting.Controls.Add($TextID)
    $FormSetting.Controls.Add($Label6)
    $FormSetting.Controls.Add($ButtonProcess)
    $FormSetting.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
    $FormSetting.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen
    $FormSetting.Text = '連続スキャン設定'
    $FormSetting.ResumeLayout($true)
}

. InitializeAppEnvilonWithLoading

. InitializeSettingForm

# 設定ウインドウ表示
if ( $null -ne $FormSetting ) {
    $FormSetting.ShowDialog() | Out-Null
}

