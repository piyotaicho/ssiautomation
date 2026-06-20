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
    . "$PSScriptRoot\Tools-AutomationCore.ps1"
    . "$PSScriptRoot\Tools-DocumentManager.ps1"

    ######################################################################
    # メインウインドウの存在確認
    $appWindow = Get-DocumentManager
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
    $fullPath = ''
    try {
        if ($Id.Trim() -eq '') {
            throw '患者IDは必須です'
        }
        if ($Id -notmatch '^(0?\d{2,7}|9\d{7})$') {
            throw '患者IDが有効なものではありません'
        }

        if ($Path.Trim() -eq '') {
            throw 'ファイルパスは必須です'
        }
        if (-not (Test-Path -PathType Leaf -Path $Path)) {
            throw '無効なファイル指定です'
        }
        if ($Path.Trim() -notlike '*.bmp' -and $Path.Trim() -notlike '*.jpg' -and $Path.Trim() -notlike '*.pdf') {
            throw 'ファイルはbmp, jpgかpdfのみ対応しています'
        }
        $fullPath = (Get-Item -Path $Path).FullName
    } catch {
        Write-Debug $_
        $errorList += "登録失敗 患者ID: ${Id} ファイル名: ${fullPath} - ${$_.Exception.Message}"
        return
    }

    # 設定画面を取得する
    # write-host '設定ウインドウへ遷移'
    $setteiWindow = Get-DMSetteiWindow -AppWindow $appWindow
    if ($null -eq $setteiWindow) {
        throw '文書管理の状態が異常です'
    }

    # 診療科コードが指定されていたら診療科コードを取得してリストから選択肢にあわせて適当そうなものを割り当てる
    $regulatedKaCode = ''
    try {
        if ($kaCode.Trim()) {
            if ($null -eq $kaCodeList) {
                $kaCodeList = Get-UIAComboBoxItems (Get-UIAComboBox -Parent $setteiWindow -Id 'CboSnk')
            }
            $regulatedKaCode = $kaCodeList | Where-Object { $_ -match "$kaCode" } | Select-Object -First 1
        }
    } catch {
        Write-Error $_
        throw '診療科コードの取得に失敗しました'
    }

    # 設定ウインドウの内容を設定してリストへ
    try {
        [void](Set-DMSetteiWindow -setteiWindow $setteiWindow -kanjyaId $Id -kaCode $regulatedKaCode -ishiCode $IshiCode.Trim() -nyuugai $Nyuugai.Trim())
    } catch {
        Write-Debug $_
        $errorList += "登録失敗 患者ID: ${Id} ファイル名: ${fullPath} - ${$_.Exception.Message}"
        return
    }

    # 新規画像を開いて諸元を入力
    $shinkiWindow = $null
    try {
        $shinkiWindow = Set-DMNewImageWindowProperties -Element $appWindow -DocumentName $Title -Description $Description -Date $Date
        # 画像を選択して保存
        [void](Register-DMNewImage -Element $shinkiWindow -Path $fullPath)
    } catch {
        Write-Debug $_
        $errorList += "登録失敗 患者ID: ${Id} ファイル名: ${fullPath} - ${$_.Exception.Message}"
        return
    }
}

end {
    if ($null -ne $appWindow) {
        Close-DocumentManager $appWindow
    }
    if ($errorList.Count -gt 0) {
        Write-Error "登録に失敗した項目が ${errorList.Count} 件ありました。詳細は以下の通りです。"
        $errorList | ForEach-Object { Write-Host $_ }
    }
}
