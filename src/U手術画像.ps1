# マイドキュメントにある取り込み対象を文書管理に保存するパイプラインオブジェクトにする
#
# ファイル：[日付] [患者ID][r:レポートフラグ].pdf
param (
    $Path = [System.Environment]::GetFolderPath("MyDocuments")
)

if (-not (Test-Path -Path $Path -PathType Container)) {
    throw "${$Path}はフォルダではありません."
}

$files = (Get-ChildItem -Path $Path | Where-Object { $_.Name -match '^20\d{6} \d{7}r?\.pdf' } | Select-Object -Property Name,FullName)

if ($null -eq $files -or $files.Count -eq 0) {
    exit 0
}

$returns = $files | % {
    $values = ([string]$_.Name).Split(' .')

    $Date = [string]$values[0].Insert(4, '-').Insert(7, '-')
    $Description = ''
    if ($values[1] -notlike '*r') {
        $Id = [string]$values[1]
    } else {
        $Id = [string]$values[1].substring(0, 7)
        $Description = 'レポート'
    }

    ,[PSCustomObject]@{
        Id = $Id
        Path = $_.FullName
        Date = $Date
        Title = '手術画像'
        Description = $Description
    }
}

$returns
