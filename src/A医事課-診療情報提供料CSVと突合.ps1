# 医事データー(診療情報提供料（１）のCSVをパースする
param (
    [Parameter(Mandatory)][string]$Path,
    [switch]$showResult = $false
)

# 診療科テーブル
. "$PSScriptRoot\Tools-MasterTables.ps1"

# 引数のチェック
if (-not (Test-Path -PathType Leaf -Path $Path)) {
    throw "$Path は有効なファイル名ではありません."
}

$content = Get-Content -Path $Path -Encoding Default

$check = $content[6]

# 抽出内容のチェック - 電算コードが正しいか - 電算コードがあるはずのところをチェックする
if ($check -notlike '*180016110*') {
    throw "$Path は診療情報提供料(1)の実施抽出データーではない可能性があります."
}

# CSVのパース
$CSVdata = $content | Select-Object -Skip 7 | Where-Object { $_ -ne '' } | Select-Object -SkipLast 1 | ConvertFrom-Csv


if ($CSVdata.Count -eq 0) {
    throw '対象がありません'
}

# 突合のため整形
$CSVdata = $CSVdata | Select-Object @( `
            @{Name='日付'; Expression={$_.'診療日'}}, `
            @{Name='ID'; Expression={$_.'患者ID'}}, `
            @{Name='患者氏名'; Expression={$_.'氏名'}}, `
            @{Name='診療科'; Expression={$MasterKaCode[[Int]$_.'診療科']}}, `
            @{Name='入外'; Expression={if ($_.'部屋' -eq '外来') { '外来' } else { '入院' }}}, `
            @{Name='担当医'; Expression={$_.'Dr氏名'}} `
            )

# 紹介状データーを取得
# 算定日付の最初の日付から4週間前から最後の日付から3日間
$fromDate = (Get-Date ($CSVdata | Sort-Object '日付' | Select-Object -First 1).'日付').AddDays(-28).ToString('yyyy/MM/dd')
$toDate = (Get-Date ($CSVdata | Sort-Object '日付' | Select-Object -Last 1).'日付').AddDays(+3).ToString('yyyy/MM/dd')

# 紹介状データーは日付逆順でソートしておく
$referral = (& "$PSScriptRoot\A紹介患者一覧.ps1" -Start $fromDate -End $toDate) | Sort-Object -Property 日付 -Descending

$result = @()

$CSVdata | ForEach-Object {
    if ($_.ID -in ($referral.ID)) {
        # 基準となる算定日
        $santeiDate = Get-Date $_.日付

        # 紹介患者一覧から抽出
        $idMatchedReferral = $referral | Where-Object ID -EQ $_.ID
        $foundReferral = $idMatchedReferral[0]

        if ($idMatchedReferral.Count -gt 1) {
            # 基準日からの絶対値でソート
            $idMatchedReferral = $idMatchedReferral |`
                Select-Object *, @{Name='Interval';Expression={[math]::Abs(((Get-Date $_.日付)-$santeiDate).TotalDays)}} | `
                ForEach-Object { [PSCustomObject]$_ } | `
                Sort-Object -Property Interval

            # 抽出(1) 医師名
            $filterdReferral = $idMatchedReferral | Where-Object 担当医 -EQ $_.担当医
            if ($filterdReferral.Count -eq 0) {
                # 抽出(2) 診療科
                $filterdReferral = $idMatchedReferral | Where-Object 診療科 -EQ $_.診療科
                if ($filterdRefferral.Count -eq 0) {
                    $result += , ($_ | Select-Object *, @{Name='紹介先医療機関名'; Expression={''}})
                    continue
                }
            }

            # 複数抽出の際はもっともインターバルの短いものを採用するので[0]
            $foundReferral = $filterdReferral[0]
        }

        if ($showResult) {
            Write-Host "算定日: $($_.日付)"
            Write-Host "ID: $($_.ID) 患者氏名: $($_.患者氏名) 入外: $($_.入外)"
            Write-Host "担当医: $($_.担当医) ($($_.診療科))"
            Write-Host -NoNewline '紹介先医療機関名: '

            # 紹介先医療機関名
            Write-Host $foundReferral.'紹介先医療機関名'
            Write-Host '---'
        }
        $result += , ($_ | Select-Object *, @{Name='紹介先医療機関名'; Expression={$foundReferral.'紹介先医療機関名'}})
    } else {
        $result += , ($_ | Select-Object *, @{Name='紹介先医療機関名'; Expression={''}})
    }
}

if ($showResult) {
    Write-Host -ForegroundColor Cyan "$($CSVdata.Count)件のうち $(($result | Where-Object 紹介先医療機関名 -NE '').Count)件の紹介先を取得しました."
}

if ($showResult -and ($result | Where-Object 紹介先医療機関名 -EQ '').Count -gt 0) {
    Write-Host -ForegroundColor Red "以下 $(($result | Where-Object 紹介先医療機関名 -EQ '').Count)件の紹介先が取得できませんでした."
    $result | Where-Object 紹介先医療機関名 -EQ '' | ForEach-Object {
        Write-Host -ForegroundColor Red "算定日: $($_.日付)"
        Write-Host -ForegroundColor Red "ID: $($_.ID) 患者氏名: $($_.患者氏名) 入外: $($_.入外)"
        Write-Host -ForegroundColor Red "担当医: $($_.担当医) ($($_.診療科))"
        Write-Host -ForegroundColor Red '---'
     }
}

return $result
