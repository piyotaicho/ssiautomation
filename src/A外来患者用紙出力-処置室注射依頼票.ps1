#
# 外来 - コメントに 注射 点滴 の記載がある場合に処置室依頼票を出力する
# 
# -noFilter : コメントの有無にかかわらず全症例を対象とする
#
param(
    [Switch]$noFilter
)

# 追加アセンブリのロード
Add-Type -AssemblyName System.Windows.Forms

# オートメーションコアなどを導入
. "$PSScriptRoot\Tools-AutomationCore.ps1"
. "$PSScriptRoot\Tools-Entrance.ps1"
. "$PSScriptRoot\Tools-DocumentManager.ps1"
. "$PSScriptRoot\Tools-ExcelWindow.ps1"

# アプリケーションウインドウ
$appBunsyo = $null # 文書管理

try {
    # エントランスから情報を取得する
    $entranceGairai = Get-GairaiList

    if ($entranceGairai.List.Count -eq 0) {
        throw '処理対象がありません'
    }

    # 文書管理を起動
    $appBunsyo = Get-DocumentManager

    # ループ
    $setteiWindow = $null
    foreach($row in $entranceGairai.List) {
        # CSVから情報を取得
        [string]$Id = $row.'患者コード'
        [string]$Comment = $row.'患者一覧用コメント'

        # ダミー紹介患者 99xxxxxx はスキップする
        if ($Id -match '^99\d{6}$') {
            continue
        }

        # コメントに「注射」または「点滴」が含まれていなければスキップする
        if (-not $noFilter -and $Comment -notmatch '注射|点滴') {
            continue
        }

        # 文書管理の設定ウインドウで患者IDを設定
        $setteiWindow = Get-DMSetteiWindow -AppWindow $appBunsyo
        if ($null -eq $setteiWindow) {
            throw '文書管理の状態が異常です'
        }
        Set-DMSetteiWindow -setteiWindow $setteiWindow -kanjyaId $Id | Out-Null

        # 新規文書作成前のエクセルのプロセス状況を取得
        $excelPIDs = (Get-Process -Name excel -ErrorAction SilentlyContinue).Id

        # 新規文書を作成
        New-DMDocument -appWindow $appBunsyo -documentName '処置室依頼票' | Out-Null

        # 直近で起動した有効なExcelを抽出
        $retryCount = 3
        $newExcelHwnd = 0
        while (($retryCount--) -gt 0) {
            $newExcelProcs = (Get-Process -Name excel -ErrorAction SilentlyContinue | Where-Object { $_.Id -notin $excelPIDs -and $_.MainWindowHandle -ne 0})
            if ($newExcelProcs.Count -ge 1) {
                $newExcelHwnd = $newExcelProcs[0].MainWindowHandle
                break
            }
            Write-Host 'Getting Excel HWND again'
            [UIATools]::Sleep(1000)
            $retryCount--
        }

        if (-not $newExcelHwnd) {
            throw '有効なExcelプロセスの取得に失敗しました'
        }

        # COMオブジェクトを取得
        $excel = $null
        $activesheet = $null

        $excel = [ExcelWindowFactory]::GetExcelApplicationFromHwnd($newExcelHwnd)
        [UIATools]::Sleep()

        if ($null -eq $excel) {
            throw 'Excelのコントロール取得に失敗しました'
        }

        try {
            # 印刷
            $excel.Sheets[1].PrintOut(1, $true)

            # 保存せず終了
            # 文書管理標準Excelテンプレートが用意しているクリニカルパス対応用の機能を利用して保存フラグチェックを回避
            try {
                $excel.Run('GlobalExcel.CriticalpClose')
            }
            catch {}
            $excel.ActiveWorkbook.Close($false)
            $excel.Quit()
        } catch {
            throw $_
        } finally {
            # COMオブジェクトメモリ解放
            try {
                [System.Runtime.InteropServices.Marshal]::ReleaseComObject($activesheet) | Out-Null
                $activesheet = $null
            } catch {}
            try {
                [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
                $excel = $null
            } catch {}
        }
    }

    # 文書管理を終了する
    Close-DocumentManager $appBunsyo | Out-Null
} catch {
    # [System.Windows.Forms.MessageBox]::Show($_.Exception.Message, '外来患者用紙出力', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information, [System.Windows.Forms.MessageBoxDefaultButton]::Button1, [System.Windows.Forms.MessageBoxOptions]::ServiceNotification)
    throw $_
}
