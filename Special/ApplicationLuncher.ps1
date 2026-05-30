param (
    [string]$JsonPath = "apps.json"
)

# 1. 必要なアセンブリのロード
Add-Type -AssemblyName PresentationFramework, System.Windows.Forms, WindowsBase, System.Drawing

# 2. サンプルJSONファイルの自動生成（ファイルがない場合のみ）
if (-not (Test-Path $JsonPath)) {
    $sampleJson = @(
        @{ Name = "メモ帳"; Path = "C:\Windows\notepad.exe"; ProcessNames = @("Notepad") },
        @{ Name = "電卓"; Path = "C:\Windows\System32\calc.exe"; ProcessNames = @("CalculatorApp") },
        @{ Name = "ペイント"; Path = "C:\Windows\System32\mspaint.exe" },
        @{ Name = "存在しないアプリ"; Path = "C:\invalid_path\error.exe" }
    ) | ConvertTo-Json -Depth 5
    [System.IO.File]::WriteAllText($JsonPath, $sampleJson, [System.Text.Encoding]::UTF8)
}

$AppList = Get-Content $JsonPath -Raw | ConvertFrom-Json

# 3. 画面レイアウト (XAML)
# WindowStyle="SingleBorderWindow" (標準枠) にし、AllowsTransparencyを外すことでタイトルバーを維持。
# 代わりに、ウィンドウ全体の不透明度を Opacity="0.9"、背景を #99000000（透過）に調整。
$xml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        Title="部門用 アプリケーション・ランチャー" 
        WindowStyle="SingleBorderWindow" 
        ResizeMode="CanMinimize"
        WindowStartupLocation="CenterScreen" 
        Width="1024" Height="768" 
        Background="#99000000" 
        Opacity="0.95"
        ShowInTaskbar="True">
    <Grid>
        <ItemsControl Name="TileContainer" Margin="30,30,30,50">
            <ItemsControl.ItemsPanel>
                <ItemsPanelTemplate>
                    <UniformGrid Margin="5"/>
                </ItemsPanelTemplate>
            </ItemsControl.ItemsPanel>
        </ItemsControl>
        
        <TextBlock Text="[F5]:強制更新  |  各タイルをクリックすると個別起動します" 
                   Foreground="White" FontSize="12" HorizontalAlignment="Center" 
                   VerticalAlignment="Bottom" Margin="0,0,0,15" Opacity="0.7"/>
    </Grid>
</Window>
"@

$reader = [System.Xml.XmlReader]::Create([System.IO.StringReader]$xml)
$window = [System.Windows.Markup.XamlReader]::Load($reader)
$container = $window.FindName("TileContainer")

# 4. 状態管理用オブジェクト
$Global:Tiles = @()
$Global:LastStateHash = ""
$Global:IdleCounter = 0
$Global:IsActiveTracking = $true
$Global:AutoLaunchDone = $false

function Resolve-AppExecutablePath ([string]$path) {
    if ([string]::IsNullOrWhiteSpace($path)) { return $null }

    if ([System.IO.Path]::IsPathRooted($path) -and (Test-Path $path)) {
        return (Resolve-Path -Path $path).Path
    }

    try {
        $cmd = Get-Command -Name $path -CommandType Application -ErrorAction Stop
        if ($cmd -and $cmd.Source) {
            return $cmd.Source
        }
    } catch {
    }

    return $null
}

function Get-AppProcessNames ([object]$tile, [string]$resolvedPath) {
    $names = @()

    # JSON側で ProcessNames が指定されていれば最優先で使用
    if ($null -ne $tile -and $null -ne $tile.ProcessNames) {
        if ($tile.ProcessNames -is [string]) {
            if (-not [string]::IsNullOrWhiteSpace($tile.ProcessNames)) {
                $names += $tile.ProcessNames
            }
        } else {
            foreach ($name in $tile.ProcessNames) {
                if (-not [string]::IsNullOrWhiteSpace([string]$name)) {
                    $names += [string]$name
                }
            }
        }
    }

    # 未指定時の後方互換: 実行ファイル名を監視対象にする
    if ($names.Count -eq 0 -and -not [string]::IsNullOrWhiteSpace($resolvedPath)) {
        $baseName = [System.IO.Path]::GetFileNameWithoutExtension($resolvedPath)
        if (-not [string]::IsNullOrWhiteSpace($baseName)) {
            $names += $baseName
        }
    }

    return $names | Select-Object -Unique
}

function Test-AppIsRunning ([object]$tile, [string]$resolvedPath) {
    $processNames = Get-AppProcessNames $tile $resolvedPath
    foreach ($name in $processNames) {
        if (Get-Process -Name $name -ErrorAction SilentlyContinue) {
            return $true
        }
    }
    return $false
}

# アイコン抽出関数
function Get-AppIconSource ([string]$path) {
    $resolvedPath = Resolve-AppExecutablePath $path
    if (-not $resolvedPath) {
        return [System.Windows.Interop.Imaging]::CreateBitmapSourceFromHIcon(
            [System.Drawing.SystemIcons]::Error.Handle,
            [System.Windows.Int32Rect]::Empty,
            [System.Windows.Media.Imaging.BitmapSizeOptions]::FromWidthAndHeight(64, 64)
        )
    }

    try {
        $icon = [System.Drawing.Icon]::ExtractAssociatedIcon($resolvedPath)
        if ($null -eq $icon) {
            return [System.Windows.Interop.Imaging]::CreateBitmapSourceFromHIcon(
                [System.Drawing.SystemIcons]::Application.Handle,
                [System.Windows.Int32Rect]::Empty,
                [System.Windows.Media.Imaging.BitmapSizeOptions]::FromWidthAndHeight(64, 64)
            )
        }
        $bitmap = $icon.ToBitmap()
        $stream = New-Object System.IO.MemoryStream
        $bitmap.Save($stream, [System.Drawing.Imaging.ImageFormat]::Png)
        $stream.Position = 0
        $bi = New-Object System.Windows.Media.Imaging.BitmapImage
        $bi.BeginInit()
        $bi.StreamSource = $stream
        $bi.CacheOption = [System.Windows.Media.Imaging.BitmapCacheOption]::OnLoad
        $bi.EndInit()
        return $bi
    } catch {
        return [System.Windows.Interop.Imaging]::CreateBitmapSourceFromHIcon(
            [System.Drawing.SystemIcons]::Application.Handle,
            [System.Windows.Int32Rect]::Empty,
            [System.Windows.Media.Imaging.BitmapSizeOptions]::FromWidthAndHeight(64, 64)
        )
    }
}

# 5. タイルUIの動的生成
foreach ($app in $AppList) {
    $border = New-Object System.Windows.Controls.Border
    $border.Margin = 8
    $border.CornerRadius = 4
    $border.Cursor = [System.Windows.Input.Cursors]::Hand
    
    $grid = New-Object System.Windows.Controls.Grid
    $border.Child = $grid
    
    # アイコン (上側60%に寄せるように配置)
    $img = New-Object System.Windows.Controls.Image
    $img.VerticalAlignment = [System.Windows.VerticalAlignment]::Center
    $img.HorizontalAlignment = [System.Windows.HorizontalAlignment]::Center
    $img.Height = 64
    $img.Width = 64
    $grid.Children.Add($img) | Out-Null
    
    # アプリ名
    $txt = New-Object System.Windows.Controls.TextBlock
    $txt.Text = $app.Name
    $txt.Foreground = [System.Windows.Media.Brushes]::White
    $txt.FontSize = 16
    $txt.FontWeight = [System.Windows.FontWeights]::Bold
    $txt.HorizontalAlignment = [System.Windows.HorizontalAlignment]::Center
    $txt.VerticalAlignment = [System.Windows.VerticalAlignment]::Bottom
    $txt.Margin = New-Object System.Windows.Thickness(0,0,0,10)
    $grid.Children.Add($txt) | Out-Null

    # アイコン反映
    $iconSource = Get-AppIconSource $app.Path
    if ($null -ne $iconSource) {
        $img.Source = $iconSource
    }

    # クリックイベント
    $border.Add_MouseDown({
        param($senderElement, $e)
        $targetApp = $senderElement.Tag
        if ($targetApp.State -eq "Red") {
            try {
                $resolvedPath = Resolve-AppExecutablePath $targetApp.Path
                if ($resolvedPath) {
                    if (-not (Test-AppIsRunning $targetApp $resolvedPath)) {
                        Start-Process $resolvedPath
                    }
                }
                $Global:IdleCounter = 0
                $Global:IsActiveTracking = $true
                Update-SystemState
            } catch {
                [System.Windows.MessageBox]::Show("起動に失敗しました: $($targetApp.Path)", "Error")
            }
        }
    })

    $tileObj = [PSCustomObject]@{
        UI       = $border
        Image    = $img
        Name     = $app.Name
        Path     = $app.Path
        ProcessNames = $app.ProcessNames
        State    = "Unknown"
    }
    $border.Tag = $tileObj
    $Global:Tiles += $tileObj
    $container.Items.Add($border) | Out-Null
}

# 6. 状態更新ロジック
function Update-SystemState {
    if (-not $Global:IsActiveTracking) { return }

    $currentStateString = ""
    $bc = New-Object System.Windows.Media.BrushConverter

    foreach ($tile in $Global:Tiles) {
        $resolvedPath = Resolve-AppExecutablePath $tile.Path
        if (-not $resolvedPath) {
            $tile.State = "Gray"
            $tile.UI.Background = $bc.ConvertFromString("#FF7F7F7F") # グレー
        }
        else {
            $proc = Test-AppIsRunning $tile $resolvedPath
            if ($proc) {
                $tile.State = "Green"
                $tile.UI.Background = $bc.ConvertFromString("#FF28A745") # 緑
            } else {
                $tile.State = "Red"
                $tile.UI.Background = $bc.ConvertFromString("#FFDC3545") # 赤
            }
        }
        $currentStateString += "$($tile.Path):$($tile.State)|"
    }

    # 10秒変化がなければエコモード（停止）
    if ($currentStateString -eq $Global:LastStateHash) {
        $Global:IdleCounter++
        if ($Global:IdleCounter -ge 10) {
            $Global:IsActiveTracking = $false
        }
    } else {
        $Global:IdleCounter = 0
        $Global:LastStateHash = $currentStateString
    }
}

# 一括起動
function Start-AllApplications {
    foreach ($tile in $Global:Tiles) {
        if ($tile.State -eq "Red") {
            try {
                $resolvedPath = Resolve-AppExecutablePath $tile.Path
                if ($resolvedPath) {
                    if (-not (Test-AppIsRunning $tile $resolvedPath)) {
                        Start-Process $resolvedPath
                    }
                }
            } catch {}
        }
    }
    $Global:IdleCounter = 0
    $Global:IsActiveTracking = $true
    Update-SystemState
}

# 7. タイマー設定 (1秒周期)
$timer = New-Object System.Windows.Threading.DispatcherTimer
$timer.Interval = [TimeSpan]::FromSeconds(1)
$autoLaunchAt = [DateTime]::Now.AddSeconds(1)

$timer.Add_Tick({
    # 自動実行は起動後1秒の一度だけ。以降の更新処理ではトリガーしない。
    if (-not $Global:AutoLaunchDone -and [DateTime]::Now -ge $autoLaunchAt) {
        $Global:AutoLaunchDone = $true
        Start-AllApplications
        return
    }

    Update-SystemState
})
$timer.Start()

# 8. ウィンドウイベント
$window.Add_KeyDown({
    param($senderElement, $e)
    if ($e.Key -eq [System.Windows.Input.Key]::F5) {
        $Global:IdleCounter = 0
        $Global:IsActiveTracking = $true
        Update-SystemState
    }
})

$window.Add_Activated({
    $Global:IdleCounter = 0
    $Global:IsActiveTracking = $true
    Update-SystemState
})

# タイマー停止処理（閉じられたとき用）
$window.Add_Closed({
    $timer.Stop()
})

# 起動
$window.ShowDialog() | Out-Null
