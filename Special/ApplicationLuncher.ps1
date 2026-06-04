param (
    [string]$JsonPath = "apps.json",
    [string]$Title = "部門用 アプリケーション・ランチャー",
    [Alias("NoAutoLaunch")]
    [switch]$NoAutolunch
)

# 1. 必要なアセンブリのロード
Add-Type -AssemblyName PresentationFramework, System.Windows.Forms, WindowsBase, System.Drawing

# 2. サンプルJSONファイルの自動生成（ファイルがない場合のみ）
if (-not (Test-Path $JsonPath)) {
    throw 'JSONファイルが見つかりませんでした。'
}

$AppList = Get-Content $JsonPath -Raw | ConvertFrom-Json

$windowTitle = $Title
if ([string]::IsNullOrWhiteSpace($windowTitle)) {
    $windowTitle = "部門用 アプリケーション・ランチャー"
}

# 3. 画面レイアウト (XAML)
# WindowStyle="SingleBorderWindow" (標準枠) にし、AllowsTransparencyを外すことでタイトルバーを維持。
# 代わりに、ウィンドウ全体の不透明度を Opacity="0.9"、背景を #99000000（透過）に調整。
$xml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    Title="$windowTitle" 
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
$Global:AutoLaunchDone = $NoAutolunch.IsPresent

function Resolve-AppExecutablePath ([string]$path) {
    if ([string]::IsNullOrWhiteSpace($path)) { return $null }

    if ([System.IO.Path]::IsPathRooted($path) -and (Test-Path $path)) {
        return (Resolve-Path -Path $path).Path
    }

    try {
        $commands = @(Get-Command -Name $path -CommandType Application -ErrorAction Stop)
        foreach ($cmd in $commands) {
            $source = [string]$cmd.Source
            if (-not [string]::IsNullOrWhiteSpace($source) -and (Test-Path -LiteralPath $source -PathType Leaf)) {
                return (Resolve-Path -LiteralPath $source).Path
            }
        }

        if ($commands.Count -gt 0) {
            $firstSource = [string]$commands[0].Source
            if (-not [string]::IsNullOrWhiteSpace($firstSource)) {
                return $firstSource
            }
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

function Get-AppArguments ([object]$tile) {
    $argumentValues = @()

    if ($null -eq $tile) {
        return $argumentValues
    }

    # 推奨: Arguments。後方互換で CommandLineOptions も受け付ける。
    $source = $tile.Arguments
    if ($null -eq $source) {
        $source = $tile.CommandLineOptions
    }

    if ($null -eq $source) {
        return $argumentValues
    }

    if ($source -is [string]) {
        if (-not [string]::IsNullOrWhiteSpace($source)) {
            $argumentValues += $source
        }
    } else {
        foreach ($arg in $source) {
            if (-not [string]::IsNullOrWhiteSpace([string]$arg)) {
                $argumentValues += [string]$arg
            }
        }
    }

    return $argumentValues
}

function Get-AppArgumentSource ([object]$tile) {
    if ($null -eq $tile) {
        return $null
    }

    if ($null -ne $tile.Arguments) {
        return $tile.Arguments
    }

    if ($null -ne $tile.CommandLineOptions) {
        return $tile.CommandLineOptions
    }

    return $null
}

function Convert-ArgumentStringToArray ([string]$argumentLine) {
    $result = @()
    if ([string]::IsNullOrWhiteSpace($argumentLine)) {
        return $result
    }

    # "quoted value" を1トークンとして扱い、それ以外は空白で分割する。
    $regexMatches = [regex]::Matches($argumentLine, '"([^"\\]|\\.)*"|\S+')
    foreach ($m in $regexMatches) {
        $token = $m.Value
        if ($token.Length -ge 2 -and $token.StartsWith('"') -and $token.EndsWith('"')) {
            $token = $token.Substring(1, $token.Length - 2)
        }
        if (-not [string]::IsNullOrWhiteSpace($token)) {
            $result += $token
        }
    }

    return $result
}

function Start-AppProcess ([string]$resolvedPath, [object]$tile) {
    $argumentSource = Get-AppArgumentSource $tile
    $workingDirectory = $null
    try {
        $candidateDirectory = [System.IO.Path]::GetDirectoryName($resolvedPath)
        if (-not [string]::IsNullOrWhiteSpace($candidateDirectory) -and (Test-Path -LiteralPath $candidateDirectory -PathType Container)) {
            $workingDirectory = $candidateDirectory
        }
    } catch {
        $workingDirectory = $null
    }

    $startProcessArgs = @{ FilePath = $resolvedPath; ErrorAction = 'Stop' }
    if ($null -ne $workingDirectory) {
        $startProcessArgs.WorkingDirectory = $workingDirectory
    }

    if ($null -eq $argumentSource) {
        Start-Process @startProcessArgs
        return
    }

    if ($argumentSource -is [string]) {
        if ([string]::IsNullOrWhiteSpace($argumentSource)) {
            Start-Process @startProcessArgs
        } else {
            $launchErrors = @()

            # 方式1: ユーザー記述どおりの文字列をそのまま渡す。
            try {
                Start-Process @startProcessArgs -ArgumentList $argumentSource
                return
            } catch {
                $launchErrors += $_.Exception.Message
            }

            # 方式2: 引数文字列を分割して配列で渡す。
            $parsedArguments = Convert-ArgumentStringToArray $argumentSource
            if ($parsedArguments.Count -gt 0) {
                try {
                    Start-Process @startProcessArgs -ArgumentList $parsedArguments
                    return
                } catch {
                    $launchErrors += $_.Exception.Message
                }
            }

            throw [System.InvalidOperationException]::new(("引数付き起動に失敗しました。{0}" -f (($launchErrors | Select-Object -Unique) -join " / ")))
        }
        return
    }

    $arguments = Get-AppArguments $tile
    if ($arguments.Count -gt 0) {
        $launchErrors = @()

        # 方式1: 配列をそのまま渡す。
        try {
            Start-Process @startProcessArgs -ArgumentList $arguments
            return
        } catch {
            $launchErrors += $_.Exception.Message
        }

        # 方式2: 連結文字列で渡す（受け取り側仕様差の吸収）。
        $joinedArguments = $arguments -join ' '
        if (-not [string]::IsNullOrWhiteSpace($joinedArguments)) {
            try {
                Start-Process @startProcessArgs -ArgumentList $joinedArguments
                return
            } catch {
                $launchErrors += $_.Exception.Message
            }
        }

        throw [System.InvalidOperationException]::new(("引数付き起動に失敗しました。{0}" -f (($launchErrors | Select-Object -Unique) -join " / ")))
    } else {
        Start-Process @startProcessArgs
    }
}

function Normalize-ProcessExecutableName ([string]$processName) {
    if ([string]::IsNullOrWhiteSpace($processName)) {
        return $null
    }

    $trimmed = $processName.Trim()
    if ($trimmed.ToLowerInvariant().EndsWith('.exe')) {
        return $trimmed
    }

    return "$trimmed.exe"
}

function Normalize-CommandLineText ([string]$text) {
    if ([string]::IsNullOrWhiteSpace($text)) {
        return ""
    }

    return ([regex]::Replace($text.Trim().ToLowerInvariant(), '\s+', ' '))
}

function Test-ArgumentTokenMatch ([string]$actualToken, [string]$expectedToken) {
    if ([string]::IsNullOrWhiteSpace($actualToken) -or [string]::IsNullOrWhiteSpace($expectedToken)) {
        return $false
    }

    $actualNormalized = $actualToken.Trim().Trim('"').ToLowerInvariant()
    $expectedNormalized = $expectedToken.Trim().Trim('"').ToLowerInvariant()

    return ($actualNormalized -eq $expectedNormalized)
}

function Test-CommandLineIncludesArguments ([string]$actualCommandLine, [string[]]$expectedArguments) {
    if ([string]::IsNullOrWhiteSpace($actualCommandLine)) {
        return $false
    }

    if ($null -eq $expectedArguments -or $expectedArguments.Count -eq 0) {
        return $true
    }

    $actualTokens = Convert-ArgumentStringToArray $actualCommandLine
    if ($actualTokens.Count -eq 0) {
        return $false
    }

    # 先頭は実行ファイルパス想定のため除外し、純粋な引数のみで比較する。
    $actualArguments = @()
    if ($actualTokens.Count -gt 1) {
        $actualArguments = $actualTokens[1..($actualTokens.Count - 1)]
    }

    $expectedTokens = @()
    foreach ($expected in $expectedArguments) {
        if ([string]::IsNullOrWhiteSpace([string]$expected)) {
            continue
        }

        $parsed = Convert-ArgumentStringToArray ([string]$expected)
        if ($parsed.Count -gt 0) {
            $expectedTokens += $parsed
        } else {
            $expectedTokens += [string]$expected
        }
    }

    if ($expectedTokens.Count -eq 0) {
        return $true
    }

    $searchIndex = 0
    foreach ($expectedToken in $expectedTokens) {
        if ([string]::IsNullOrWhiteSpace($expectedToken)) {
            continue
        }

        $matched = $false
        for ($i = $searchIndex; $i -lt $actualArguments.Count; $i++) {
            if (Test-ArgumentTokenMatch -actualToken $actualArguments[$i] -expectedToken $expectedToken) {
                $matched = $true
                $searchIndex = $i + 1
                break
            }
        }

        if (-not $matched) {
            return $false
        }
    }

    return $true
}

function Test-AppIsRunning ([object]$tile, [string]$resolvedPath) {
    $processNames = Get-AppProcessNames $tile $resolvedPath
    $arguments = Get-AppArguments $tile
    $hasExplicitArgumentFilter = ($null -ne $tile -and (($null -ne $tile.Arguments) -or ($null -ne $tile.CommandLineOptions)))

    if ($hasExplicitArgumentFilter) {
        if ($arguments.Count -eq 0) {
            return $false
        }

        foreach ($name in $processNames) {
            $exeName = Normalize-ProcessExecutableName $name
            if ([string]::IsNullOrWhiteSpace($exeName)) {
                continue
            }

            $escapedExeName = $exeName.Replace("'", "''")
            $processes = Get-CimInstance -ClassName Win32_Process -Filter "Name='$escapedExeName'" -ErrorAction SilentlyContinue
            foreach ($proc in $processes) {
                if (Test-CommandLineIncludesArguments $proc.CommandLine $arguments) {
                    return $true
                }
            }
        }

        return $false
    }

    foreach ($name in $processNames) {
        $lookupName = [System.IO.Path]::GetFileNameWithoutExtension($name)
        if (Get-Process -Name $lookupName -ErrorAction SilentlyContinue) {
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
                        Start-AppProcess -resolvedPath $resolvedPath -tile $targetApp
                    }
                }
                $Global:IdleCounter = 0
                $Global:IsActiveTracking = $true
                Update-SystemState
            } catch {
                [System.Windows.MessageBox]::Show("起動に失敗しました: $($targetApp.Path)`n$($_.Exception.Message)", "Error")
            }
        }
    })

    $tileObj = [PSCustomObject]@{
        UI       = $border
        Image    = $img
        Name     = $app.Name
        Path     = $app.Path
        ProcessNames = $app.ProcessNames
        Arguments = $app.Arguments
        CommandLineOptions = $app.CommandLineOptions
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
        try {
            $resolvedPath = Resolve-AppExecutablePath $tile.Path
            if ($resolvedPath) {
                # 色のキャッシュ値ではなく都度の起動判定を使い、引数一致まで確認する。
                if (-not (Test-AppIsRunning $tile $resolvedPath)) {
                    Start-AppProcess -resolvedPath $resolvedPath -tile $tile
                }
            }
        } catch {}
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
