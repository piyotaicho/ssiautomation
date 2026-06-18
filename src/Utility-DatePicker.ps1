# 表示されたウインドウで日付を取得します
# returns: 選択された日付 (string[]) - 期間で選択された場合は開始日と終了日を返します
Param(
    [string]$Title = "日付を選択してください",
    [string]$InitialDate = (Get-Date).ToString("yyyy/MM/dd"),
    [string]$Format = "yyyyMMdd",
    [switch]$Duration = $false
)

Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase, System.Xaml

function ConvertTo-DatePickerDate {
    param([string]$Value)

    try {
        return [datetime](Get-Date -Date $Value -ErrorAction Stop)
    } catch {
        throw "InitialDate を解釈できません: $Value"
    }
}

function ConvertTo-XamlText {
    param([string]$Value)
    return [System.Security.SecurityElement]::Escape($Value)
}

function Format-DatePickerDate {
    param([datetime]$Value, [string]$Format = $Script:Format)

    return $Value.ToString($Format)
}

function Get-DatePickerMonthStart {
    param([datetime]$Value)

    return Get-Date -Year $Value.Year -Month $Value.Month -Day 1
}

function Test-DatePickerInRange {
    param(
        [datetime]$Value,
        [datetime]$Start,
        $End
    )

    # Durationモードではない場合は常にfalseを返す
    if ($null -eq $Start -or $null -eq $End) {
        return $false
    }

    if ($End.GetType().Name -ne 'DateTime') {
        throw "End パラメーターは DateTime 型もしくは Nullである必要があります."
    }
    return ($Value.Date -ge $Start.Date -and $Value.Date -le $End.Date)
}

function Update-DatePickerDisplay {
    if ($Duration) {
        $txtStartDate.Text = if ($null -ne $script:SelectedStart) { Format-DatePickerDate $script:SelectedStart 'yyyy/MM/dd'} else { '' }
        $txtEndDate.Text = if ($null -ne $script:SelectedEnd) { Format-DatePickerDate $script:SelectedEnd 'yyyy/MM/dd'} else { '' }
        $txtDateLabelDuration.Text = if ($null -ne $script:SelectedEnd) { "$Title（OKで確定します）" } else { "$Title（終了日を選択してください）" }
    } else {
        $txtSingleDate.Text = if ($null -ne $script:SelectedStart) { Format-DatePickerDate $script:SelectedStart 'yyyy/MM/dd'} else { '' }
        $txtDateLabelSingle.Text = "$Title（OKで確定します）"
    }

    $btnOk.IsEnabled = if ($Duration) { $null -ne $script:SelectedStart -and $null -ne $script:SelectedEnd } else { $null -ne $script:SelectedStart }
}

function Update-DatePickerCalendar {
    $script:ViewMonth = Get-DatePickerMonthStart $script:ViewMonth
    $txtMonthTitle.Text = $script:ViewMonth.ToString('yyyy年 M月')

    $firstDay = Get-Date -Year $script:ViewMonth.Year -Month $script:ViewMonth.Month -Day 1
    $offset = [int]$firstDay.DayOfWeek
    $daysInMonth = [datetime]::DaysInMonth($script:ViewMonth.Year, $script:ViewMonth.Month)
    $today = (Get-Date).Date

    for ($index = 0; $index -lt 42; $index++) {
        $button = $script:DateButtons[$index]
        $dayNumber = $index - $offset + 1

        # 日付が有効な範囲外の場合はボタンを無効化
        if ($dayNumber -lt 1 -or $dayNumber -gt $daysInMonth) {
            $button.Content = ''
            $button.Tag = $null
            $button.IsEnabled = $false
            $button.Background = [System.Windows.Media.Brushes]::WhiteSmoke
            $button.Foreground = [System.Windows.Media.Brushes]::Gray
            $button.BorderBrush = [System.Windows.Media.Brushes]::Gainsboro
            continue
        }

        $dayDate = Get-Date -Year $script:ViewMonth.Year -Month $script:ViewMonth.Month -Day $dayNumber
        $button.Content = $dayNumber
        $button.FontSize = 17
        $button.FontWeight = 'Normal'
        $button.Background = [System.Windows.Media.Brushes]::White
        $button.Foreground = [System.Windows.Media.Brushes]::Black

        $button.Tag = $dayDate

        # ボタンの表示スタイルをいじる
        if ($dayDate.Date -eq $today) {
            $button.FontWeight = 'Bold'
        }

        # Durationモードでも$isStartが選択される
        $isStart = $null -ne $script:SelectedStart -and $dayDate.Date -eq $script:SelectedStart.Date
        $isEnd = $null -ne $script:SelectedEnd -and $dayDate.Date -eq $script:SelectedEnd.Date

        # 選択日のスタイルを変更
        if ($isStart -or $isEnd) {
            $button.Background = [System.Windows.Media.Brushes]::DodgerBlue
            $button.Foreground = [System.Windows.Media.Brushes]::White
            $button.BorderBrush = [System.Windows.Media.Brushes]::DodgerBlue
            $button.FontWeight = 'Bold'
            continue
        }

        # 選択範囲内の日付のスタイルを変更
        if (Test-DatePickerInRange -Value $dayDate -Start $script:SelectedStart -End $script:SelectedEnd) {
            $button.Background = [System.Windows.Media.Brushes]::LightSteelBlue
            $button.BorderBrush = [System.Windows.Media.Brushes]::DodgerBlue
            continue
        }

        # 選択に被らないものはデフォルトのスタイルに戻す
        $button.Background = [System.Windows.Media.Brushes]::White
        $button.Foreground = [System.Windows.Media.Brushes]::Black
        $button.BorderBrush = [System.Windows.Media.Brushes]::LightGray
    }

    Update-DatePickerDisplay
}

function Select-DatePickerDate {
    param([datetime]$Value)

    $selectedDate = $Value.Date

    if (-not $Duration) {
        $script:SelectedStart = $selectedDate
        $script:SelectedEnd = $null
    } elseif ($null -eq $script:SelectedStart -or $null -ne $script:SelectedEnd) {
        $script:SelectedStart = $selectedDate
        $script:SelectedEnd = $null
    } elseif ($selectedDate -lt $script:SelectedStart.Date) {
        $script:SelectedEnd = $script:SelectedStart
        $script:SelectedStart = $selectedDate
    } elseif ($selectedDate -eq $script:SelectedStart.Date) {
        $script:SelectedEnd = $script:SelectedStart
    } else {
        $script:SelectedEnd = $selectedDate
    }

    $script:ViewMonth = Get-DatePickerMonthStart $selectedDate
    Update-DatePickerCalendar
}

$script:SelectedStart = ConvertTo-DatePickerDate $InitialDate
$script:SelectedEnd = $null
$script:ViewMonth = Get-DatePickerMonthStart $script:SelectedStart
$script:DialogAccepted = $false

$singleVisibility = if ($Duration) { 'Collapsed' } else { 'Visible' }
$rangeVisibility = if ($Duration) { 'Visible' } else { 'Collapsed' }

$xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="日付選択"
        Width="520" Height="565"
        MinWidth="520" MinHeight="565"
        WindowStartupLocation="CenterScreen"
        ResizeMode="NoResize"
        Background="#FFF7F7F7"
        FontFamily="Yu Gothic UI">
    <Grid Margin="8">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto" />
            <RowDefinition Height="Auto" />
        </Grid.RowDefinitions>

        <Border Grid.Row="0"
                Background="White"
                BorderBrush="#D6D6D6"
                BorderThickness="1"
                CornerRadius="8"
                Padding="6">
            <Grid>
                <Grid.RowDefinitions>
                    <RowDefinition Height="Auto" />
                    <RowDefinition Height="Auto" />
                    <RowDefinition Height="Auto" />
                </Grid.RowDefinitions>

                <Grid Grid.Row="0" Margin="0,0,0,8">
                    <Grid.ColumnDefinitions>
                        <ColumnDefinition Width="Auto" />
                        <ColumnDefinition Width="*" />
                        <ColumnDefinition Width="Auto" />
                    </Grid.ColumnDefinitions>

                    <Button Name="BtnPrevMonth"
                            Content="&lt;"
                            Width="32"
                            Height="28"
                            Background="#F0F0F0"
                            Foreground="#222"
                            BorderThickness="0"
                            FontSize="15" />

                    <TextBlock Name="TxtMonthTitle"
                               Grid.Column="1"
                               HorizontalAlignment="Center"
                               VerticalAlignment="Center"
                               FontSize="17"
                               FontWeight="SemiBold"
                               Foreground="#222" />

                    <Button Name="BtnNextMonth"
                            Grid.Column="2"
                            Content="&gt;"
                            Width="32"
                            Height="28"
                            Background="#F0F0F0"
                            Foreground="#222"
                            BorderThickness="0"
                            FontSize="15" />
                </Grid>

                <UniformGrid Name="CalendarGrid"
                             Grid.Row="1"
                             Rows="7"
                             Columns="7"
                             Margin="0,2,0,0"
                             HorizontalAlignment="Center"
                             VerticalAlignment="Center" />

                <Grid Grid.Row="2" Margin="0,6,0,0">
                    <Grid.RowDefinitions>
                        <RowDefinition Height="Auto" />
                        <RowDefinition Height="Auto" />
                    </Grid.RowDefinitions>

                    <Grid Grid.Row="0">
                        <Border Background="#FAFAFA"
                                BorderBrush="#E0E0E0"
                                BorderThickness="1"
                                CornerRadius="6"
                                Padding="8"
                                Margin="0,0,0,6">
                            <Grid>
                                <StackPanel Name="PnlSingle" Visibility="$singleVisibility">
                                    <TextBlock Name="TxtDateLabelSingle" Text="日付（OKで確定します）" Foreground="#555" Margin="0,0,0,4" FontSize="14" />
                                    <TextBox Name="TxtSingleDate"
                                             IsReadOnly="True"
                                             FontSize="15"
                                             Padding="7,5"
                                             Background="White"
                                             BorderBrush="#CFCFCF" />
                                </StackPanel>

                                <StackPanel Name="PnlDuration" Visibility="$rangeVisibility">
                                    <TextBlock Name="TxtDateLabelDuration" Text="日付（終了日を選択してください）" Foreground="#555" Margin="0,0,0,8" FontSize="14" />
                                    <StackPanel Orientation="Horizontal">
                                        <StackPanel Width="185" Margin="0,0,10,0">
                                            <TextBlock Text="開始日" Foreground="#555" Margin="0,0,0,4" FontSize="13" />
                                            <TextBox Name="TxtStartDate"
                                                     IsReadOnly="True"
                                                     FontSize="15"
                                                     Padding="7,5"
                                                     Background="White"
                                                     BorderBrush="#CFCFCF" />
                                        </StackPanel>
                                        <StackPanel Width="185">
                                            <TextBlock Text="終了日" Foreground="#555" Margin="0,0,0,4" FontSize="13" />
                                            <TextBox Name="TxtEndDate"
                                                     IsReadOnly="True"
                                                     FontSize="15"
                                                     Padding="7,5"
                                                     Background="White"
                                                     BorderBrush="#CFCFCF" />
                                        </StackPanel>
                                    </StackPanel>
                                </StackPanel>
                            </Grid>
                        </Border>
                    </Grid>

                    <StackPanel Grid.Row="1"
                                Orientation="Horizontal"
                                HorizontalAlignment="Right">
                        <Button Name="BtnOk"
                                Content="OK"
                                Width="100"
                                Height="32"
                                Margin="0,0,8,0"
                                IsDefault="True"
                                Background="#2F6FED"
                                Foreground="White"
                                BorderThickness="0" />
                        <Button Name="BtnCancel"
                                Content="キャンセル"
                                Width="100"
                                Height="32"
                                IsCancel="True"
                                Background="#E6E6E6"
                                Foreground="#222"
                                BorderThickness="0" />
                    </StackPanel>
                </Grid>
            </Grid>
        </Border>
    </Grid>
</Window>
"@

$reader = [System.Xml.XmlReader]::Create([System.IO.StringReader]$xaml)
$window = [System.Windows.Markup.XamlReader]::Load($reader)

$txtMonthTitle = $window.FindName('TxtMonthTitle')
$calendarGrid = $window.FindName('CalendarGrid')
$btnPrevMonth = $window.FindName('BtnPrevMonth')
$btnNextMonth = $window.FindName('BtnNextMonth')
$txtSingleDate = $window.FindName('TxtSingleDate')
$txtStartDate = $window.FindName('TxtStartDate')
$txtEndDate = $window.FindName('TxtEndDate')
$txtDateLabelSingle = $window.FindName('TxtDateLabelSingle')
$txtDateLabelDuration = $window.FindName('TxtDateLabelDuration')
$btnOk = $window.FindName('BtnOk')
$btnCancel = $window.FindName('BtnCancel')

$script:DateButtons = @()

foreach ($weekday in @('日', '月', '火', '水', '木', '金', '土')) {
    $header = New-Object System.Windows.Controls.TextBlock
    $header.Text = $weekday
    $header.HorizontalAlignment = 'Center'
    $header.VerticalAlignment = 'Center'
    $header.Margin = '2'
    $header.FontWeight = 'SemiBold'
    if ($weekday -eq '日') {
        $header.Foreground = [System.Windows.Media.Brushes]::Crimson
    } elseif ($weekday -eq '土') {
        $header.Foreground = [System.Windows.Media.Brushes]::DodgerBlue
    } else {
        $header.Foreground = [System.Windows.Media.Brushes]::DimGray
    }
    [void]$calendarGrid.Children.Add($header)
}

for ($index = 0; $index -lt 42; $index++) {
    $dayButton = New-Object System.Windows.Controls.Button
    $dayButton.Width = 58
    $dayButton.Height = 42
    $dayButton.Margin = '2'
    $dayButton.Padding = '0'
    $dayButton.FontSize = 16
    $dayButton.Background = [System.Windows.Media.Brushes]::White
    $dayButton.Foreground = [System.Windows.Media.Brushes]::Black
    $dayButton.BorderBrush = [System.Windows.Media.Brushes]::LightGray
    $dayButton.BorderThickness = '1'
    $dayButton.Tag = $null
    $dayButton.Add_Click({
        param($sender, $eventArgs)
        if ($null -ne $sender.Tag) {
            Select-DatePickerDate -Value ([datetime]$sender.Tag)
        }
    })
    $script:DateButtons += $dayButton
    [void]$calendarGrid.Children.Add($dayButton)
}

$btnPrevMonth.Add_Click({
    $script:ViewMonth = $script:ViewMonth.AddMonths(-1)
    Update-DatePickerCalendar
})

$btnNextMonth.Add_Click({
    $script:ViewMonth = $script:ViewMonth.AddMonths(1)
    Update-DatePickerCalendar
})

$btnOk.Add_Click({
    $script:DialogAccepted = $true
    $window.Close()
})

$btnCancel.Add_Click({
    $script:DialogAccepted = $false
    $window.Close()
})

$window.Add_Closing({
    if (-not $script:DialogAccepted) {
        $script:DialogAccepted = $false
    }
})

Update-DatePickerCalendar

$null = $window.ShowDialog()

if (-not $script:DialogAccepted) {
    return $null
}

if ($Duration) {
    if ($null -eq $script:SelectedEnd) {
        $script:SelectedEnd = $script:SelectedStart
    }
    return @(
        (Format-DatePickerDate $script:SelectedStart),
        (Format-DatePickerDate $script:SelectedEnd)
    )
}

return @((Format-DatePickerDate $script:SelectedStart))
