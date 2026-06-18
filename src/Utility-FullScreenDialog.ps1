param (
    [string]$Message = "処理がエラーで中断もしくは終了しました。",
    [string]$Color = "#88FF0000" # デフォルト値
)

Add-Type -AssemblyName PresentationFramework, System.Windows.Forms, WindowsBase

# パラメーターを修正
$Message = $Message.Trim('"').Trim("'")
$Color = $Color.Trim('"').Trim("'").ToUpper()
if ($Color -notlike '#*') {
    $Color = '#' + $Color
}

# プライマリ画面の解像度（サイズ）を取得
$primaryScreen = [System.Windows.Forms.Screen]::PrimaryScreen
$screenWidth  = $primaryScreen.Bounds.Width
$screenHeight = $primaryScreen.Bounds.Height

$xml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        Title="Error Notification" WindowStyle="None" AllowsTransparency="True" Topmost="True" 
        Left="0" Top="0" Width="$screenWidth" Height="$screenHeight"
        ShowInTaskbar="True">
    
    <Grid HorizontalAlignment="Center" VerticalAlignment="Center">
        <Border Background="#F5F5F5" CornerRadius="8" Padding="30" Width="500" Margin="20">
            <Border.Effect>
                <DropShadowEffect BlurRadius="20" ShadowDepth="5" Opacity="0.5" Color="Black"/>
            </Border.Effect>
            
            <StackPanel HorizontalAlignment="Center">
                <TextBlock Text="⚠️" FontSize="50" HorizontalAlignment="Center" Margin="0,0,0,10"/>
                
                <TextBlock Name="TxtMessage" Text="$Message" FontSize="18" FontWeight="Bold" 
                           Foreground="#333333" TextWrapping="Wrap" HorizontalAlignment="Center" 
                           TextAlignment="Center" Margin="0,0,0,25"/>
                
                <Button Name="BtnOK" Content="OK" Width="120" Height="35" IsDefault="True"
                        Background="#DC3545" Foreground="White" FontSize="14" FontWeight="Bold"
                        Cursor="Hand">
                    <Button.Resources>
                        <Style TargetType="Button">
                            <Setter Property="Template">
                                <Setter.Value>
                                    <ControlTemplate TargetType="Button">
                                        <Border Background="{TemplateBinding Background}" CornerRadius="4">
                                            <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
                                        </Border>
                                    </ControlTemplate>
                                </Setter.Value>
                            </Setter>
                        </Style>
                    </Button.Resources>
                </Button>
            </StackPanel>
        </Border>
    </Grid>
</Window>
"@

$reader = [System.Xml.XmlReader]::Create([System.IO.StringReader]$xml)
$window = [System.Windows.Markup.XamlReader]::Load($reader)

$bc = New-Object System.Windows.Media.BrushConverter
$window.Background = $bc.ConvertFromString($Color)

$btnOK = $window.FindName("BtnOK")
$btnOK.Add_Click({
    $window.Close()
})

$window.ShowDialog() | Out-Null