param (
    [string]$Message = "自動実行中さわらないでね！",
    [string]$Color = "#440000FF" 
)

Add-Type -AssemblyName PresentationFramework, System.Windows.Forms, WindowsBase

# プライマリ画面の解像度（サイズ）を取得
$primaryScreen = [System.Windows.Forms.Screen]::PrimaryScreen
$screenWidth  = $primaryScreen.Bounds.Width
$screenHeight = $primaryScreen.Bounds.Height

$code = @"
using System;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Interop;

public class WindowHelper {
    const int GWL_EXSTYLE = -20;
    const int WS_EX_TRANSPARENT = 0x00000020;
    const int WS_EX_LAYERED = 0x00080000;
    const int WS_EX_NOACTIVATE = 0x08000000;

    [DllImport("user32.dll")]
    static extern int GetWindowLong(IntPtr hWnd, int nIndex);

    [DllImport("user32.dll")]
    static extern int SetWindowLong(IntPtr hWnd, int nIndex, int dwNewLong);

    [DllImport("kernel32.dll")]
    public static extern uint SetThreadExecutionState(uint esFlags);
    public const uint ES_CONTINUOUS = 0x80000000;
    public const uint ES_DISPLAY_REQUIRED = 0x00000002;
    public const uint ES_SYSTEM_REQUIRED = 0x00000001;

    public static void MakeTransparentAndNoActivate(Window window) {
        var hwnd = new WindowInteropHelper(window).Handle;
        int extendedStyle = GetWindowLong(hwnd, GWL_EXSTYLE);
        SetWindowLong(hwnd, GWL_EXSTYLE, extendedStyle | WS_EX_TRANSPARENT | WS_EX_LAYERED | WS_EX_NOACTIVATE);
    }
}
"@
Add-Type -TypeDefinition $code -ReferencedAssemblies "PresentationFramework", "PresentationCore", "WindowsBase", "System.Xaml"

# XAML: WindowStateを削除し、Width/Height/Left/Top をPowerShell側から埋め込む
$xml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        Title="Overlay" WindowStyle="None" AllowsTransparency="True" Topmost="True" 
        Left="0" Top="0" Width="$screenWidth" Height="$screenHeight"
        Background="$Color" ShowInTaskbar="False">
    <Grid>
        <TextBlock Text="$Message" Foreground="White" FontSize="60" FontWeight="Bold"
                   HorizontalAlignment="Center" VerticalAlignment="Center" Opacity="0.6">
            <TextBlock.Effect>
                <DropShadowEffect BlurRadius="10" ShadowDepth="0" Color="Black"/>
            </TextBlock.Effect>
        </TextBlock>
    </Grid>
</Window>
"@

$reader = [System.Xml.XmlReader]::Create([System.IO.StringReader]$xml)
$window = [System.Windows.Markup.XamlReader]::Load($reader)

$window.Add_SourceInitialized({
    [WindowHelper]::MakeTransparentAndNoActivate($this)
    [WindowHelper]::SetThreadExecutionState([WindowHelper]::ES_CONTINUOUS -bor [WindowHelper]::ES_DISPLAY_REQUIRED -bor [WindowHelper]::ES_SYSTEM_REQUIRED)
})

$window.Add_Closed({
    [WindowHelper]::SetThreadExecutionState([WindowHelper]::ES_CONTINUOUS)
})

$window.ShowDialog() | Out-Null