/*
    * AutomationOverlay - オートメーション実行中に全画面で表示するオーバーレイ
    * 
    * 使い方:
    *   1. ビルドしてAutomationOverlay.exeを作成
    *   2. コマンドプロンプトやバッチファイルから以下のように実行
    *      AutomationOverlay.exe "オートメーション実行中..." "#440000FF"
    * 
    * 引数:
    *   1. 表示するメッセージ（省略可、デフォルトは「オートメーション実行中...」）
    *   2. 背景色のカラーコード（省略可、デフォルトは半透明青 #440000FF）
    * 
    * 注意:
    *   - このオーバーレイはマウスとキーボードを透過するため、下のアプリケーションは通常通り操作できます。
    *   - スリープ抑制機能も組み込まれているため、長時間のオートメーション実行に適しています。
*/
using System;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using System.Windows.Interop;
using System.Runtime.InteropServices;

namespace AutomationOverlay {
    public class Program {
        [STAThread]
        public static void Main(string[] args) {
            string message = args.Length > 0 ? args[0] : "オートメーション実行中...";
            string colorCode = args.Length > 1 ? args[1] : "#440000FF"; // デフォルト半透明青

            var app = new Application();
            var win = new Window {
                Title = "AutomationOverlay",
                WindowStyle = WindowStyle.None,
                AllowsTransparency = true,
                Topmost = true,
                WindowState = WindowState.Maximized,
                ShowInTaskbar = false,
                Background = (Brush)new System.Windows.Media.BrushConverter().ConvertFromString(colorCode),
                Content = new Grid {
                    Children = {
                        new TextBlock {
                            Text = message,
                            Foreground = Brushes.White,
                            FontSize = 60,
                            FontWeight = FontWeights.Bold,
                            HorizontalAlignment = HorizontalAlignment.Center,
                            VerticalAlignment = VerticalAlignment.Center,
                            Opacity = 0.6
                        }
                    }
                }
            };

            win.SourceInitialized += (s, e) => {
                // ウィンドウハンドル（HWND）を取得
                var hwnd = new WindowInteropHelper(win).Handle;
                
                // 現在の拡張スタイルを取得
                int extendedStyle = GetWindowLong(hwnd, GWL_EXSTYLE);
                
                // WS_EX_TRANSPARENT (マウス透過) と WS_EX_NOACTIVATE (キーボードフォーカス非アクティブ化) を追加
                SetWindowLong(hwnd, GWL_EXSTYLE, extendedStyle | WS_EX_TRANSPARENT | WS_EX_LAYERED | WS_EX_NOACTIVATE);

                // スリープ抑制
                SetThreadExecutionState(ES_CONTINUOUS | ES_DISPLAY_REQUIRED | ES_SYSTEM_REQUIRED);
            };

            win.Closed += (s, e) => {
                // スリープ抑制解除
                SetThreadExecutionState(ES_CONTINUOUS);
            };

            app.Run(win);
        }

        // Windows API
        [DllImport("user32.dll")] static extern int GetWindowLong(IntPtr hWnd, int nIndex);
        [DllImport("user32.dll")] static extern int SetWindowLong(IntPtr hWnd, int nIndex, int dwNewLong);
        [DllImport("kernel32.dll")] static extern uint SetThreadExecutionState(uint esFlags);

        const int GWL_EXSTYLE = -20;
        const int WS_EX_TRANSPARENT = 0x00000020;
        const int WS_EX_LAYERED = 0x00080000;
        const int WS_EX_NOACTIVATE = 0x08000000; // ★これを追加：アクティブ化を禁止するスタイル
        const uint ES_CONTINUOUS = 0x80000000;
        const uint ES_DISPLAY_REQUIRED = 0x00000002;
        const uint ES_SYSTEM_REQUIRED = 0x00000001;
    }
}
