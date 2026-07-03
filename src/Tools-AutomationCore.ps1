# PowerShell UIautomation helper functions
# last modified: 2026-06-11

# アセンブリのロード-

Add-Type -AssemblyName "UIAutomationClient"
Add-Type -AssemblyName "UIAutomationTypes"
Add-Type -AssemblyName WindowsBase

# UIAutomationClientProvidersをロードする必要があるが、PowerShell単体ではロードできないので
# C#コードをAdd-Typeでコンパイルして、UIAutomationClientの型を参照してUIAutomationClientProvidersの機能を呼び出す
$sourceGetMainWindow = @'
using System;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using System.Threading;
using System.Windows;
using System.Windows.Automation;

public class UIATools
{
        public enum ClickType {
            Left,
            Right,
            Double
        }

        // Win32 structs
        [StructLayout(LayoutKind.Sequential)]
        struct MOUSEINPUT {
            public int dx;
            public int dy;
            public uint mouseData;
            public uint dwFlags;
            public uint time;
            public IntPtr dwExtraInfo;
        }

        [StructLayout(LayoutKind.Sequential)]
        struct KEYBDINPUT {
            public ushort wVk;
            public ushort wScan;
            public uint dwFlags;
            public uint time;
            public IntPtr dwExtraInfo;
        }

        [StructLayout(LayoutKind.Explicit)]
        struct INPUT {
            [FieldOffset(0)] public int type;
            //
            [FieldOffset(8)] public MOUSEINPUT mi;
            [FieldOffset(8)] public KEYBDINPUT ki;
            // 64bit環境でのサイズ調整として、FieldOffsetを8にしている
            // 32bit環境ではFieldOffset(4)にする必要がある
        }

        // Import Win32 APIs
        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool IsWindow(IntPtr hWnd);

        [DllImport("user32.dll")]
        static extern IntPtr GetForegroundWindow();

        [DllImport("user32.dll")]
        static extern bool SetForegroundWindow(IntPtr hWnd);

        [DllImport("user32.dll")]
        static extern void SwitchToThisWindow(IntPtr hWnd, bool fAltTab);

        [DllImport("user32.dll")]
        static extern bool PostMessage(IntPtr hWnd, uint Msg, IntPtr wParam, IntPtr lParam);

        [DllImport("user32.dll")]
        static extern bool ScreenToClient(IntPtr hWnd, ref System.Drawing.Point lpPoint);

        [DllImport("user32.dll")]
        static extern uint SendInput(uint nInputs, INPUT[] pInputs, int cbSize);

        [DllImport("user32.dll")]
        static extern int GetSystemMetrics(int nIndex);

        [DllImport("imm32.dll")]
        static extern IntPtr ImmGetContext(IntPtr hWnd);

        [DllImport("imm32.dll")]
        static extern bool ImmSetOpenStatus(IntPtr hIMC, bool fOpen);

        [DllImport("imm32.dll")]
        static extern bool ImmReleaseContext(IntPtr hWnd, IntPtr hIMC);

        private const uint WK_LBUTTON     = 0x0001;
        private const uint WM_LBUTTONDOWN = 0x0201;
        private const uint WM_LBUTTONUP   = 0x0202;

        private const uint MOUSEEVENTF_MOVE      = 0x0001;
        private const uint MOUSEEVENTF_LEFTDOWN  = 0x0002;
        private const uint MOUSEEVENTF_LEFTUP    = 0x0004;
        private const uint MOUSEEVENTF_RIGHTDOWN = 0x0008;
        private const uint MOUSEEVENTF_RIGHTUP   = 0x0010;
        private const uint MOUSEEVENTF_ABSOLUTE  = 0x8000;

        private const uint KEYEVENTF_KEYDOWN = 0x0000;
        private const uint KEYEVENTF_KEYUP   = 0x0002;

        private const int SM_CXSCREEN = 0;
        private const int SM_CYSCREEN = 1;

        public static AutomationElement RootElement
        {
            get
            {
                return AutomationElement.RootElement;
            }
        }

        // AutomationElementが有効かを確認するユーティリティ関数
        public static bool IsAlive(AutomationElement element) {
            if (element == null) return false;

            try {
                object hwndProp = element.GetCurrentPropertyValue(AutomationElement.NativeWindowHandleProperty);

                if (hwndProp != null) {
                    int hwndInt = (int)hwndProp;
                    if (hwndInt != 0) {
                        IntPtr hWnd = (IntPtr)hwndInt;
                        return IsWindow(hWnd);
                    }
                }
            } catch (ElementNotAvailableException) {
                // elementが既に消滅している
                return false;
            }
            return false;
        }

        // 名前にワイルドカード(*)を使用してトップレベルウインドウを取得するユーティリティ関数
        public static AutomationElement GetMainWindowByName(string name) {
            if (string.IsNullOrEmpty(name)) {
                return null;
            }

            string escaped = Regex.Escape(name).Replace("\\*", ".*");
            Regex wildcardRegex = new Regex("^" + escaped + "$", RegexOptions.IgnoreCase);
            Condition cond = new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Window);
            AutomationElementCollection windows = RootElement.FindAll(TreeScope.Element | TreeScope.Children, cond);
            foreach (AutomationElement window in windows) {
                if (wildcardRegex.IsMatch(window.Current.Name ?? string.Empty)) {
                    return window;
                }
            }
            return null;
        }

        // プロセスIDからトップレベルウインドウを取得するユーティリティ関数
        public static AutomationElement GetMainWindowByProcessID(int processId) {
            if (processId <= 0) {
                return null;
            }

            Condition cond = new AndCondition(
                new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Window),
                new PropertyCondition(AutomationElement.ProcessIdProperty, processId)
            );
            return RootElement.FindFirst(TreeScope.Element | TreeScope.Children, cond);
        }

        // オートメーションIDからトップレベルウインドウを取得するユーティリティ関数
        public static AutomationElement GetMainWindowByAutomationID(string automationId) {
            if (string.IsNullOrEmpty(automationId)) {
                return null;
            }

            Condition cond = new AndCondition(
                new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Window),
                new PropertyCondition(AutomationElement.AutomationIdProperty, automationId)
            );
            return RootElement.FindFirst(TreeScope.Element | TreeScope.Children, cond);
        }

        // UIオートメーションの要素をクリックするユーティリティ関数。
        // 背面にあってもコントロールに直接メッセージを送信する。
        // 通常のInvokeやPatternの呼び出しでうまくいかない場合に使用する。
        public static void PostClick(AutomationElement targetElement, AutomationElement windowElement) {
            if (targetElement == null || windowElement == null) throw new Exception("オブジェクトが指定されていません.");

            IntPtr hWnd = (IntPtr)windowElement.Current.NativeWindowHandle;
            if (hWnd == IntPtr.Zero) throw new Exception("親Windowsのハンドルが取得できません.");

            // ClickablePointを取得
            Point clickablePoint = targetElement.GetClickablePoint();
            if (clickablePoint == null) throw new Exception("ClickablePointが取得できません.");
            var clickPoint = new System.Drawing.Point((int)clickablePoint.X, (int)clickablePoint.Y);

            // 相対ポイントに変換
            System.Drawing.Point clientPoint = clickPoint;
            ScreenToClient(hWnd, ref clientPoint);

            // メッセージの送信
            IntPtr lParam = (IntPtr)((clientPoint.Y << 16) | (clientPoint.X & 0xffff));
            IntPtr wParam = (IntPtr)WK_LBUTTON;

            PostMessage(hWnd, WM_LBUTTONDOWN, wParam, lParam);
            PostMessage(hWnd, WM_LBUTTONUP, wParam, lParam);
        }

        // UIオートメーションの要素を強制的にクリックするユーティリティ関数。
        // フォーカスもできないようなUI要素に対して、座標を指定してマウスクリックを送ることで操作する。
        // 通常のInvokeやPatternの呼び出しでうまくいかない場合に使用する。
        // コントロールを含むウインドウが最前面にないと正確にクリックできないため、事前にウインドウを最前面にするなどの対策が必要。
        public static void ForceClick(AutomationElement element) {
            ForceClick(element, ClickType.Left);
        }

        public static void ForceClick(AutomationElement element, ClickType clickType) {
            if (element == null) throw new Exception("オブジェクトが指定されていません.");

            // ClickablePointを取得
            Point clickablePoint = element.GetClickablePoint();
            if (clickablePoint == null) throw new Exception("ClickablePointが取得できません.");

            // 絶対座標に変換
            int screenWidth = GetSystemMetrics(SM_CXSCREEN);
            int screenHeight = GetSystemMetrics(SM_CYSCREEN);

            int absX = ((int)clickablePoint.X) * 65536 / screenWidth;
            int absY = ((int)clickablePoint.Y) * 65536 / screenHeight;

            // SendInput入力の生成と送信
            INPUT[] mouseInputs = new INPUT[1];
            mouseInputs[0].type = 0; // INPUT_MOUSE
            mouseInputs[0].mi.dx = absX;
            mouseInputs[0].mi.dy = absY;
            mouseInputs[0].mi.dwFlags = MOUSEEVENTF_MOVE | MOUSEEVENTF_ABSOLUTE;
            SendInput(1, mouseInputs, Marshal.SizeOf(new INPUT()));

            switch (clickType) {
                case ClickType.Left:
                    SendMouseInput(mouseInputs, MOUSEEVENTF_LEFTDOWN);
                    SendMouseInput(mouseInputs, MOUSEEVENTF_LEFTUP);
                    break;
                case ClickType.Right:
                    SendMouseInput(mouseInputs, MOUSEEVENTF_RIGHTDOWN);
                    SendMouseInput(mouseInputs, MOUSEEVENTF_RIGHTUP);
                    break;
                case ClickType.Double:
                    SendMouseInput(mouseInputs, MOUSEEVENTF_LEFTDOWN);
                    SendMouseInput(mouseInputs, MOUSEEVENTF_LEFTUP);
                    SendMouseInput(mouseInputs, MOUSEEVENTF_LEFTDOWN);
                    SendMouseInput(mouseInputs, MOUSEEVENTF_LEFTUP);
                    break;
                default:
                    throw new ArgumentOutOfRangeException("clickType", clickType, null);
            }
        }

        private static void SendMouseInput(INPUT[] mouseInputs, uint mouseEventFlag) {
            mouseInputs[0].mi.dwFlags = mouseEventFlag | MOUSEEVENTF_ABSOLUTE;
            SendInput(1, mouseInputs, Marshal.SizeOf(new INPUT()));
        }

        // IMEの状態をoffにする
        // 最前面のウインドウに限るのでコントロールにSetFocus()してウインドウを前面に移動しておく対応が必要。
        public static void DisableIME() {
            IntPtr hWnd = GetForegroundWindow();
            if (hWnd == IntPtr.Zero) return;

            // IMEコンテキストの取得
            IntPtr hIMC = ImmGetContext(hWnd);
            if (hIMC != IntPtr.Zero) {
                // IME off
                ImmSetOpenStatus(hIMC, false);

                // IMEコンテキストを開放
                ImmReleaseContext(hWnd, hIMC);
            }
        }

        // デフォルトは 200ms の待機を提供するユーティリティ関数
        public static void Sleep() {
            Thread.Sleep(200);
        }

        public static void Sleep(int milliseconds) {
            Thread.Sleep(milliseconds);
        }

        public static void Sleep(int? milliseconds, int? seconds) {
            if (seconds.HasValue && milliseconds.HasValue) {
                throw new Exception("-Seconds と -Milliseconds は同時に指定できません");
            }

            if (seconds.HasValue) {
                Thread.Sleep(seconds.Value * 1000);
                return;
            }

            if (milliseconds.HasValue) {
                Thread.Sleep(milliseconds.Value);
                return;
            }

            Thread.Sleep(200);
        }
}
'@

if (-not ([System.Management.Automation.PSTypeName]'UIATools').Type) {
    Add-Type -TypeDefinition $sourceGetMainWindow -Language CSharp -ReferencedAssemblies("UIAutomationClient", "UIAutomationTypes", "System.Drawing", "WindowsBase", "System.Threading.Thread")
}

# --- [1] 汎用検索エンジン ---
function Get-UIAControls {
    <#
    .SYNOPSIS
        条件に一致するすべてのUI要素を配列で取得します。Nameにワイルドカード(*)を使用可能です。
    #>
    param(
        [Parameter(Mandatory)][Windows.Automation.AutomationElement]$Parent,
        [string]$Id,
        [string]$Name,
        [Windows.Automation.ControlType]$Type,
        [Windows.Automation.TreeScope]$Scope = ([Windows.Automation.TreeScope]::Descendants),
        [Windows.Automation.Condition]$Condition # 直接Conditionを渡す場合に使用
    )

    # Write-Host 'Get-UIAControls:'

    if ($null -eq $Parent) {
        throw "Parent要素は必須です"
    }

    $hasWildcardName = (-not [string]::IsNullOrEmpty($Name)) -and $Name.Contains('*')
    
    $searchCond = $Condition
    if ($null -eq $searchCond) {
        $conditions = @()
        if ($Id) { 
            $conditions += New-Object Windows.Automation.PropertyCondition([Windows.Automation.AutomationElement]::AutomationIdProperty, $Id)
        }
        if ($Name -and -not $hasWildcardName) {
            $conditions += New-Object Windows.Automation.PropertyCondition([Windows.Automation.AutomationElement]::NameProperty, $Name)
        }
        if ($Type) {
            $conditions += New-Object Windows.Automation.PropertyCondition([Windows.Automation.AutomationElement]::ControlTypeProperty, $Type)
        }

        $searchCond = if ($conditions.Count -eq 0) { [Windows.Automation.Condition]::TrueCondition }
                     elseif ($conditions.Count -gt 1) { New-Object Windows.Automation.AndCondition(,[Windows.Automation.PropertyCondition[]]$conditions) }
                     else { $conditions[0] }
    }

    $elements = $Parent.FindAll($Scope, $searchCond)

    # ワイルドカードのName指定がある場合は、Nameプロパティをフィルタリングする
    if ($hasWildcardName) {
        $elements = $elements | Where-Object {
            $_.Current.Name -like $Name
        }
    }
    return @($elements)
}

function Get-UIAControl {
    <#
    .SYNOPSIS
        条件に一致する最初のUI要素を取得します。Nameにワイルドカード(*)を使用可能です。
        ワイルドカードを使用した場合はGet-UIAControls後にフィルタリング(FindFirst相当)で動作するのでパフォーマンスがやや落ちるので注意。
    #>
    param(
        [Parameter(Mandatory)][Windows.Automation.AutomationElement]$Parent, 
        [string]$Id, 
        [string]$Name, 
        [Windows.Automation.ControlType]$Type, 
        [Windows.Automation.TreeScope]$Scope = [Windows.Automation.TreeScope]::Descendants,
        [Windows.Automation.Condition]$Condition,
        [int]$TimeoutSec = 10
    )

    if ($null -eq $Parent) {
        throw "Parent要素は必須です"
    }

    # ワイルドカードがある場合は、Get-UIAControlsの結果から最初の要素を取得する。なければFindFirstで高速に取得する。
    $hasWildcardName = ($null -ne $Name) -and $Name.Contains('*')
    if ($hasWildcardName) {
        return Get-UIAControls -Parent $Parent -Id $Id -Name $Name -Type $Type -Condition $Condition -Scope $Scope | Select-Object -First 1
    }

    # Conditionが未指定の場合のみ、Id, NameとTypeから組み立てる
    $searchCond = $Condition
    if ($null -eq $searchCond) {
        $conditions = @()
        if ($Id) {
            $conditions += New-Object Windows.Automation.PropertyCondition([Windows.Automation.AutomationElement]::AutomationIdProperty, $Id)
        }
        if ($Name) {
            $conditions += New-Object Windows.Automation.PropertyCondition([Windows.Automation.AutomationElement]::NameProperty, $Name)
        }
        if ($Type) {
            $conditions += New-Object Windows.Automation.PropertyCondition([Windows.Automation.AutomationElement]::ControlTypeProperty, $Type)
        }
        
        if ($conditions.Count -gt 1) {
            # AndCondition は配列を 1 引数で渡す必要がある（先頭のカンマで配列を保持）
            $searchCond = New-Object Windows.Automation.AndCondition(,[Windows.Automation.PropertyCondition[]]$conditions)
        } elseif ($conditions.Count -eq 1) {
            $searchCond = $conditions[0]
        } else {
            $searchCond = [Windows.Automation.Condition]::TrueCondition
        }
    }

    $stopwatch = [System.Diagnostics.Stopwatch]::StartNew()
    while ($stopwatch.Elapsed.TotalSeconds -lt $TimeoutSec) {
        # write-host 'Finding'
        $element = $Parent.FindFirst($Scope, $searchCond)
        if ($null -ne $element) { return $element }
        [UIATools]::Sleep(300)
    }
    throw "要素取得タイムアウト: Name=$Name, Id=$Id"
}

# --- [2] コントロール別 取得関数 ---

# デスクトップ直下からアプリウィンドウを探す（名前またはPID指定）
function Get-UIAAppWindow { 
    <#
    .SYNOPSIS
        デスクトップ直下からアプリウィンドウを探します。名前, オートメーションIDまたはPIDで指定可能です。
        GetMainWindow*のラッパー。
    #>
    param(
        [string]$Name,
        [string]$Id,
        [int]$ProcessId = 0
    ) 

    # 1. 確実に判定できるPIDがある場合はそれで対応
    if ($ProcessId -gt 0) { 
        return [UIATools]::GetMainWindowByProcessId($ProcessId)
    }

    # 2. AutomationIdで検索
    if ($Id) {
        return [UIATools]::GetMainWindowByAutomationID($Id)
    }

    # 3. Nameで検索
    if ($Name) {
        return [UIATools]::GetMainWindowByName($Name)
    }

    return $null
}

# 任意の親要素（Windowなど）の下にある子要素を探す
function Get-UIAWindow      { param($Parent, $Id, $Name) Get-UIAControl -Parent $Parent -Id $Id -Name $Name -Type ([Windows.Automation.ControlType]::Window) }
function Get-UIAWindows     { param($Parent, $Id, $Name) Get-UIAControls -Parent $Parent -Id $Id -Name $Name -Type ([Windows.Automation.ControlType]::Window) }
function Get-UIAChildWindows { param($Parent, $Id, $Name) Get-UIAControls -Parent $Parent -Id $Id -Name $Name -Type ([Windows.Automation.ControlType]::Window) -Scope ([Windows.Automation.TreeScope]::Children) }

function Get-UIAPane        { param($Parent, $Id, $Name) Get-UIAControl -Parent $Parent -Id $Id -Name $Name -Type ([Windows.Automation.ControlType]::Pane) }
function Get-UIAPanes       { param($Parent, $Id, $Name) Get-UIAControls -Parent $Parent -Id $Id -Name $Name -Type ([Windows.Automation.ControlType]::Pane) }
function Get-UIAChildPane   { param($Parent, $Id, $Name) Get-UIAControl -Parent $Parent -Id $Id -Name $Name -Type ([Windows.Automation.ControlType]::Pane) -Scope ([Windows.Automation.TreeScope]::Children) }
function Get-UIAChildPanes  { param($Parent, $Id, $Name) Get-UIAControls -Parent $Parent -Id $Id -Name $Name -Type ([Windows.Automation.ControlType]::Pane) -Scope ([Windows.Automation.TreeScope]::Children) }

function Get-UIAButton      { param($Parent, $Id, $Name) Get-UIAControl -Parent $Parent -Id $Id -Name $Name -Type ([Windows.Automation.ControlType]::Button) }
function Get-UIAButtons     { param($Parent, $Id, $Name) Get-UIAControls -Parent $Parent -Id $Id -Name $Name -Type ([Windows.Automation.ControlType]::Button) }
function Get-UIAText        { param($Parent, $Id, $Name) Get-UIAControl -Parent $Parent -Id $Id -Name $Name -Type ([Windows.Automation.ControlType]::Text) }
function Get-UIADocument    { param($Parent, $Id, $Name) Get-UIAControl -Parent $Parent -Id $Id -Name $Name -Type ([Windows.Automation.ControlType]::Document) }
function Get-UIAEdit        { param($Parent, $Id, $Name) Get-UIAControl -Parent $Parent -Id $Id -Name $Name -Type ([Windows.Automation.ControlType]::Edit) }
function Get-UIACheckBox    { param($Parent, $Id, $Name) Get-UIAControl -Parent $Parent -Id $Id -Name $Name -Type ([Windows.Automation.ControlType]::CheckBox) }
function Get-UIAComboBox    { param($Parent, $Id, $Name) Get-UIAControl -Parent $Parent -Id $Id -Name $Name -Type ([Windows.Automation.ControlType]::ComboBox) }
function Get-UIAList        { param($Parent, $Id, $Name) Get-UIAControl -Parent $Parent -Id $Id -Name $Name -Type ([Windows.Automation.ControlType]::List) }
function Get-UIAListItems   { param($Parent, $Id, $Name) Get-UIAControls -Parent $Parent -Id $Id -Name $Name -Type ([Windows.Automation.ControlType]::ListItem) }
function Get-UIAGroup       { param($Parent, $Id, $Name) Get-UIAControl -Parent $Parent -Id $Id -Name $Name -Type ([Windows.Automation.ControlType]::Group) }
function Get-UIARadio       { param($Parent, $Id, $Name) Get-UIAControl -Parent $Parent -Id $Id -Name $Name -Type ([Windows.Automation.ControlType]::RadioButton) }
function Get-UIARadios      { param($Parent, $Id, $Name) Get-UIAControl -Parent $Parent -Id $Id -Name $Name -Type ([Windows.Automation.ControlType]::RadioButton) }
function Get-UIAChildRadios { param($Parent, $Id, $Name) Get-UIAControls -Parent $Parent -Id $Id -Name $Name -Type ([Windows.Automation.ControlType]::RadioButton) -Scope ([Windows.Automation.TreeScope]::Children) }
function Get-UIAMenuBar     { param($Parent, $Id, $Name) Get-UIAControl -Parent $Parent -Id $Id -Name $Name -Type ([Windows.Automation.ControlType]::MenuBar) }
function Get-UIAChildMenuBar{ param($Parent, $Id, $Name) Get-UIAControl -Parent $Parent -Id $Id -Name $Name -Type ([Windows.Automation.ControlType]::MenuBar) -Scope ([Windows.Automation.TreeScope]::Children) }
function Get-UIAMenuItem    { param($Parent, $Id, $Name) Get-UIAControl -Parent $Parent -Id $Id -Name $Name -Type ([Windows.Automation.ControlType]::MenuItem) }
function Get-UIADataGrid    { param($Parent, $Id, $Name) Get-UIAControl -Parent $Parent -Id $Id -Name $Name -Type ([Windows.Automation.ControlType]::DataGrid) }
function Get-UIADataItems   { param($Parent, $Id, $Name) Get-UIAControls -Parent $Parent -Id $Id -Name $Name -Type ([Windows.Automation.ControlType]::DataItem) }
function Get-UIATable       { param($Parent, $Id, $Name) Get-UIAControl -Parent $Parent -Id $Id -Name $Name -Type ([Windows.Automation.ControlType]::Table) }

# --- [3] 操作・ユーティリティ関数 ---

function Invoke-UIAElement {
    <#
    .SYNOPSIS
        UIオートメーションの要素を操作します。InvokePatternまたはTogglePatternを使用します。
    #>
    param([Parameter(Mandatory, Position = 0)]
        [Windows.Automation.AutomationElement]$Element
    )

    $pattern = $null
    if ($Element.TryGetCurrentPattern([Windows.Automation.InvokePattern]::Pattern, [ref]$pattern)) {
        $pattern.Invoke()
    } elseif ($Element.TryGetCurrentPattern([Windows.Automation.TogglePattern]::Pattern, [ref]$pattern)) {
        $pattern.Toggle()
    } elseif ($Element.TryGetCurrentPattern([Windows.Automation.SelectionItemPattern]::Pattern, [ref]$pattern)) {
        $pattern.Select()
    } else {
        throw "Invoke,Toggle,Select非対応の要素です"
    }
}

function Invoke-UIAForceClick {
    <#
    .SYNOPSIS
        UIオートメーションの要素をClickします。
        PaneなどではInvokePatternやTogglePatternが動作しないことが多いので、ForceClickで直接クリックする。
        コントロールを含むウインドウが最前面にないと正確にクリックできないため、
        事前にウインドウを最前面にするなどの対策が必要なので $ParentWindow パラメーターも必須にする。
    #>
    param(
        [Parameter(Mandatory, Position = 0)] [Windows.Automation.AutomationElement]$Element,
        [Parameter(Mandatory, Position = 1)] [Windows.Automation.AutomationElement]$ParentWindow
    )
    if ($null -eq $Element -or $null -eq $ParentWindow) {
        throw 'Element要素, ParentWindow要素は必須です'
    }

    # 親Windowが最小化していないことが大前提
    $windowPattern = $null
    if ($ParentWindow.TryGetCurrentPattern([Windows.Automation.WindowPatternIdentifiers]::Pattern, [ref]$windowPattern)) {
        $windowPattern.SetWindowVisualState('Normal')
    } else {
        throw '無効なParentWindowです'
    }

    [UIATools]::ForceClick($Element)
}

function Set-UIAElementExpanded {
    <#
    .SYNOPSIS
        UIオートメーションの要素を展開します。ExpandCollapsePatternを使用します。
    #>
    param([Parameter(Mandatory, Position = 0)]
        [Windows.Automation.AutomationElement]$Element
    )

    $expandCollapsePattern = $null
    if ($Element.TryGetCurrentPattern([Windows.Automation.ExpandCollapsePattern]::Pattern, [ref]$expandCollapsePattern)) {
        $expandCollapsePattern.Expand()
    } else {
        throw "展開非対応の要素です"
    }
}

function Set-UIAElementCollapsed {
    <#
    .SYNOPSIS
        UIオートメーションの要素を折りたたみます。ExpandCollapsePatternを使用します。
    #>
    param([Parameter(Mandatory, Position = 0)]
        [Windows.Automation.AutomationElement]$Element
    )

    $expandCollapsePattern = $null
    if ($Element.TryGetCurrentPattern([Windows.Automation.ExpandCollapsePattern]::Pattern, [ref]$expandCollapsePattern)) {
        $expandCollapsePattern.Collapse()
    } else {
        throw "折りたたみ非対応の要素です"
    }
}

function Set-UIAWindowActive {
    <#
    .SYNOPSIS
        UIオートメーションのウィンドウを最前面にします。WindowPatternを使用します。
    #>
    param([Parameter(Mandatory, Position = 0)]
        [Windows.Automation.AutomationElement]$Element
    )

    if ($Element.Current.ControlType -ne [Windows.Automation.ControlType]::Window) {
        throw "ウィンドウ以外の要素を操作できません"
    }

    $windowPattern = $null
    if ($Element.TryGetCurrentPattern([Windows.Automation.WindowPattern]::Pattern, [ref]$windowPattern)) {
        $windowPattern.SetWindowVisualState([Windows.Automation.WindowVisualState]::Normal)
        $Element.SetFocus()
    } else {
        throw "ウィンドウステート更新非対応の要素です"
    }
}

function Close-UIAWindow {
    <#
    .SYNOPSIS
        UIオートメーションのウィンドウを閉じます。WindowPatternを使用します。
    #>
    param([Parameter(Mandatory, Position = 0)]
        [Windows.Automation.AutomationElement]$Element
    )

    if ($Element.Current.ControlType -ne [Windows.Automation.ControlType]::Window) {
        throw "ウィンドウ以外の要素を閉じることはできません"
    }

    $windowPattern = $null
    if ($Element.TryGetCurrentPattern([Windows.Automation.WindowPattern]::Pattern, [ref]$windowPattern)) {
        $windowPattern.Close()
    } else {
        throw "ウィンドウクローズ非対応の要素です"
    }
}

function Get-UIAName {
    <#
    .SYNOPSIS
        要素の Name プロパティ（Windowのタイトル、Buttonのテキスト、Textラベルなど）を取得します。
    #>
    param(
        [Parameter(Mandatory, Position = 0)]
        [Windows.Automation.AutomationElement]$Element
    )

    return $Element.Current.Name
}

function Get-UIAValue {
    <#
    .SYNOPSIS
        Text, Edit, ListItem などの要素から表示されている文字列を取得します。
    #>
    param(
        [Parameter(Mandatory, Position = 0)]
        [Windows.Automation.AutomationElement]$Element
    )

    # 1. ValuePattern を持っている場合 (Edit, ComboBoxなど)
    $pattern = $null
    if ($Element.TryGetCurrentPattern([Windows.Automation.ValuePattern]::Pattern, [ref]$pattern)) {
        return $pattern.Current.Value
    }
    
    # 2. CheckBoxの場合
    if ($Element.TryGetCurrentPattern([Windows.Automation.TogglePattern]::Pattern, [ref]$pattern)) {
        return ($pattern.Current.ToggleState -eq [Windows.Automation.ToggleState]::On)
    }

    # 3. RadioButtonの場合
    if ($Element.TryGetCurrentPattern([Windows.Automation.SelectionItemPattern]::Pattern, [ref]$pattern)) {
        return $pattern.Current.IsSelected
    }

    # 4. Pattern がない場合 (Text, Button, ListItemなど)
    # 多くの場合は Name プロパティに文字列が入っている - Get-UIAName を呼び出す
    return Get-UIAName $Element
}

function Set-UIAValue {
    <#
    .SYNOPSIS
        UIオートメーションの要素に値を設定します。基本的にValuePatternを使用します。
        -ForceでSendWaitで強制的にコントロールに送信します。
    #>
    param(
        [Parameter(Mandatory, Position = 0)] [Windows.Automation.AutomationElement]$Element,
        [Parameter(Mandatory, Position = 1)] [string]$Value,
        [switch]$Force = $false,
        [switch]$OmitEscape = $false
    )

    if ($Force -eq $false) {
        $valuePattern = $null
        $Element.TryGetCurrentPattern([Windows.Automation.ValuePattern]::Pattern, [ref]$valuePattern)
        if ($null -eq $valuePattern) {
            throw "値入力非対応の要素です"
        }
        $valuePattern.SetValue($Value)
    } else {
        # コントロールにSendKey.SendWaitで送信する
        try {
            $Element.SetFocus()
            [UIATools]::Sleep(50)
        } catch {}

        if ($OmitEscape) {
            [System.Windows.Forms.SendKeys]::SendWait($Value)
        } else {
            # SendKeysのコントロール文字列をエスケープして送信
            [System.Windows.Forms.SendKeys]::SendWait(($Value -replace '([+^%~(){}\[\]])','{$1}'))
        }
    }
}

function Get-UIAComboBoxItems {
    <#
    .SYNOPSIS
        コンボボックスを展開し、内部のアイテム一覧を String[] として取得します。
    #>
    param([Parameter(Mandatory, Position = 0)] [Windows.Automation.AutomationElement]$ComboBox)

    $expandPattern = $true
    try {
        Set-UIAElementExpanded -Element $ComboBox
        [UIATools]::Sleep() # 展開待ち
    } catch {
        $expandPattern = $false
    }

    $listItems = Get-UIAControls -Parent $ComboBox -Type ([Windows.Automation.ControlType]::ListItem) -Scope Descendants
    $itemNames = foreach ($item in $listItems) { Get-UIAName $item }

    if ($expandPattern) {
        Set-UIAElementCollapsed -Element $ComboBox
    }
    return [string[]]$itemNames
}

function Set-UIAComboBoxValue {
    <#
    .SYNOPSIS
        コンボボックスの値を下位のTextに対して設定する。
    #>
    param(
        [Parameter(Mandatory, Position = 0)] [Windows.Automation.AutomationElement]$Element,
        [Parameter(Mandatory, Position = 1)] [string]$Value
    )

    if ($null -eq $Element -or $Element.Current.ControlType -ne [Windows.Automation.ControlType]::ComboBox) {
        throw 'コンボボックスを指定して下さい'
    }

    # SetTextが可能ならそれで実施する
    try {
        Set-UIAValue $Element $Value
        return
    } catch {}

    # SetText出来ない場合はリストとして項目を列挙して選択する
    Set-UIAListSelection $Element $Value
}

function Set-UIAListSelection {
    <#
    .SYNOPSIS
        リストコントロールのアイテムを選択します。
    #>
    param(
        [Parameter(Mandatory, Position = 0)] [Windows.Automation.AutomationElement]$List,
        [Parameter(Mandatory, Position = 1)] [string]$Value,
        [switch] [boolean]$UseMatch = $false
    )

    $listItems = Get-UIAListItems -Parent $List
    if ($listItems.Count -eq 0) {
        throw '選択対象がありません'
    }

    $selectionPattern = $null
    foreach ($item in $listItems) {
        if ($UseMatch) {
            if ($item.Current.Name -match $Value) {
                if ($item.TryGetCurrentPattern([Windows.Automation.SelectionItemPattern]::Pattern, [ref]$selectionPattern)) {
                    $selectionPattern.Select()
                    return
                }
            }
        } else {
            if ($item.Current.Name -eq $Value) {
                if ($item.TryGetCurrentPattern([Windows.Automation.SelectionItemPattern]::Pattern, [ref]$selectionPattern)) {
                    $selectionPattern.Select()
                    return
                }
            }
        }
    }
    throw '値を選択できません'
}

function Get-UIAListItemValues {
    <#
    .SYNOPSIS
        リストコントロールのアイテム一覧を String[] として取得します。
    #>
    param(
        [Parameter(Mandatory, Position = 0)] [Windows.Automation.AutomationElement]$List
    )

    $listItems = Get-UIAListItems -Parent $List
    return [string[]]($listItems | ForEach-Object { $_.Current.Name })
}

function Get-UIATableContents {
    <#
    .SYNOPSIS
        TableやDataGridの内容を headerとvaluesで取得します。
        ヘッダーが無い場合、headerは空配列になります。
    #>
    param(
        [Parameter(Mandatory, Position = 0)]
        [Alias('Table', 'DataGrid')]
        [Windows.Automation.AutomationElement]$Element
    )

    $supportedControlTypes = @(
        [Windows.Automation.ControlType]::Table,
        [Windows.Automation.ControlType]::DataGrid
    )
    if ($supportedControlTypes -notcontains $Element.Current.ControlType) {
        throw "Get-UIATableContents は Table または DataGrid 要素のみ対応しています"
    }

    $headers = Get-UIAControls -Parent $Element -Type ([Windows.Automation.ControlType]::Header) -Scope ([Windows.Automation.TreeScope]::Descendants)
    $headerNames = foreach ($header in $headers) { Get-UIAValue $header }
    $rows = Get-UIAControls -Parent $Element -Type ([Windows.Automation.ControlType]::DataItem) -Scope ([Windows.Automation.TreeScope]::Descendants)
    $values = foreach ($row in $rows) {
        # 多くの実装では DataItem の直下がセルだが、実装差異を考慮して子要素優先 + 子孫要素へフォールバックする
        $cells = Get-UIAControls -Parent $row -Scope [Windows.Automation.TreeScope]::Children -Condition ([Windows.Automation.Condition]::TrueCondition)
        if ($cells.Count -eq 0) {
            $cells = Get-UIAControls -Parent $row -Scope [Windows.Automation.TreeScope]::Descendants -Condition ([Windows.Automation.Condition]::TrueCondition)
        }

        $cellValues = foreach ($cell in $cells) {
            Get-UIAValue $cell
        }
        ,$cellValues # 2次元配列にするため、行ごとにカンマで配列化
    }
    return @{
        Header = [string[]]$headerNames
        Values = [string[][]]$values
    }
}

# --- [4] UIオートメーション複合関数 ---
function GetAppWindow {
    param(
        [Parameter(Mandatory)] [string]$Name,
        [string]$ExecutablePath
    )

    $window = $null
    # 既に起動していたらウィンドウを取得して返す
    try {
        # Write-Host "Check Window $($Name)"
        $window = Get-UIAAppWindow -Name $Name
    } catch {}

    if ($null -ne $window) {
        return $window
    }

    # 見つからない場合は新規起動する
    if ($ExecutablePath -and (Test-Path $ExecutablePath)) {
        # write-host "アプリを起動します: $ExecutablePath"
        $process = Start-Process -FilePath $ExecutablePath -PassThru
        [UIATools]::Sleep(5000) # 起動待ち
        try {
            $window = Get-UIAAppWindow -ProcessId $process.Id
        } catch {}
    } else {
        # Write-Host "Executable $($ExecutablePath) is not found"
    }
    return $window
}

function FindByTreewalker {
    param(
        [Windows.Automation.AutomationElement]$element,
        [Windows.Automation.ControlType]$type,
        $walker = $null
    )

    if ($element.Current.ControlType -eq $type) {
        return $element
    }

    if ($null -eq $walker) {
        $walker = [Windows.Automation.TreeWalker]::ControlViewWalker
    }

    $child = $walker.GetFirstChild($element)

    while ($null -ne $child) {
        # 再帰呼び出しでwalkerを引き継ぐ
        $found = FindByTreewalker -element $child -type $type -walker $walker
        if ($null -ne $found) {
            return $found
        }

        $child = $walker.GetNextSibling($child)
    }

    return $null
}

# EOF
