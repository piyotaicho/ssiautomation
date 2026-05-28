using System;
using System.Diagnostics;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Threading;
using System.Windows.Automation;

namespace SsIAutomation.Special
{
    internal static class WebLogout
    {
        private const string TargetProcessName = "Entrance";
        private const string MainWindowAutomationId = "MainFrm";
        private const string MainWindowNamePrefix = "エントランス";
        private const string BrowserPaneAutomationId = "tlpWebBrowser";
        private const string TargetTabName = "★ガルーン★";

        [STAThread]
        private static int Main()
        {
            try
            {
                if (!IsEntranceRunning())
                {
                    Console.Error.WriteLine("Entrance.exe が起動していません。");
                    return 1;
                }

                AutomationElement mainWindow = FindEntranceWindow();
                if (mainWindow == null)
                {
                    Console.Error.WriteLine("Entrance のメインウィンドウを取得できませんでした。");
                    return 1;
                }

                AutomationElement webBrowserPane = FindWebBrowserPane(mainWindow);
                if (webBrowserPane == null)
                {
                    Console.Error.WriteLine("automationId='tlpWebBrowser' の Pane を取得できませんでした。");
                    return 1;
                }

                AutomationElement tabWebContents = FindTabWebContents(webBrowserPane);
                if (tabWebContents == null)
                {
                    Console.Error.WriteLine("automationId='tabWebContents' (または 'tabWebContens') の Tab を取得できませんでした。");
                    return 1;
                }

                if (!SelectTabItem(tabWebContents, TargetTabName))
                {
                    Console.Error.WriteLine("TabItem '★ガルーン★' を選択できませんでした。");
                    return 1;
                }

                AutomationElement selectedPane = FindSelectedTabPane(tabWebContents, TargetTabName);
                if (selectedPane == null)
                {
                    Console.Error.WriteLine("Tab 配下の Pane[0] が Name='★ガルーン★' であることを確認できませんでした。");
                    return 1;
                }

                AutomationElement ieHost = FindInternetExplorerHost(selectedPane);
                if (ieHost == null)
                {
                    Console.Error.WriteLine("埋め込み Internet Explorer コントロールを取得できませんでした。");
                    return 1;
                }

                bool attempted = TryCallLogout(ieHost);
                if (attempted)
                {
                    Console.WriteLine("logout() 呼び出しを試行しました。関数未定義でもエラーにはしていません。");
                }
                else
                {
                    Console.WriteLine("logout() 呼び出し対象の DOM にアクセスできませんでしたが、エラーにはしません。");
                }

                return 0;
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine("予期しないエラー: " + ex.Message);
                return 1;
            }
        }

        private static bool IsEntranceRunning()
        {
            return Process.GetProcessesByName(TargetProcessName).Length > 0;
        }

        private static AutomationElement FindEntranceWindow()
        {
            AutomationElement root = AutomationElement.RootElement;
            if (root == null)
            {
                return null;
            }

            AutomationElementCollection windows = root.FindAll(TreeScope.Children, new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Window));
            foreach (AutomationElement window in windows)
            {
                if (window == null)
                {
                    continue;
                }

                string automationId = SafeCurrentString(window, AutomationElement.AutomationIdProperty);
                string name = SafeCurrentString(window, AutomationElement.NameProperty);
                int processId = SafeCurrentInt(window, AutomationElement.ProcessIdProperty);

                if (!string.Equals(automationId, MainWindowAutomationId, StringComparison.Ordinal))
                {
                    continue;
                }

                if (name == null || !name.StartsWith(MainWindowNamePrefix, StringComparison.Ordinal))
                {
                    continue;
                }

                Process process;
                try
                {
                    process = Process.GetProcessById(processId);
                }
                catch
                {
                    continue;
                }

                if (process == null || !string.Equals(process.ProcessName, TargetProcessName, StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                return window;
            }

            return null;
        }

        private static AutomationElement FindWebBrowserPane(AutomationElement mainWindow)
        {
            AutomationElement pane0 = GetChildByTypeAndIndex(mainWindow, ControlType.Pane, 0);
            AutomationElement pane1 = GetChildByTypeAndIndex(pane0, ControlType.Pane, 0);
            AutomationElement pane2 = GetChildByTypeAndIndex(pane1, ControlType.Pane, 0);

            if (pane2 != null)
            {
                string id = SafeCurrentString(pane2, AutomationElement.AutomationIdProperty);
                if (string.Equals(id, BrowserPaneAutomationId, StringComparison.Ordinal))
                {
                    return pane2;
                }
            }

            return mainWindow.FindFirst(
                TreeScope.Descendants,
                new AndCondition(
                    new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Pane),
                    new PropertyCondition(AutomationElement.AutomationIdProperty, BrowserPaneAutomationId)));
        }

        private static AutomationElement FindTabWebContents(AutomationElement webBrowserPane)
        {
            AutomationElement tab = webBrowserPane.FindFirst(
                TreeScope.Descendants,
                new AndCondition(
                    new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Tab),
                    new PropertyCondition(AutomationElement.AutomationIdProperty, "tabWebContents")));

            if (tab != null)
            {
                return tab;
            }

            return webBrowserPane.FindFirst(
                TreeScope.Descendants,
                new AndCondition(
                    new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Tab),
                    new PropertyCondition(AutomationElement.AutomationIdProperty, "tabWebContens")));
        }

        private static bool SelectTabItem(AutomationElement tabControl, string tabName)
        {
            AutomationElement tabItem = tabControl.FindFirst(
                TreeScope.Descendants,
                new AndCondition(
                    new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.TabItem),
                    new PropertyCondition(AutomationElement.NameProperty, tabName)));

            if (tabItem == null)
            {
                return false;
            }

            object pattern;
            if (tabItem.TryGetCurrentPattern(SelectionItemPattern.Pattern, out pattern))
            {
                ((SelectionItemPattern)pattern).Select();
                Thread.Sleep(300);
                return true;
            }

            if (tabItem.TryGetCurrentPattern(InvokePattern.Pattern, out pattern))
            {
                ((InvokePattern)pattern).Invoke();
                Thread.Sleep(300);
                return true;
            }

            return false;
        }

        private static AutomationElement FindSelectedTabPane(AutomationElement tabControl, string expectedName)
        {
            AutomationElement firstPane = GetChildByTypeAndIndex(tabControl, ControlType.Pane, 0);
            if (firstPane != null)
            {
                string paneName = SafeCurrentString(firstPane, AutomationElement.NameProperty);
                if (string.Equals(paneName, expectedName, StringComparison.Ordinal))
                {
                    return firstPane;
                }
            }

            return tabControl.FindFirst(
                TreeScope.Descendants,
                new AndCondition(
                    new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Pane),
                    new PropertyCondition(AutomationElement.NameProperty, expectedName)));
        }

        private static AutomationElement FindInternetExplorerHost(AutomationElement scope)
        {
            AutomationElement ieServer = scope.FindFirst(
                TreeScope.Descendants,
                new PropertyCondition(AutomationElement.ClassNameProperty, "Internet Explorer_Server"));

            if (ieServer != null)
            {
                return ieServer;
            }

            return scope.FindFirst(TreeScope.Descendants, new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Document));
        }

        private static bool TryCallLogout(AutomationElement ieHost)
        {
            int hwnd = SafeCurrentInt(ieHost, AutomationElement.NativeWindowHandleProperty);
            if (hwnd == 0)
            {
                return false;
            }

            IntPtr lResult;
            uint htmlGetObjectMsg = RegisterWindowMessage("WM_HTML_GETOBJECT");
            IntPtr sendResult = SendMessageTimeout(
                new IntPtr(hwnd),
                htmlGetObjectMsg,
                IntPtr.Zero,
                IntPtr.Zero,
                SendMessageTimeoutFlags.SMTO_ABORTIFHUNG,
                2000,
                out lResult);

            if (sendResult == IntPtr.Zero || lResult == IntPtr.Zero)
            {
                return false;
            }

            Guid iidDispatch = new Guid("00020400-0000-0000-C000-000000000046");
            object documentObject;
            int hr = ObjectFromLresult(lResult, ref iidDispatch, IntPtr.Zero, out documentObject);
            if (hr != 0 || documentObject == null)
            {
                return false;
            }

            try
            {
                object parentWindow = documentObject.GetType().InvokeMember(
                    "parentWindow",
                    BindingFlags.GetProperty,
                    null,
                    documentObject,
                    null);

                if (parentWindow == null)
                {
                    return false;
                }

                // logout が未定義でも例外を飲み込む JavaScript を実行する
                object[] args = { "try{if(typeof logout === 'function'){logout();}}catch(e){}", "JavaScript" };
                parentWindow.GetType().InvokeMember(
                    "execScript",
                    BindingFlags.InvokeMethod,
                    null,
                    parentWindow,
                    args);

                return true;
            }
            catch
            {
                return false;
            }
            finally
            {
                ReleaseComObjectQuietly(documentObject);
            }
        }

        private static AutomationElement GetChildByTypeAndIndex(AutomationElement parent, ControlType controlType, int index)
        {
            if (parent == null)
            {
                return null;
            }

            AutomationElementCollection children = parent.FindAll(
                TreeScope.Children,
                new PropertyCondition(AutomationElement.ControlTypeProperty, controlType));

            if (children == null || index < 0 || children.Count <= index)
            {
                return null;
            }

            return children[index];
        }

        private static string SafeCurrentString(AutomationElement element, AutomationProperty property)
        {
            if (element == null)
            {
                return null;
            }

            try
            {
                object value = element.GetCurrentPropertyValue(property, true);
                return value as string;
            }
            catch
            {
                return null;
            }
        }

        private static int SafeCurrentInt(AutomationElement element, AutomationProperty property)
        {
            if (element == null)
            {
                return 0;
            }

            try
            {
                object value = element.GetCurrentPropertyValue(property, true);
                if (value is int)
                {
                    return (int)value;
                }
            }
            catch
            {
                // no-op
            }

            return 0;
        }

        private static void ReleaseComObjectQuietly(object comObject)
        {
            if (comObject == null)
            {
                return;
            }

            try
            {
                if (Marshal.IsComObject(comObject))
                {
                    Marshal.FinalReleaseComObject(comObject);
                }
            }
            catch
            {
                // no-op
            }
        }

        [Flags]
        private enum SendMessageTimeoutFlags : uint
        {
            SMTO_NORMAL = 0x0,
            SMTO_BLOCK = 0x1,
            SMTO_ABORTIFHUNG = 0x2,
            SMTO_NOTIMEOUTIFNOTHUNG = 0x8
        }

        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        private static extern uint RegisterWindowMessage(string lpString);

        [DllImport("user32.dll", SetLastError = true)]
        private static extern IntPtr SendMessageTimeout(
            IntPtr hWnd,
            uint Msg,
            IntPtr wParam,
            IntPtr lParam,
            SendMessageTimeoutFlags fuFlags,
            uint uTimeout,
            out IntPtr lpdwResult);

        [DllImport("oleacc.dll")]
        private static extern int ObjectFromLresult(
            IntPtr lResult,
            ref Guid riid,
            IntPtr wParam,
            [MarshalAs(UnmanagedType.Interface)] out object ppvObject);
    }
}