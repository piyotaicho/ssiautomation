#
# 電子カルテシステムへのログイン状態を確認する 
#
Add-Type -AssemblyName UIAutomationClient
Add-Type -AssemblyName UIAutomationTypes

$sourceCheckLogin = @'
using System;
using System.Diagnostics;
using System.Windows.Automation;

public class SSICheckLogin
{
    private const string ProcessName = "LoginOutWindow";
    private const string LogoutWindowName = "ログアウト";
    private const string UserNameAutomationId = "lblUserName";

    private static AutomationElement RootElement
    {
        get
        {
            return AutomationElement.RootElement;
        }
    }

    private static AutomationElement GetMainWindowByProcessID(int processId)
    {
        if (processId <= 0)
        {
            return null;
        }

        Condition cond = new AndCondition(
            new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Window),
            new PropertyCondition(AutomationElement.ProcessIdProperty, processId)
        );
        return RootElement.FindFirst(TreeScope.Element | TreeScope.Children, cond);
    }

    private static AutomationElement GetChildWindowByName(AutomationElement parent, string name)
    {
        if (parent == null || string.IsNullOrEmpty(name))
        {
            return null;
        }

        Condition cond = new AndCondition(
            new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Window),
            new PropertyCondition(AutomationElement.NameProperty, name)
        );
        return parent.FindFirst(TreeScope.Children, cond);
    }

    private static AutomationElement GetDescendantTextByAutomationId(AutomationElement parent, string automationId)
    {
        if (parent == null || string.IsNullOrEmpty(automationId))
        {
            return null;
        }

        Condition cond = new AndCondition(
            new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Text),
            new PropertyCondition(AutomationElement.AutomationIdProperty, automationId)
        );
        return parent.FindFirst(TreeScope.Descendants, cond);
    }

    private static AutomationElement GetLoginOutMainWindow()
    {
        Process[] processes = Process.GetProcessesByName(ProcessName);
        foreach (Process process in processes)
        {
            AutomationElement appWindow = GetMainWindowByProcessID(process.Id);
            if (appWindow != null)
            {
                return appWindow;
            }
        }
        return null;
    }

    public static bool check()
    {
        AutomationElement appWindow = GetLoginOutMainWindow();
        if (appWindow == null)
        {
            return false;
        }

        AutomationElement logoutWindow = GetChildWindowByName(appWindow, LogoutWindowName);
        return logoutWindow != null;
    }

    public static string fetchName()
    {
        AutomationElement appWindow = GetLoginOutMainWindow();
        if (appWindow == null)
        {
            throw new InvalidOperationException("電子カルテシステムが起動していません");
        }

        AutomationElement logoutWindow = GetChildWindowByName(appWindow, LogoutWindowName);
        if (logoutWindow == null)
        {
            return null;
        }

        AutomationElement userNameLabel = GetDescendantTextByAutomationId(logoutWindow, UserNameAutomationId);
        if (userNameLabel == null)
        {
            return null;
        }

        return userNameLabel.Current.Name;
    }
}
'@

if (-not ([System.Management.Automation.PSTypeName]'SSICheckLogin').Type) {
    Add-Type -TypeDefinition $sourceCheckLogin -Language CSharp -ReferencedAssemblies @(
        [Windows.Automation.AutomationElement].Assembly.Location,
        [Windows.Automation.TreeScope].Assembly.Location,
        [System.Diagnostics.Process].Assembly.Location,
        [System.ComponentModel.Component].Assembly.Location
    )
}

# . でライブラリとして利用していない場合のデフォルトアクション
function Invoke-CheckLogin {
    return [SSICheckLogin]::check()
}

if ($MyInvocation.InvocationName -ne '.') {
    if (Invoke-CheckLogin) {
        'ログインしています' | Out-Default
    } else {
        'ログインしていません' | Out-Default
    }
}
