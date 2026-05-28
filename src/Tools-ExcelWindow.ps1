# ExcelのCOMオブジェクトをhWndに基づいて取得する
$excelWindowFactory = @"
using System;
using System.Runtime.InteropServices;

public class ExcelWindowFactory {
    [DllImport("user32.dll", SetLastError = true)]
    private static extern IntPtr FindWindowEx(IntPtr hwndParent, IntPtr hwndChildAfter, string lpszClass, string lpszWindow);

    [DllImport("oleacc.dll", PreserveSig = false)]
    private static extern void AccessibleObjectFromWindow(IntPtr hwnd, uint dwId, ref Guid riid, [MarshalAs(UnmanagedType.IDispatch)] out object ppvObject);

    [DllImport("user32.dll")]
    public static extern IntPtr GetForgroundWindow();

    private static readonly Guid IID_IDispatch = new Guid("{00020400-0000-0000-C000-000000000046}");
    private const uint OBJID_NATIVEOM = 0xFFFFFFF0;

    public static object GetExcelApplicationFromHwnd(IntPtr hwnd) {
        if (hwnd == IntPtr.Zero) return null;
        // search the window handle which has 'XLDESK' -> 'EXCEL7' tree
        IntPtr hwndDesk = FindWindowEx(hwnd, IntPtr.Zero, "XLDESK", null);
        IntPtr hwndSheet = FindWindowEx(hwndDesk, IntPtr.Zero, "EXCEL7", null);

        if (hwndSheet  == IntPtr.Zero) return null;

        // Get COM object from Window
        Guid guid = IID_IDispatch;
        object dispatchObject;
        AccessibleObjectFromWindow(hwndSheet, OBJID_NATIVEOM, ref guid, out dispatchObject);

        if (dispatchObject == null) return null;

        try {
            dynamic window = dispatchObject;
            return window.Application;
        }
        catch {
            try {
                dynamic direct = dispatchObject;
                return direct.Application;
            }
            catch {
                return null;
            }
        }
    }
}
"@

if (-not ([System.Management.Automation.PSTypeName]'ExcelWindowFactory').Type) {
    Add-Type -TypeDefinition $excelWindowFactory -Language CSharp -ReferencedAssemblies @('System.Core', 'Microsoft.CSharp')
}
