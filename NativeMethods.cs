using System.Runtime.InteropServices;

namespace WindowSnapper;

internal static class NativeMethods
{
    // ── Delegates ────────────────────────────────────────────────────────────
    public delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);
    public delegate IntPtr HookProc(int nCode, IntPtr wParam, IntPtr lParam);

    // ── Structs ───────────────────────────────────────────────────────────────
    [StructLayout(LayoutKind.Sequential)]
    public struct RECT
    {
        public int Left, Top, Right, Bottom;
        public System.Drawing.Rectangle ToRectangle() =>
            System.Drawing.Rectangle.FromLTRB(Left, Top, Right, Bottom);
        public int Width  => Right  - Left;
        public int Height => Bottom - Top;
    }

    [StructLayout(LayoutKind.Sequential)]
    public struct POINT { public int X, Y; }

    [StructLayout(LayoutKind.Sequential)]
    public struct MSLLHOOKSTRUCT
    {
        public POINT pt;
        public uint mouseData, flags, time;
        public IntPtr dwExtraInfo;
    }

    [StructLayout(LayoutKind.Sequential)]
    public struct KBDLLHOOKSTRUCT
    {
        public uint vkCode, scanCode, flags, time;
        public IntPtr dwExtraInfo;
    }

    // ── Constants ─────────────────────────────────────────────────────────────
    public const int WH_MOUSE_LL    = 14;
    public const int WH_KEYBOARD_LL = 13;
    public const int WM_MOUSEMOVE    = 0x0200;
    public const int WM_LBUTTONDOWN  = 0x0201;
    public const int WM_RBUTTONDOWN  = 0x0204;
    public const int WM_KEYDOWN      = 0x0100;
    public const int WM_VSCROLL      = 0x0115;
    public const int WM_HSCROLL      = 0x0114;

    public const int VK_ESCAPE  = 0x1B;
    public const int VK_NEXT    = 0x22;  // Page Down
    public const int VK_HOME    = 0x24;
    public const int VK_CONTROL = 0x11;

    public const uint KEYEVENTF_KEYUP = 0x0002;

    public const int SB_TOP       = 6;
    public const int SB_LEFT      = 6;
    public const int SB_PAGEDOWN  = 3;
    public const int SB_PAGERIGHT = 3;

    public const int SW_RESTORE = 9;
    public const int DWMWA_EXTENDED_FRAME_BOUNDS = 9;

    public const int SRCCOPY = 0x00CC0020;

    public const uint MOUSEEVENTF_HWHEEL = 0x01000;
    public const uint MOUSEEVENTF_WHEEL  = 0x0800;

    // ── Window enumeration / info ─────────────────────────────────────────────
    [DllImport("user32.dll")]
    public static extern bool EnumWindows(EnumWindowsProc lpEnumFunc, IntPtr lParam);

    [DllImport("user32.dll")]
    public static extern bool IsWindowVisible(IntPtr hWnd);

    [DllImport("user32.dll")]
    public static extern bool GetWindowRect(IntPtr hWnd, out RECT rect);

    [DllImport("dwmapi.dll")]
    public static extern int DwmGetWindowAttribute(
        IntPtr hwnd, int dwAttribute,
        out RECT pvAttribute, int cbAttribute);

    // ── Window control ────────────────────────────────────────────────────────
    [DllImport("user32.dll")]
    public static extern bool SetForegroundWindow(IntPtr hWnd);

    [DllImport("user32.dll")]
    public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

    [DllImport("user32.dll")]
    public static extern bool PostMessage(IntPtr hWnd, uint msg, IntPtr wParam, IntPtr lParam);

    // ── Input ─────────────────────────────────────────────────────────────────
    [DllImport("user32.dll")]
    public static extern bool SetCursorPos(int x, int y);

    [DllImport("user32.dll")]
    public static extern void keybd_event(byte bVk, byte bScan, uint dwFlags, UIntPtr dwExtraInfo);

    [DllImport("user32.dll")]
    public static extern void mouse_event(uint dwFlags, int dx, int dy, int dwData, int dwExtraInfo);

    // ── Hooks ─────────────────────────────────────────────────────────────────
    [DllImport("user32.dll", SetLastError = true)]
    public static extern IntPtr SetWindowsHookEx(int idHook, HookProc lpfn, IntPtr hMod, uint dwThreadId);

    [DllImport("user32.dll", SetLastError = true)]
    public static extern bool UnhookWindowsHookEx(IntPtr hhk);

    [DllImport("user32.dll")]
    public static extern IntPtr CallNextHookEx(IntPtr hhk, int nCode, IntPtr wParam, IntPtr lParam);

    [DllImport("kernel32.dll", CharSet = CharSet.Auto)]
    public static extern IntPtr GetModuleHandle(string? lpModuleName);

    // ── Screen capture ────────────────────────────────────────────────────────
    [DllImport("gdi32.dll")]
    public static extern bool BitBlt(
        IntPtr hdc, int nXDest, int nYDest, int nWidth, int nHeight,
        IntPtr hdcSrc, int nXSrc, int nYSrc, int dwRop);

    [DllImport("user32.dll")]
    public static extern IntPtr GetDC(IntPtr hWnd);

    [DllImport("user32.dll")]
    public static extern int ReleaseDC(IntPtr hWnd, IntPtr hDC);

    // ── DPI ───────────────────────────────────────────────────────────────────
    [DllImport("shcore.dll")]
    public static extern int SetProcessDpiAwareness(int value);

    // ── Window placement (used to detect minimised vs maximised) ──────────────
    [StructLayout(LayoutKind.Sequential)]
    public struct WINDOWPLACEMENT
    {
        public int length, flags, showCmd;
        public POINT ptMinPosition, ptMaxPosition;
        public RECT  rcNormalPosition, rcDevice;
    }

    public const int SW_SHOWMINIMIZED = 2;

    public const int GWL_STYLE    = -16;
    public const int WS_HSCROLL   = 0x00100000;
    public const int WS_VSCROLL   = 0x00200000;
    public const int SM_CXVSCROLL = 2;
    public const int SM_CYHSCROLL = 3;

    public const int OBJID_HSCROLL            = unchecked((int)0xFFFFFFFA);
    public const int OBJID_VSCROLL            = unchecked((int)0xFFFFFFFB);
    public const int STATE_SYSTEM_INVISIBLE   = 0x00008000;
    public const int STATE_SYSTEM_UNAVAILABLE = 0x00000001;

    [StructLayout(LayoutKind.Sequential)]
    public struct SCROLLBARINFO
    {
        public int  cbSize;
        public RECT rcScrollBar;
        public int  dxyLineButton, xyThumbTop, xyThumbBottom, reserved;
        [MarshalAs(UnmanagedType.ByValArray, SizeConst = 6)]
        public int[] rgstate;
    }

    [DllImport("user32.dll")]
    public static extern int GetWindowLong(IntPtr hWnd, int nIndex);

    [DllImport("user32.dll")]
    public static extern int GetSystemMetrics(int nIndex);

    [DllImport("user32.dll")]
    public static extern bool GetScrollBarInfo(IntPtr hwnd, int idObject, ref SCROLLBARINFO psbi);

    [DllImport("user32.dll")]
    public static extern bool GetWindowPlacement(IntPtr hWnd, ref WINDOWPLACEMENT lpwndpl);

    [DllImport("user32.dll")]
    public static extern IntPtr WindowFromPoint(POINT pt);

    [DllImport("user32.dll")]
    public static extern bool GetClientRect(IntPtr hWnd, out RECT lpRect);

    [DllImport("user32.dll")]
    public static extern bool ClientToScreen(IntPtr hWnd, ref POINT lpPoint);

    [DllImport("user32.dll")]
    public static extern IntPtr GetParent(IntPtr hWnd);
}
