using System.Drawing;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace WindowSnapper;

/// <summary>
/// Manages window-selection mode.
///
/// No transparent overlay — that approach causes black-window rendering bugs
/// when WS_EX_TRANSPARENT, OptimizedDoubleBuffer and TransparencyKey are combined.
///
/// Instead:
///   HintForm      – small opaque bar at the top of the primary screen.
///   HighlightForm – Region-clipped opaque border frame; interior is absent
///                   (not transparent — literally not part of the window) so
///                   clicks pass straight through to the window below.
///   Two global LL hooks handle all mouse and keyboard input.
/// </summary>
internal sealed class SelectorOverlay
{
    // ── Mode colours / labels ─────────────────────────────────────────────────
    private static readonly Dictionary<string, Color> ModeColors = new()
    {
        ["single"]     = Color.FromArgb(0,   255, 136),
        ["vertical"]   = Color.FromArgb(0,   170, 255),
        ["horizontal"] = Color.FromArgb(255, 170,   0),
        ["all"]        = Color.FromArgb(255,  68, 255),
    };
    private static readonly Dictionary<string, string> ModeLabels = new()
    {
        ["single"]     = "Single Screen Snap",
        ["vertical"]   = "Vertical Scroll Snap",
        ["horizontal"] = "Horizontal Scroll Snap",
        ["all"]        = "All Scrolls Snap",
    };

    // ── Events ────────────────────────────────────────────────────────────────
    public event Action<IntPtr>? WindowSelected;
    public event Action?         SelectionCancelled;

    // ── State ─────────────────────────────────────────────────────────────────
    private bool   _active;
    private IntPtr _currentHwnd;

    // ── Visuals ───────────────────────────────────────────────────────────────
    private HintForm?      _hint;
    private HighlightForm? _highlight;

    // ── Hooks (keep delegate refs alive — GC must not collect them) ───────────
    private NativeMethods.HookProc? _mouseProc;
    private NativeMethods.HookProc? _kbdProc;
    private IntPtr _mouseHook;
    private IntPtr _kbdHook;

    // ── Constructor / Show / Cancel ───────────────────────────────────────────

    public SelectorOverlay(string mode)
    {
        var color = ModeColors.TryGetValue(mode, out var c) ? c : Color.LimeGreen;
        var label = ModeLabels.TryGetValue(mode, out var l) ? l : mode;

        _hint      = new HintForm(label, color);
        _highlight = new HighlightForm(color);
    }

    public void Show()
    {
        _active = true;
        _hint!.Show();
        InstallHooks();
    }

    public void CancelSelection()
    {
        if (!_active) return;
        _active = false;
        UninstallHooks();
        _highlight?.Close();
        _hint?.Close();
        SelectionCancelled?.Invoke();
    }

    // ── Internal selection complete ───────────────────────────────────────────

    private void CompleteSelection(IntPtr hwnd)
    {
        // Called on the UI thread via BeginInvoke from the hook callback.
        _active = false;
        UninstallHooks();
        _highlight?.Close();
        _hint?.Close();

        // Small pause so both forms fully disappear before BitBlt starts.
        var timer = new System.Windows.Forms.Timer { Interval = 350 };
        timer.Tick += (_, _) =>
        {
            timer.Stop();
            timer.Dispose();
            WindowSelected?.Invoke(hwnd);
        };
        timer.Start();
    }

    // ── Hooks ─────────────────────────────────────────────────────────────────

    private void InstallHooks()
    {
        var hMod = NativeMethods.GetModuleHandle(null);
        _mouseProc = MouseHook;
        _kbdProc   = KeyboardHook;
        _mouseHook = NativeMethods.SetWindowsHookEx(NativeMethods.WH_MOUSE_LL, _mouseProc, hMod, 0);
        _kbdHook   = NativeMethods.SetWindowsHookEx(NativeMethods.WH_KEYBOARD_LL, _kbdProc, hMod, 0);
    }

    private void UninstallHooks()
    {
        if (_mouseHook != IntPtr.Zero) { NativeMethods.UnhookWindowsHookEx(_mouseHook); _mouseHook = IntPtr.Zero; }
        if (_kbdHook   != IntPtr.Zero) { NativeMethods.UnhookWindowsHookEx(_kbdHook);   _kbdHook   = IntPtr.Zero; }
    }

    private IntPtr MouseHook(int nCode, IntPtr wParam, IntPtr lParam)
    {
        if (nCode >= 0 && _active)
        {
            var info = Marshal.PtrToStructure<NativeMethods.MSLLHOOKSTRUCT>(lParam);

            switch ((int)wParam)
            {
                case NativeMethods.WM_MOUSEMOVE:
                    OnMouseMove(info.pt.X, info.pt.Y);
                    break;

                case NativeMethods.WM_LBUTTONDOWN:
                    // Use the currently highlighted window — avoids accidentally
                    // picking up our own border/hint form handles.
                    if (_currentHwnd != IntPtr.Zero)
                    {
                        var hwnd = _currentHwnd;
                        // BeginInvoke returns immediately, hook callback stays fast.
                        _hint!.BeginInvoke(() => CompleteSelection(hwnd));
                        return (IntPtr)1;   // swallow — don't let the click reach the window
                    }
                    break;

                case NativeMethods.WM_RBUTTONDOWN:
                    _hint!.BeginInvoke(CancelSelection);
                    return (IntPtr)1;
            }
        }
        return NativeMethods.CallNextHookEx(_mouseHook, nCode, wParam, lParam);
    }

    private IntPtr KeyboardHook(int nCode, IntPtr wParam, IntPtr lParam)
    {
        if (nCode >= 0 && _active && (int)wParam == NativeMethods.WM_KEYDOWN)
        {
            var info = Marshal.PtrToStructure<NativeMethods.KBDLLHOOKSTRUCT>(lParam);
            if ((int)info.vkCode == NativeMethods.VK_ESCAPE)
            {
                _hint!.BeginInvoke(CancelSelection);
                return (IntPtr)1;
            }
        }
        return NativeMethods.CallNextHookEx(_kbdHook, nCode, wParam, lParam);
    }

    private void OnMouseMove(int x, int y)
    {
        var hwnd = FindWindowAtPoint(x, y);
        if (hwnd == _currentHwnd) return;
        _currentHwnd = hwnd;

        if (hwnd == IntPtr.Zero) { _highlight?.Hide(); return; }

        var rect = GetWindowBounds(hwnd);
        _highlight!.SetWindow(rect);
    }

    // ── Helpers ───────────────────────────────────────────────────────────────

    private IntPtr FindWindowAtPoint(int x, int y)
    {
        // Exclude our own UI forms so they can never become the "selected" window.
        var hintHwnd      = _hint?.IsHandleCreated      == true ? _hint.Handle      : IntPtr.Zero;
        var highlightHwnd = _highlight?.IsHandleCreated == true ? _highlight.Handle : IntPtr.Zero;

        IntPtr result = IntPtr.Zero;
        NativeMethods.EnumWindows((hwnd, _) =>
        {
            if (hwnd == hintHwnd || hwnd == highlightHwnd) return true;
            if (!NativeMethods.IsWindowVisible(hwnd))      return true;
            if (!NativeMethods.GetWindowRect(hwnd, out var r)) return true;
            if (r.Width < 20 || r.Height < 20)             return true;
            if (x >= r.Left && x < r.Right && y >= r.Top && y < r.Bottom)
            {
                result = hwnd;
                return false;   // EnumWindows is Z-order top-first; first hit wins
            }
            return true;
        }, IntPtr.Zero);
        return result;
    }

    private static Rectangle GetWindowBounds(IntPtr hwnd)
    {
        int hr = NativeMethods.DwmGetWindowAttribute(
            hwnd, NativeMethods.DWMWA_EXTENDED_FRAME_BOUNDS,
            out var r, Marshal.SizeOf<NativeMethods.RECT>());
        if (hr != 0) NativeMethods.GetWindowRect(hwnd, out r);
        return r.ToRectangle();
    }

    // ── Inner form: hint bar ──────────────────────────────────────────────────

    private sealed class HintForm : Form
    {
        public HintForm(string modeLabel, Color modeColor)
        {
            FormBorderStyle = FormBorderStyle.None;
            TopMost         = true;
            ShowInTaskbar   = false;
            BackColor       = Color.FromArgb(30, 30, 30);
            Cursor          = Cursors.Cross;

            var screen = Screen.PrimaryScreen?.Bounds ?? Screen.AllScreens[0].Bounds;
            Bounds = new Rectangle(screen.X, screen.Y, screen.Width, 40);

            Controls.Add(new Label
            {
                Text      = $"  {modeLabel}  |  Click a window to capture  |  ESC / Right-click to cancel  ",
                ForeColor = modeColor,
                BackColor = Color.FromArgb(30, 30, 30),
                Font      = new Font("Segoe UI", 11, FontStyle.Regular),
                Dock      = DockStyle.Fill,
                TextAlign = ContentAlignment.MiddleCenter,
                Cursor    = Cursors.Cross,
            });
        }
    }

    // ── Inner form: highlight border ──────────────────────────────────────────

    private sealed class HighlightForm : Form
    {
        public HighlightForm(Color color)
        {
            FormBorderStyle = FormBorderStyle.None;
            TopMost         = true;
            ShowInTaskbar   = false;
            BackColor       = color;
            Cursor          = Cursors.Cross;
        }

        public void SetWindow(Rectangle winRect)
        {
            const int bw = 3;

            // Position the form slightly OUTSIDE the target window.
            Bounds = new Rectangle(
                winRect.X - bw, winRect.Y - bw,
                winRect.Width + bw * 2, winRect.Height + bw * 2);

            // Clip to just the frame; the interior region is absent — clicks fall
            // through to whatever is underneath (the target window).
            var frame = new Region(new Rectangle(0, 0, Width, Height));
            frame.Exclude(new Rectangle(bw, bw, Width - bw * 2, Height - bw * 2));
            Region = frame;

            if (!Visible) Show();
        }
    }
}
