using System.Drawing;
using System.Drawing.Drawing2D;
using System.Windows.Forms;

namespace WindowSnapper;

/// <summary>
/// Manages the system tray icon and top-level menu.
/// Extends ApplicationContext so we have no visible main window.
/// </summary>
internal sealed class SysTrayApp : ApplicationContext
{
    private readonly NotifyIcon   _notify;
    private readonly ScreenCapture _capture = new();
    private SelectorOverlay?      _overlay;

    public SysTrayApp()
    {
        _notify = new NotifyIcon
        {
            Icon             = BuildIcon(),
            Text             = "Window Snapper",
            Visible          = true,
            ContextMenuStrip = BuildMenu(),
        };

        // Left-click also opens the menu (friendlier UX)
        _notify.MouseClick += (_, e) =>
        {
            if (e.Button == MouseButtons.Left)
                _notify.ContextMenuStrip!.Show(Cursor.Position);
        };
    }

    // ── Menu ──────────────────────────────────────────────────────────────────

    private ContextMenuStrip BuildMenu()
    {
        var menu = new ContextMenuStrip();
        menu.Items.Add("Single Screen Snap",     null, (_, _) => StartCapture("single"));
        menu.Items.Add("Vertical Scroll Snap",   null, (_, _) => StartCapture("vertical"));
        menu.Items.Add("Horizontal Scroll Snap", null, (_, _) => StartCapture("horizontal"));
        menu.Items.Add("All Scrolls Snap",       null, (_, _) => StartCapture("all"));
        menu.Items.Add(new ToolStripSeparator());
        menu.Items.Add("Exit", null, (_, _) => Exit());
        return menu;
    }

    // ── Capture flow ──────────────────────────────────────────────────────────

    private void StartCapture(string mode)
    {
        if (_overlay != null) return;   // already in selection mode

        _overlay = new SelectorOverlay(mode);

        _overlay.WindowSelected += hwnd =>
        {
            _overlay = null;
            // Run capture on an STA thread so SaveFileDialog works without invoking.
            var t = new Thread(() => _capture.CaptureAndSave(hwnd, mode));
            t.SetApartmentState(ApartmentState.STA);
            t.IsBackground = true;
            t.Start();
        };

        _overlay.SelectionCancelled += () => _overlay = null;

        _overlay.Show();
    }

    private void Exit()
    {
        _overlay?.CancelSelection();
        _notify.Visible = false;
        _notify.Dispose();
        Application.Exit();
    }

    // ── Icon creation ─────────────────────────────────────────────────────────

    private static Icon BuildIcon()
    {
        using var bmp = new Bitmap(32, 32);
        using var g   = Graphics.FromImage(bmp);
        g.SmoothingMode = SmoothingMode.AntiAlias;
        g.Clear(Color.Transparent);

        // Camera body
        using var bodyBrush  = new SolidBrush(Color.FromArgb(50, 120, 200));
        using var outlinePen = new Pen(Color.FromArgb(200, 225, 255), 1.5f);
        using var bodyPath   = RoundedRect(new RectangleF(2, 9, 28, 17), 3f);
        g.FillPath(bodyBrush, bodyPath);
        g.DrawPath(outlinePen, bodyPath);

        // Viewfinder bump
        g.FillRectangle(bodyBrush, 10, 4, 8, 7);
        g.DrawRectangle(outlinePen, 10, 4, 8, 7);

        // Lens ring
        using var lensBrush = new SolidBrush(Color.FromArgb(20, 20, 70));
        g.FillEllipse(lensBrush, 8, 11, 14, 13);
        g.DrawEllipse(outlinePen, 8, 11, 14, 13);

        // Lens inner
        using var innerBrush = new SolidBrush(Color.FromArgb(80, 130, 220));
        g.FillEllipse(innerBrush, 11, 14, 8, 7);

        // Flash
        using var flashBrush = new SolidBrush(Color.FromArgb(255, 215, 50));
        g.FillRectangle(flashBrush, 23, 5, 5, 5);

        var hIcon = bmp.GetHicon();
        return Icon.FromHandle(hIcon);
    }

    private static GraphicsPath RoundedRect(RectangleF b, float r)
    {
        var p = new GraphicsPath();
        float d = r * 2;
        p.AddArc(b.X, b.Y, d, d, 180, 90);
        p.AddArc(b.Right - d, b.Y, d, d, 270, 90);
        p.AddArc(b.Right - d, b.Bottom - d, d, d, 0, 90);
        p.AddArc(b.X, b.Bottom - d, d, d, 90, 90);
        p.CloseFigure();
        return p;
    }
}
