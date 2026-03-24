using System.Drawing;
using System.Drawing.Imaging;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace WindowSnapper;

/// <summary>
/// Captures a window in one of four modes and saves the result.
/// Runs on a dedicated STA background thread so it can show WinForms dialogs
/// (SaveFileDialog) without needing to marshal to the UI thread.
/// </summary>
internal sealed class ScreenCapture
{
    private readonly string _defaultDir;

    public ScreenCapture()
    {
        _defaultDir = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.MyPictures),
            "WindowSnapper");
        Directory.CreateDirectory(_defaultDir);
    }

    // ── Public entry point ────────────────────────────────────────────────────

    public void CaptureAndSave(IntPtr hwnd, string mode)
    {
        Bitmap? img = null;
        try
        {
            img = mode switch
            {
                "single"     => CaptureSingle(hwnd),
                "vertical"   => CaptureVertical(hwnd),
                "horizontal" => CaptureHorizontal(hwnd),
                "all"        => CaptureAll(hwnd),
                _            => CaptureSingle(hwnd),
            };

            SaveWithDialog(img);
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Capture failed:\n{ex.Message}", "Window Snapper",
                            MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
        finally
        {
            img?.Dispose();
        }
    }

    // ── Capture modes ─────────────────────────────────────────────────────────

    private Bitmap CaptureSingle(IntPtr hwnd)
    {
        BringToFront(hwnd);
        var cr = FindContentRect(hwnd);
        ScrollToTop(hwnd, cr);
        ScrollToLeft(hwnd, cr);
        Thread.Sleep(150);
        return GrabRect(cr);
    }

    private Bitmap CaptureVertical(IntPtr hwnd)
    {
        BringToFront(hwnd);
        var cr = FindContentRect(hwnd);
        ScrollToTop(hwnd, cr);

        var frames = new List<Bitmap>();
        Bitmap? prev = null;

        for (int i = 0; i < 60; i++)
        {
            var frame = GrabRect(cr);
            bool resized = prev != null && frame.Size != prev.Size;
            if (resized) { frame.Dispose(); break; }

            bool done = prev != null && AreSimilar(frame, prev);
            prev?.Dispose();
            prev = (Bitmap)frame.Clone();
            frames.Add(frame);   // always keep — stitcher will skip if truly identical
            if (done) break;
            ScrollPageDown(hwnd, cr);
        }
        prev?.Dispose();

        return StitchVertical(frames);
    }

    private Bitmap CaptureHorizontal(IntPtr hwnd)
    {
        BringToFront(hwnd);
        var cr = FindContentRect(hwnd);
        ScrollToLeft(hwnd, cr);

        var frames = new List<Bitmap>();
        Bitmap? prev = null;

        for (int i = 0; i < 20; i++)
        {
            var frame = GrabRect(cr);
            bool resized = prev != null && frame.Size != prev.Size;
            if (resized) { frame.Dispose(); break; }

            bool done = prev != null && AreSimilar(frame, prev);
            prev?.Dispose();
            prev = (Bitmap)frame.Clone();
            frames.Add(frame);
            if (done) break;
            ScrollPageRight(hwnd, cr);
        }
        prev?.Dispose();

        return StitchHorizontal(frames);
    }

    private Bitmap CaptureAll(IntPtr hwnd)
    {
        BringToFront(hwnd);
        var cr = FindContentRect(hwnd);
        ScrollToTop(hwnd, cr);
        ScrollToLeft(hwnd, cr);

        var columns      = new List<Bitmap>();
        Bitmap? prevColFirst = null;

        for (int col = 0; col < 20; col++)
        {
            ScrollToTop(hwnd, cr);

            var frames = new List<Bitmap>();
            Bitmap? prev = null;

            for (int row = 0; row < 60; row++)
            {
                var frame = GrabRect(cr);
                bool resized = prev != null && frame.Size != prev.Size;
                if (resized) { frame.Dispose(); break; }

                bool done = prev != null && AreSimilar(frame, prev);
                prev?.Dispose();
                prev = (Bitmap)frame.Clone();
                frames.Add(frame);
                if (done) break;
                ScrollPageDown(hwnd, cr);
            }
            prev?.Dispose();

            if (frames.Count == 0) break;

            var firstFrame = frames[0];
            if (prevColFirst != null && AreSimilar(firstFrame, prevColFirst))
            {
                foreach (var f in frames) f.Dispose();
                break;
            }
            prevColFirst?.Dispose();
            prevColFirst = (Bitmap)firstFrame.Clone();

            columns.Add(StitchVertical(frames));
            ScrollPageRight(hwnd, cr);
        }
        prevColFirst?.Dispose();

        return StitchHorizontal(columns);
    }

    // ── Content area detection ────────────────────────────────────────────────

    /// <summary>
    /// Finds the rectangle (in screen coordinates) of the scrollable content
    /// child window, falling back to the window's client area if no suitable
    /// child is found.
    /// </summary>
    private static Rectangle FindContentRect(IntPtr hwnd)
    {
        var winRect = GetWindowRect(hwnd);

        // Probe at 65% width, 50% height — right of any nav pane, away from scrollbar.
        var probe = new NativeMethods.POINT
        {
            X = winRect.X + winRect.Width  * 65 / 100,
            Y = winRect.Y + winRect.Height / 2,
        };

        var child = NativeMethods.WindowFromPoint(probe);
        if (child != IntPtr.Zero && child != hwnd && IsDescendant(hwnd, child))
        {
            NativeMethods.GetClientRect(hwnd,  out var parentClient);
            NativeMethods.GetClientRect(child, out var childClient);

            int parentArea = parentClient.Width * parentClient.Height;
            int childArea  = childClient.Width  * childClient.Height;

            // Only use the child if it covers at least 1/3 of the parent's client
            // area — rules out tiny buttons/tabs that happen to be at the probe point.
            if (parentArea > 0 && childArea * 3 >= parentArea)
                return TrimScrollBars(child, ClientRectToScreen(child, childClient));
        }

        // Fallback: full client area of the top-level window (excludes title bar
        // and resize border but not toolbars/ribbons).
        NativeMethods.GetClientRect(hwnd, out var fallback);
        return TrimScrollBars(hwnd, ClientRectToScreen(hwnd, fallback));
    }

    /// <summary>Walks the parent chain to check if <paramref name="child"/> is
    /// a descendant of <paramref name="ancestor"/>.</summary>
    private static bool IsDescendant(IntPtr ancestor, IntPtr child)
    {
        var cur = NativeMethods.GetParent(child);
        while (cur != IntPtr.Zero)
        {
            if (cur == ancestor) return true;
            cur = NativeMethods.GetParent(cur);
        }
        return false;
    }

    /// <summary>Converts a client-coordinate RECT to screen-coordinate Rectangle.</summary>
    private static Rectangle ClientRectToScreen(IntPtr hwnd, NativeMethods.RECT clientRect)
    {
        var origin = new NativeMethods.POINT { X = 0, Y = 0 };
        NativeMethods.ClientToScreen(hwnd, ref origin);
        return new Rectangle(origin.X, origin.Y, clientRect.Width, clientRect.Height);
    }

    /// <summary>
    /// Returns <paramref name="cr"/> with scrollbar-occupied edges removed.
    /// First tries WindowFromPoint edge probing (finds separate scrollbar child windows).
    /// Falls back to luminance-variance analysis: a scrollbar track is near-solid grey
    /// (very low variance) while actual content is varied (high variance).
    /// </summary>
    private static Rectangle TrimScrollBars(IntPtr contentChild, Rectangle cr)
    {
        int sbW = NativeMethods.GetSystemMetrics(NativeMethods.SM_CXVSCROLL);
        int sbH = NativeMethods.GetSystemMetrics(NativeMethods.SM_CYHSCROLL);

        // ── Vertical scrollbar (right edge) ───────────────────────────────────
        if (cr.Width > sbW * 3 && cr.Height > sbH * 2)
        {
            var probe  = new NativeMethods.POINT { X = cr.Right - sbW / 2, Y = cr.Y + cr.Height / 2 };
            var rChild = NativeMethods.WindowFromPoint(probe);
            bool hasV  = (rChild != IntPtr.Zero && rChild != contentChild)
                      || EdgeIsScrollBar(cr, vertical: true,  sbW, sbH);
            if (hasV) cr = new Rectangle(cr.X, cr.Y, cr.Width - sbW, cr.Height);
        }

        // ── Horizontal scrollbar (bottom edge) ────────────────────────────────
        if (cr.Width > sbW * 2 && cr.Height > sbH * 3)
        {
            var probe  = new NativeMethods.POINT { X = cr.X + cr.Width / 2, Y = cr.Bottom - sbH / 2 };
            var bChild = NativeMethods.WindowFromPoint(probe);
            bool hasH  = (bChild != IntPtr.Zero && bChild != contentChild)
                      || EdgeIsScrollBar(cr, vertical: false, sbW, sbH);
            if (hasH) cr = new Rectangle(cr.X, cr.Y, cr.Width, cr.Height - sbH);
        }

        return cr;
    }

    /// <summary>
    /// Captures a strip at the edge and a reference strip just inward, then compares
    /// their luminance variance.  Scrollbar track ≈ solid grey → near-zero variance.
    /// Content (file names, icons, dates) → high variance.
    /// </summary>
    private static bool EdgeIsScrollBar(Rectangle cr, bool vertical, int sbW, int sbH)
    {
        Rectangle edge, inside;
        if (vertical)
        {
            int h  = Math.Min(cr.Height / 2, 200);
            int y  = cr.Y + (cr.Height - h) / 2;
            edge   = new Rectangle(cr.Right - sbW,         y, sbW, h);
            inside = new Rectangle(cr.Right - sbW * 3 - 2, y, sbW, h);
        }
        else
        {
            int w  = Math.Min(cr.Width / 2, 400);
            int x  = cr.X + (cr.Width - w) / 2;
            edge   = new Rectangle(x, cr.Bottom - sbH,         w, sbH);
            inside = new Rectangle(x, cr.Bottom - sbH * 3 - 2, w, sbH);
        }

        float edgeVar   = LuminanceVariance(edge);
        float insideVar = LuminanceVariance(inside);

        // Edge must be very uniform (< 200) AND at least 4× less varied than the content
        // strip inward.  The second guard avoids trimming blank / uniform-background windows.
        return edgeVar < 200f && insideVar > 200f && edgeVar < insideVar * 0.25f;
    }

    private static float LuminanceVariance(Rectangle sr)
    {
        if (sr.Width <= 0 || sr.Height <= 0) return 0f;

        using var bmp = new Bitmap(sr.Width, sr.Height, PixelFormat.Format32bppArgb);
        using var g   = Graphics.FromImage(bmp);
        var destDC = g.GetHdc();
        var srcDC  = NativeMethods.GetDC(IntPtr.Zero);
        try
        {
            NativeMethods.BitBlt(destDC, 0, 0, sr.Width, sr.Height,
                                 srcDC, sr.X, sr.Y, NativeMethods.SRCCOPY);
        }
        finally { g.ReleaseHdc(destDC); NativeMethods.ReleaseDC(IntPtr.Zero, srcDC); }

        const int step = 2;
        var data  = bmp.LockBits(new Rectangle(0, 0, sr.Width, sr.Height),
                                 ImageLockMode.ReadOnly, PixelFormat.Format32bppArgb);
        var bytes = new byte[data.Stride * sr.Height];
        Marshal.Copy(data.Scan0, bytes, 0, bytes.Length);
        bmp.UnlockBits(data);

        float sum = 0, sum2 = 0;
        int   n   = 0;
        for (int y = 0; y < sr.Height; y += step)
        {
            int rowOff = y * data.Stride;
            for (int x = 0; x < sr.Width; x += step)
            {
                int off = rowOff + x * 4;
                if (off + 2 >= bytes.Length) continue;
                float lum = 0.114f * bytes[off] + 0.587f * bytes[off + 1] + 0.299f * bytes[off + 2];
                sum  += lum;
                sum2 += lum * lum;
                n++;
            }
        }
        if (n == 0) return 0f;
        float mean = sum / n;
        return sum2 / n - mean * mean;   // Var(X) = E[X²] − E[X]²
    }

    // ── Window helpers ────────────────────────────────────────────────────────

    private static Rectangle GetWindowRect(IntPtr hwnd)
    {
        int hr = NativeMethods.DwmGetWindowAttribute(
            hwnd,
            NativeMethods.DWMWA_EXTENDED_FRAME_BOUNDS,
            out var r,
            Marshal.SizeOf<NativeMethods.RECT>());

        if (hr != 0) NativeMethods.GetWindowRect(hwnd, out r);
        return r.ToRectangle();
    }

    private static void BringToFront(IntPtr hwnd)
    {
        // Only un-minimise — never un-maximise (that would resize the window and
        // break both the captured rect and the similarity comparison).
        var wp = new NativeMethods.WINDOWPLACEMENT
            { length = Marshal.SizeOf<NativeMethods.WINDOWPLACEMENT>() };
        NativeMethods.GetWindowPlacement(hwnd, ref wp);
        if (wp.showCmd == NativeMethods.SW_SHOWMINIMIZED)
            NativeMethods.ShowWindow(hwnd, NativeMethods.SW_RESTORE);

        NativeMethods.SetForegroundWindow(hwnd);
        Thread.Sleep(250);
    }

    // ── Scroll helpers ────────────────────────────────────────────────────────

    private static void ScrollToTop(IntPtr hwnd, Rectangle cr)
    {
        NativeMethods.PostMessage(hwnd, NativeMethods.WM_VSCROLL,
            new IntPtr(NativeMethods.SB_TOP), IntPtr.Zero);
        Thread.Sleep(40);
        PlaceCursorInContent(cr);
        NativeMethods.mouse_event(NativeMethods.MOUSEEVENTF_WHEEL, 0, 0, 120 * 200, 0);
        Thread.Sleep(350);
    }

    private static void ScrollToLeft(IntPtr hwnd, Rectangle cr)
    {
        NativeMethods.PostMessage(hwnd, NativeMethods.WM_HSCROLL,
            new IntPtr(NativeMethods.SB_LEFT), IntPtr.Zero);
        Thread.Sleep(40);
        PlaceCursorInContent(cr);
        NativeMethods.mouse_event(NativeMethods.MOUSEEVENTF_HWHEEL, 0, 0, 120 * 200, 0);
        Thread.Sleep(300);
    }

    private static void ScrollPageDown(IntPtr hwnd, Rectangle cr)
    {
        PlaceCursorInContent(cr);
        // Each -120 unit ≈ 3 lines ≈ ~50 px.  Scroll ~65% of the viewport so
        // there is always a generous overlap for the stitcher to work with.
        int ticks = Math.Max(5, cr.Height * 65 / 100 / 50);
        NativeMethods.mouse_event(NativeMethods.MOUSEEVENTF_WHEEL, 0, 0, -120 * ticks, 0);
        Thread.Sleep(300);
    }

    private static void ScrollPageRight(IntPtr hwnd, Rectangle cr)
    {
        NativeMethods.PostMessage(hwnd, NativeMethods.WM_HSCROLL,
            new IntPtr(NativeMethods.SB_PAGERIGHT), IntPtr.Zero);
        PlaceCursorInContent(cr);
        int ticks = Math.Max(5, cr.Width * 65 / 100 / 50);
        NativeMethods.mouse_event(NativeMethods.MOUSEEVENTF_HWHEEL, 0, 0, -120 * ticks, 0);
        Thread.Sleep(300);
    }

    /// <summary>
    /// Places the cursor in the centre of the content rectangle so that mouse-
    /// wheel scroll events reach the right control.
    /// </summary>
    private static void PlaceCursorInContent(Rectangle cr)
    {
        NativeMethods.SetCursorPos(cr.X + cr.Width / 2, cr.Y + cr.Height / 2);
        Thread.Sleep(30);
    }

    // ── Screenshot via BitBlt ─────────────────────────────────────────────────

    private static Bitmap GrabRect(Rectangle rect)
    {
        if (rect.Width <= 0 || rect.Height <= 0)
            throw new InvalidOperationException("Capture area has zero size.");

        // Park the cursor above the content area before snapping.
        // This clears any hover highlight (Explorer file rows, browser links, etc.)
        // so that two screenshots of identical content are pixel-identical — required
        // for both AreSimilar and FindVerticalOverlap.
        NativeMethods.SetCursorPos(rect.X + Math.Min(120, rect.Width / 2), rect.Y - 30);
        Thread.Sleep(60);

        var bmp    = new Bitmap(rect.Width, rect.Height, PixelFormat.Format32bppArgb);
        using var g = Graphics.FromImage(bmp);
        var destDC = g.GetHdc();
        var srcDC  = NativeMethods.GetDC(IntPtr.Zero);  // whole-screen DC
        try
        {
            NativeMethods.BitBlt(destDC, 0, 0, rect.Width, rect.Height,
                                 srcDC, rect.X, rect.Y, NativeMethods.SRCCOPY);
        }
        finally
        {
            g.ReleaseHdc(destDC);
            NativeMethods.ReleaseDC(IntPtr.Zero, srcDC);
        }
        return bmp;
    }

    // ── Image comparison ──────────────────────────────────────────────────────

    private static bool AreSimilar(Bitmap a, Bitmap b, float threshold = 0.99f)
    {
        if (a.Width != b.Width || a.Height != b.Height) return false;

        var dataA = a.LockBits(new Rectangle(0, 0, a.Width, a.Height),
                               ImageLockMode.ReadOnly, PixelFormat.Format32bppArgb);
        var dataB = b.LockBits(new Rectangle(0, 0, b.Width, b.Height),
                               ImageLockMode.ReadOnly, PixelFormat.Format32bppArgb);
        try
        {
            int len = dataA.Stride * a.Height;
            var ba  = new byte[len];
            var bb  = new byte[len];
            Marshal.Copy(dataA.Scan0, ba, 0, len);
            Marshal.Copy(dataB.Scan0, bb, 0, len);

            long diff = 0;
            for (int i = 0; i < len; i++) diff += Math.Abs(ba[i] - bb[i]);

            return (1f - (float)diff / (255f * len)) >= threshold;
        }
        finally
        {
            a.UnlockBits(dataA);
            b.UnlockBits(dataB);
        }
    }

    // ── Vertical overlap detection ────────────────────────────────────────────

    /// <summary>
    /// Finds how many pixels at the top of <paramref name="below"/> are already
    /// present near the bottom of <paramref name="above"/>.
    /// Uses a downsampled grayscale strip for speed.
    /// </summary>
    private static int FindVerticalOverlap(Bitmap above, Bitmap below, int stripH = 80)
    {
        int w  = Math.Min(above.Width,  below.Width);
        int hA = above.Height;
        int hB = below.Height;
        if (hB < stripH || hA < stripH || w == 0) return 0;

        const int step = 4;

        // The overlap must lie within the last 'hB' rows of 'above'.
        // Searching the whole accumulated image risks matching a similar-looking
        // row elsewhere, which overestimates the overlap and silently eats content.
        int searchFrom = Math.Max(hA / 2, hA - hB);
        int searchH    = hA - searchFrom;
        if (searchH < stripH) return 0;

        var dataA = above.LockBits(new Rectangle(0, searchFrom, w, searchH),
                                   ImageLockMode.ReadOnly, PixelFormat.Format32bppArgb);
        var dataB = below.LockBits(new Rectangle(0, 0, w, Math.Min(hB, stripH * 2)),
                                   ImageLockMode.ReadOnly, PixelFormat.Format32bppArgb);
        try
        {
            var bytesA = new byte[dataA.Stride * searchH];
            Marshal.Copy(dataA.Scan0, bytesA, 0, bytesA.Length);

            var bytesB = new byte[dataB.Stride * dataB.Height];
            Marshal.Copy(dataB.Scan0, bytesB, 0, bytesB.Length);

            // Template = first stripH rows of 'below' (downsampled)
            float[] tmpl = SampleGray(bytesB, dataB.Stride, 0, stripH, step, w);

            float bestDiff = float.MaxValue;
            int   bestLocalRow = searchH - 1;  // local to the locked sub-region

            for (int row = 0; row <= searchH - stripH; row += step)
            {
                float[] cand = SampleGray(bytesA, dataA.Stride, row, stripH, step, w);
                float diff   = MeanAbsDiff(tmpl, cand);
                if (diff < bestDiff) { bestDiff = diff; bestLocalRow = row; }
            }

            if (bestDiff > 8f) return 0;
            // Convert local row back to absolute position in 'above'
            int bestRow = searchFrom + bestLocalRow;
            return Math.Max(0, Math.Min(hA - bestRow, hB));
        }
        finally
        {
            above.UnlockBits(dataA);
            below.UnlockBits(dataB);
        }
    }

    private static float[] SampleGray(byte[] bytes, int stride, int startRow, int height, int step, int width)
    {
        var list = new List<float>();
        for (int y = 0; y < height; y += step)
        {
            int rowOff = (startRow + y) * stride;
            for (int x = 0; x < width; x += step)
            {
                int off = rowOff + x * 4;
                if (off + 2 >= bytes.Length) continue;
                // BGRA layout
                list.Add(0.114f * bytes[off] + 0.587f * bytes[off + 1] + 0.299f * bytes[off + 2]);
            }
        }
        return [.. list];
    }

    private static float MeanAbsDiff(float[] a, float[] b)
    {
        int len = Math.Min(a.Length, b.Length);
        if (len == 0) return float.MaxValue;
        float s = 0;
        for (int i = 0; i < len; i++) s += Math.Abs(a[i] - b[i]);
        return s / len;
    }

    // ── Stitching ─────────────────────────────────────────────────────────────

    private static Bitmap StitchVertical(List<Bitmap> frames)
    {
        if (frames.Count == 0) throw new InvalidOperationException("No frames captured.");
        if (frames.Count == 1) return frames[0];

        Bitmap result = frames[0];

        for (int i = 1; i < frames.Count; i++)
        {
            var next    = frames[i];
            int overlap = FindVerticalOverlap(result, next);
            int cropY   = Math.Max(0, overlap);
            if (cropY >= next.Height) continue;

            int newH   = result.Height + next.Height - cropY;
            int newW   = Math.Max(result.Width, next.Width);
            var canvas = new Bitmap(newW, newH, PixelFormat.Format32bppArgb);
            using var g = Graphics.FromImage(canvas);
            g.Clear(Color.White);
            g.DrawImage(result, 0, 0);
            g.DrawImage(next,   new Rectangle(0, result.Height, next.Width, next.Height - cropY),
                                new Rectangle(0, cropY,          next.Width, next.Height - cropY),
                                GraphicsUnit.Pixel);

            if (i > 1) result.Dispose();   // don't dispose frames[0] yet; caller might need it
            result = canvas;
        }
        return result;
    }

    private static Bitmap StitchHorizontal(List<Bitmap> cols)
    {
        if (cols.Count == 0) throw new InvalidOperationException("No columns captured.");
        if (cols.Count == 1) return cols[0];

        int totalW = cols.Sum(c => c.Width);
        int maxH   = cols.Max(c => c.Height);
        var canvas = new Bitmap(totalW, maxH, PixelFormat.Format32bppArgb);
        using var g = Graphics.FromImage(canvas);
        g.Clear(Color.White);

        int x = 0;
        foreach (var col in cols)
        {
            g.DrawImage(col, x, 0);
            x += col.Width;
        }
        foreach (var col in cols) col.Dispose();
        return canvas;
    }

    // ── Save ──────────────────────────────────────────────────────────────────

    private void SaveWithDialog(Bitmap? img)
    {
        if (img is null)
        {
            MessageBox.Show("Capture returned an empty image.", "Window Snapper",
                            MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }

        var timestamp   = DateTime.Now.ToString("yyyyMMdd_HHmmss");
        var defaultName = $"snapshot_{timestamp}.png";

        using var dlg = new SaveFileDialog
        {
            Title            = "Save Screenshot",
            InitialDirectory = _defaultDir,
            FileName         = defaultName,
            DefaultExt       = "png",
            Filter           = "PNG image|*.png|JPEG image|*.jpg;*.jpeg|All files|*.*",
        };

        if (dlg.ShowDialog() != DialogResult.OK) return;

        img.Save(dlg.FileName);

        MessageBox.Show($"Screenshot saved to:\n{dlg.FileName}",
                        "Window Snapper", MessageBoxButtons.OK, MessageBoxIcon.Information);
    }
}
