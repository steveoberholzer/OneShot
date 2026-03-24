# OneShot — Window Snapper

A Windows system-tray utility that captures full-page screenshots of any window, including content that extends beyond the visible screen area. Select a window with a crosshair cursor, and OneShot automatically scrolls and stitches the result into a single seamless image.

## Features

- **Single Screen Snap** — captures the visible content area at its current scroll position (after resetting to top-left).
- **Vertical Scroll Snap** — scrolls the window from top to bottom, stitching every frame into one tall image.
- **Horizontal Scroll Snap** — scrolls left to right, stitching frames side by side.
- **All Scrolls Snap** — covers the full two-dimensional content area (vertical + horizontal).
- Captures only the **scrollable content area** — no title bar, ribbon, navigation pane, or chrome.
- Automatically **excludes scrollbars** from the captured area using pixel-variance analysis (works with native and DirectUI custom scrollbars).
- Seamless stitch using **template-matching overlap detection** — no visible seams.
- Saves as PNG (or JPEG) via a standard Save dialog. Files default to `Pictures\WindowSnapper\`.
- Lives quietly in the **system tray** — no main window.

## Requirements

- Windows 10 or 11 (64-bit)
- [.NET 10 Runtime](https://dotnet.microsoft.com/download/dotnet/10.0) (Windows Desktop)

## Build

```
dotnet build -c Release
```

The output is a single `.exe` in `bin\Release\net10.0-windows\`.

## Usage

1. Run `WindowSnapper.exe` — it appears in the system tray as a camera icon.
2. Left-click or right-click the tray icon and choose a snap mode.
3. The cursor becomes a crosshair and a hint bar appears at the top of the screen.
4. Hover over any window to highlight it, then **left-click** to capture.
5. Press **Escape** or **right-click** to cancel.
6. A Save dialog appears when the capture is complete.

## How It Works

| Step | Detail |
|------|--------|
| Window selection | Global low-level mouse/keyboard hooks (`WH_MOUSE_LL`, `WH_KEYBOARD_LL`). A region-clipped highlight border tracks the hovered window — no transparent overlay, so no black-window rendering bugs. |
| Content detection | `WindowFromPoint` finds the deepest scrollable child at 65 % width / 50 % height. Size-checked against the parent (≥ 1/3 area) to avoid toolbar hits. Falls back to the window client area. |
| Scrollbar exclusion | Edge strips are sampled via `BitBlt` and their luminance variance is compared to the content just inside. Near-zero variance = scrollbar track → trimmed. |
| Scroll-to-origin | `WM_VSCROLL(SB_TOP)` + 200 mouse-wheel ticks up/left before any capture. |
| Scrolling | Mouse-wheel events (`MOUSEEVENTF_WHEEL`) positioned over the content area — more universal than keyboard shortcuts, which only reach the focused child control. |
| Frame capture | `BitBlt` from the screen DC (`GetDC(NULL)`). Cursor is parked above the content area 60 ms before each grab to clear hover highlights (Explorer file rows, etc.). |
| Overlap stitching | Downsampled grayscale template matching (`FindVerticalOverlap`). Search window limited to the last frame's height in the accumulated image to prevent false matches on similar-looking rows. |
| End detection | `AreSimilar` (99 % pixel threshold). The final frame is always included — if truly identical the stitcher detects 100 % overlap and skips it harmlessly. |
