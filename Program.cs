using System.Windows.Forms;
using WindowSnapper;

// Enable Per-Monitor DPI awareness so GetWindowRect coords match BitBlt coords.
try { NativeMethods.SetProcessDpiAwareness(2); } catch { /* pre-Win8 */ }

Application.SetHighDpiMode(HighDpiMode.PerMonitorV2);
Application.EnableVisualStyles();
Application.SetCompatibleTextRenderingDefault(false);
Application.Run(new SysTrayApp());
