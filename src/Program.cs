using System;
using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;
using System.Threading;
using System.Windows.Forms;
using Microsoft.Win32;

namespace OfficeTabs
{
    public static class Program
    {
        private static Mutex? _mutex;
        private static NotifyIcon? _trayIcon;
        private static ContextMenuStrip? _trayMenu;
        private static TabbedWorkspaceForm? _workspaceForm;
        private static IntPtr _windowHookHandle;
        private static Win32.WinEventDelegate? _hookDelegate;

        [STAThread]
        public static void Main()
        {
            // מניעת הרצה של יותר ממופע אחד במקביל
            _mutex = new Mutex(true, "Global\\OfficeTabsUniqueMutexName_2026", out bool createdNew);
            if (!createdNew)
            {
                MessageBox.Show("אפליקציית Office Tabs כבר רצה ברקע.", "מידע", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            ApplicationConfiguration.Initialize();
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);

            InitializeAutoStart();
            InitializeTrayIcon();
            InitializeHook();

            _workspaceForm = new TabbedWorkspaceForm();
            
            // הרצת לולאת ההודעות של Windows
            Application.Run();

            Cleanup();
        }

        private static void InitializeAutoStart()
        {
            try
            {
                string runKey = @"SOFTWARE\Microsoft\Windows\CurrentVersion\Run";
                using RegistryKey? key = Registry.CurrentUser.OpenSubKey(runKey, true);
                if (key != null)
                {
                    string? exePath = Process.GetCurrentProcess().MainModule?.FileName;
                    if (!string.IsNullOrEmpty(exePath))
                    {
                        key.SetValue("OfficeTabs", $"\"{exePath}\"");
                    }
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"שגיאה ברישום להפעלה אוטומטית: {ex.Message}");
            }
        }

        private static void InitializeTrayIcon()
        {
            _trayMenu = new ContextMenuStrip();
            _trayMenu.Items.Add("הצג מרכז שליטה", null, OnShowDashboard);
            _trayMenu.Items.Add("החלף ערכת נושא (כהה/בהיר)", null, OnToggleTheme);
            _trayMenu.Items.Add("-");
            _trayMenu.Items.Add("יציאה", null, OnExit);

            _trayIcon = new NotifyIcon
            {
                Text = "Office Tabs Workspace Manager",
                Icon = SystemIcons.Application,
                ContextMenuStrip = _trayMenu,
                Visible = true
            };

            _trayIcon.DoubleClick += OnShowDashboard;
        }

        private static void InitializeHook()
        {
            _hookDelegate = new Win32.WinEventDelegate(WinEventProc);
            
            // האזנה לאירועי יצירת חלונות והבאתם לקדמת הבמה ברמת מערכת ההפעלה
            _windowHookHandle = Win32.SetWinEventHook(
                Win32.EVENT_OBJECT_SHOW,
                Win32.EVENT_SYSTEM_FOREGROUND,
                IntPtr.Zero,
                _hookDelegate,
                0,
                0,
                Win32.WINEVENT_OUTOFCONTEXT
            );
        }

        private static void WinEventProc(IntPtr hWinEventHook, uint eventType, IntPtr hwnd, int idObject, int idChild, uint dwEventThread, uint dwmsEventTime)
        {
            if (idObject != Win32.OBJID_WINDOW || idChild != Win32.CHILDID_SELF || hwnd == IntPtr.Zero)
                return;

            // זיהוי אם החלון שנוצר או הופעל שייך ל-Word, Excel או PowerPoint
            string className = Win32.GetClassName(hwnd);
            if (IsOfficeWindowClass(className))
            {
                // הרצה בצורה אסינכרונית כדי לא לתקוע את ה-Hook thread של ה-OS
                _workspaceForm?.BeginInvoke(new Action(() =>
                {
                    _workspaceForm.HandleNewOfficeWindow(hwnd, className);
                }));
            }
        }

        private static bool IsOfficeWindowClass(string className)
        {
            // OpusApp = Word, XLMAIN = Excel, PPTFrameClass = PowerPoint
            return className == "OpusApp" || className == "XLMAIN" || className == "PPTFrameClass";
        }

        private static void OnShowDashboard(object? sender, EventArgs e)
        {
            if (_workspaceForm != null)
            {
                _workspaceForm.Show();
                _workspaceForm.WindowState = FormWindowState.Normal;
                _workspaceForm.BringToFront();
            }
        }

        private static void OnToggleTheme(object? sender, EventArgs e)
        {
            _workspaceForm?.ToggleTheme();
        }

        private static void OnExit(object? sender, EventArgs e)
        {
            Cleanup();
            Application.Exit();
        }

        private static void Cleanup()
        {
            if (_windowHookHandle != IntPtr.Zero)
            {
                Win32.UnhookWinEvent(_windowHookHandle);
                _windowHookHandle = IntPtr.Zero;
            }

            if (_trayIcon != null)
            {
                _trayIcon.Visible = false;
                _trayIcon.Dispose();
                _trayIcon = null;
            }

            if (_workspaceForm != null && !_workspaceForm.IsDisposed)
            {
                _workspaceForm.RestoreAllCapturedWindows();
                _workspaceForm.Dispose();
            }

            _mutex?.ReleaseMutex();
            _mutex?.Dispose();
        }
    }
}
