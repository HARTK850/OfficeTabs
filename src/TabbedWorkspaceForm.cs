using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace OfficeTabs
{
    public class TabbedWorkspaceForm : Form
    {
        private readonly List<OfficeTabItem> _tabs = new();
        private int _selectedIndex = -1;
        private bool _isDarkMode = true;
        private readonly System.Windows.Forms.Timer _stateMonitorTimer;

        // צבעי ממשק עבור Fluent Dark/Light Modes
        private Color _bgColor;
        private Color _tabBarColor;
        private Color _activeTabColor;
        private Color _inactiveTabColor;
        private Color _textColor;
        private Color _borderColor;
        private Color _accentColor = Color.FromArgb(0, 120, 215); // Office Accent Default Blue

        private const int TabHeight = 40;
        private const int TabMinWidth = 120;
        private const int TabMaxWidth = 220;

        public TabbedWorkspaceForm()
        {
            this.Text = "Office Workspace Manager";
            this.Size = new Size(1200, 800);
            this.MinimumSize = new Size(600, 400);
            this.DoubleBuffered = true;
            this.KeyPreview = true;

            // תמיכה ב-DPI Aware
            this.AutoScaleMode = AutoScaleMode.Dpi;

            ApplyThemeColors();

            // טיימר לבקרת מצב החלונות (בדיקה שחלון לא נסגר באופן חיצוני)
            _stateMonitorTimer = new System.Windows.Forms.Timer();
            _stateMonitorTimer.Interval = 1000;
            _stateMonitorTimer.Tick += MonitorTimer_Tick;
            _stateMonitorTimer.Start();

            this.Paint += TabbedWorkspaceForm_Paint;
            this.MouseClick += TabbedWorkspaceForm_MouseClick;
            this.MouseMove += TabbedWorkspaceForm_MouseMove;
            this.SizeChanged += TabbedWorkspaceForm_SizeChanged;
            this.FormClosing += TabbedWorkspaceForm_FormClosing;
        }

        public void ToggleTheme()
        {
            _isDarkMode = !_isDarkMode;
            ApplyThemeColors();
            this.Invalidate();
        }

        private void ApplyThemeColors()
        {
            if (_isDarkMode)
            {
                _bgColor = Color.FromArgb(28, 28, 28);
                _tabBarColor = Color.FromArgb(36, 36, 36);
                _activeTabColor = Color.FromArgb(45, 45, 45);
                _inactiveTabColor = Color.FromArgb(32, 32, 32);
                _textColor = Color.White;
                _borderColor = Color.FromArgb(55, 55, 55);
            }
            else
            {
                _bgColor = Color.FromArgb(243, 243, 243);
                _tabBarColor = Color.FromArgb(230, 230, 230);
                _activeTabColor = Color.White;
                _inactiveTabColor = Color.FromArgb(220, 220, 220);
                _textColor = Color.Black;
                _borderColor = Color.FromArgb(200, 200, 200);
            }
            this.BackColor = _bgColor;
        }

        public void HandleNewOfficeWindow(IntPtr hwnd, string className)
        {
            // הימנעות מלכידה כפולה
            if (_tabs.Exists(t => t.Hwnd == hwnd))
                return;

            string docTitle = Win32.GetWindowText(hwnd);
            if (string.IsNullOrEmpty(docTitle))
            {
                docTitle = className switch
                {
                    "OpusApp" => "Word Document",
                    "XLMAIN" => "Excel Sheet",
                    "PPTFrameClass" => "PowerPoint Presentation",
                    _ => "Office Document"
                };
            }

            // שמירת המאפיינים המקוריים של החלון לצורך שחזורו בעתיד
            long originalStyle = Win32.GetWindowLongPtr(hwnd, Win32.GWL_STYLE);
            long originalExStyle = Win32.GetWindowLongPtr(hwnd, Win32.GWL_EXSTYLE);

            // לכידת החלון - הפיכתו לחלון בן (Child) של הטופס הנוכחי
            Win32.SetParent(hwnd, this.Handle);

            // הסרת גבולות, מסגרות, וכותרות כדי שהסגנון יהיה Fluent ונקי
            long style = originalStyle;
            style &= ~Win32.WS_CAPTION;
            style &= ~Win32.WS_THICKFRAME;
            style &= ~Win32.WS_MINIMIZEBOX;
            style &= ~Win32.WS_MAXIMIZEBOX;
            style &= ~Win32.WS_SYSMENU;
            Win32.SetWindowLongPtr(hwnd, Win32.GWL_STYLE, style);

            // עדכון המיקום והגודל של החלון החדש
            PositionEmbeddedWindow(hwnd);

            // הוספה לרשימת הכרטיסיות
            OfficeTabItem tab = new OfficeTabItem(hwnd, docTitle, className, originalStyle, originalExStyle);
            _tabs.Add(tab);

            _selectedIndex = _tabs.Count - 1;
            UpdateActiveTabVisibility();
            this.Invalidate();
        }

        private void PositionEmbeddedWindow(IntPtr hwnd)
        {
            if (hwnd == IntPtr.Zero) return;

            // המיקום של החלון הלכוד מתחיל מתחת לפס הכרטיסיות
            int x = 0;
            int y = TabHeight;
            int w = this.ClientSize.Width;
            int h = this.ClientSize.Height - TabHeight;

            Win32.MoveWindow(hwnd, x, y, w, h, true);
            Win32.ShowWindow(hwnd, Win32.SW_SHOWMAXIMIZED);
        }

        private void UpdateActiveTabVisibility()
        {
            for (int i = 0; i < _tabs.Count; i++)
            {
                if (i == _selectedIndex)
                {
                    Win32.ShowWindow(_tabs[i].Hwnd, Win32.SW_SHOW);
                    PositionEmbeddedWindow(_tabs[i].Hwnd);
                    Win32.SetFocus(_tabs[i].Hwnd);
                    
                    // עדכון צבע מותאם אישית לפי סוג המסמך (כחול ל-Word, ירוק ל-Excel, כתום ל-PPT)
                    _accentColor = _tabs[i].ClassName switch
                    {
                        "OpusApp" => Color.FromArgb(43, 87, 154),
                        "XLMAIN" => Color.FromArgb(33, 115, 70),
                        "PPTFrameClass" => Color.FromArgb(183, 71, 42),
                        _ => Color.FromArgb(0, 120, 215)
                    };
                }
                else
                {
                    Win32.ShowWindow(_tabs[i].Hwnd, Win32.SW_HIDE);
                }
            }
        }

        private void MonitorTimer_Tick(object? sender, EventArgs e)
        {
            bool dynamicUpdateRequired = false;

            for (int i = _tabs.Count - 1; i >= 0; i--)
            {
                // אם חלון ה-Office נסגר עצמאית על ידי המשתמש (למשל Alt+F4 או קובץ->יציאה)
                if (!Win32.IsWindow(_tabs[i].Hwnd))
                {
                    _tabs.RemoveAt(i);
                    dynamicUpdateRequired = true;
                    if (_selectedIndex >= _tabs.Count)
                    {
                        _selectedIndex = _tabs.Count - 1;
                    }
                }
                else
                {
                    // עדכון דינמי של שם הכותרת אם הוא השתנה (למשל בגלל "שמור בשם")
                    string currentTitle = Win32.GetWindowText(_tabs[i].Hwnd);
                    if (!string.IsNullOrEmpty(currentTitle) && currentTitle != _tabs[i].Title)
                    {
                        _tabs[i].Title = currentTitle;
                        dynamicUpdateRequired = true;
                    }
                }
            }

            if (dynamicUpdateRequired)
            {
                UpdateActiveTabVisibility();
                this.Invalidate();
            }
        }

        private void TabbedWorkspaceForm_Paint(object? sender, PaintEventArgs e)
        {
            Graphics g = e.Graphics;
            g.SmoothingMode = SmoothingMode.AntiAlias;

            // ציור רקע ה-Tab Bar
            using (SolidBrush barBrush = new SolidBrush(_tabBarColor))
            {
                g.FillRectangle(barBrush, 0, 0, this.ClientSize.Width, TabHeight);
            }

            if (_tabs.Count == 0)
            {
                // הודעה למשתמש כאשר אין מסמכים פתוחים
                string emptyMsg = "פתח קובץ Word, Excel או PowerPoint כדי להתחיל לעבוד עם כרטיסיות.";
                using Font font = new Font("Segoe UI", 11F, FontStyle.Regular);
                using SolidBrush textBrush = new SolidBrush(_textColor);
                SizeF size = g.MeasureString(emptyMsg, font);
                g.DrawString(emptyMsg, font, textBrush, (this.ClientSize.Width - size.Width) / 2, (this.ClientSize.Height - size.Height) / 2);
                return;
            }

            // חישוב רוחב הכרטיסיות בצורה דינמית בהתאם למספר המסמכים הפתוחים
            int tabWidth = Math.Min(TabMaxWidth, Math.Max(TabMinWidth, this.ClientSize.Width / _tabs.Count));

            for (int i = 0; i < _tabs.Count; i++)
            {
                int x = i * tabWidth;
                _tabs[i].Bounds = new Rectangle(x, 0, tabWidth, TabHeight);

                bool isActive = (i == _selectedIndex);
                Color tabBg = isActive ? _activeTabColor : _inactiveTabColor;

                // ציור רקע הכרטיסייה
                using (SolidBrush tabBrush = new SolidBrush(tabBg))
                {
                    g.FillRectangle(tabBrush, _tabs[i].Bounds);
                }

                // קו עליון דקורטיבי ברוחב 3 פיקסלים המציין את צבע התוכנה הפתוחה
                if (isActive)
                {
                    using Pen activePen = new Pen(_accentColor, 3);
                    g.DrawLine(activePen, x, 1, x + tabWidth, 1);
                }

                // ציור כפתור סגירה (X)
                Rectangle closeRect = new Rectangle(x + tabWidth - 24, (TabHeight - 16) / 2, 16, 16);
                _tabs[i].CloseButtonBounds = closeRect;
                
                using (Pen xPen = new Pen(_textColor, 1.5f))
                {
                    g.DrawLine(xPen, closeRect.X + 4, closeRect.Y + 4, closeRect.Right - 4, closeRect.Bottom - 4);
                    g.DrawLine(xPen, closeRect.Right - 4, closeRect.Y + 4, closeRect.X + 4, closeRect.Bottom - 4);
                }

                // ציור כותרת המסמך
                string drawText = _tabs[i].Title;
                using Font font = new Font("Segoe UI", 9F, isActive ? FontStyle.Bold : FontStyle.Regular);
                using SolidBrush textBrush = new SolidBrush(_textColor);
                
                // קיצור הכותרת אם היא ארוכה מדי
                Rectangle textRect = new Rectangle(x + 10, 0, tabWidth - 38, TabHeight);
                TextRenderer.DrawText(g, drawText, font, textRect, _textColor, TextFormatFlags.VerticalCenter | TextFormatFlags.EndEllipsis);

                // קו מפריד בין כרטיסיות
                if (!isActive && i < _tabs.Count - 1)
                {
                    using Pen sepPen = new Pen(_borderColor, 1);
                    g.DrawLine(sepPen, x + tabWidth - 1, 8, x + tabWidth - 1, TabHeight - 8);
                }
            }
        }

        private void TabbedWorkspaceForm_MouseClick(object? sender, MouseEventArgs e)
        {
            for (int i = 0; i < _tabs.Count; i++)
            {
                if (_tabs[i].Bounds.Contains(e.Location))
                {
                    // בדיקה אם לחצו על כפתור הסגירה (X)
                    if (_tabs[i].CloseButtonBounds.Contains(e.Location))
                    {
                        CloseOfficeDocument(_tabs[i]);
                        return;
                    }

                    // מעבר לכרטיסייה שנבחרה
                    _selectedIndex = i;
                    UpdateActiveTabVisibility();
                    this.Invalidate();
                    return;
                }
            }
        }

        private void TabbedWorkspaceForm_MouseMove(object? sender, MouseEventArgs e)
        {
            // שינוי עיצוב סמן העכבר כאשר מרחפים מעל כפתור סגירה של טאב
            bool overCloseButton = false;
            foreach (var tab in _tabs)
            {
                if (tab.CloseButtonBounds.Contains(e.Location))
                {
                    overCloseButton = true;
                    break;
                }
            }
            this.Cursor = overCloseButton ? Cursors.Hand : Cursors.Default;
        }

        private void CloseOfficeDocument(OfficeTabItem tab)
        {
            // שליחת בקשת סגירה לחלון ה-Office באמצעות WM_CLOSE
            Win32.PostMessage(tab.Hwnd, Win32.WM_CLOSE, IntPtr.Zero, IntPtr.Zero);
            
            // הסרה מרשימת הטאבים
            _tabs.Remove(tab);
            if (_selectedIndex >= _tabs.Count)
            {
                _selectedIndex = _tabs.Count - 1;
            }
            UpdateActiveTabVisibility();
            this.Invalidate();
        }

        private void TabbedWorkspaceForm_SizeChanged(object? sender, EventArgs e)
        {
            if (_selectedIndex >= 0 && _selectedIndex < _tabs.Count)
            {
                PositionEmbeddedWindow(_tabs[_selectedIndex].Hwnd);
            }
            this.Invalidate();
        }

        public void RestoreAllCapturedWindows()
        {
            foreach (var tab in _tabs)
            {
                if (Win32.IsWindow(tab.Hwnd))
                {
                    // שחרור החלון מהסרת ה-Parenting
                    Win32.SetParent(tab.Hwnd, IntPtr.Zero);
                    
                    // שחזור הסטייל המקורי
                    Win32.SetWindowLongPtr(tab.Hwnd, Win32.GWL_STYLE, tab.OriginalStyle);
                    Win32.SetWindowLongPtr(tab.Hwnd, Win32.GWL_EXSTYLE, tab.OriginalExStyle);

                    // עדכון המיקום והגודל של החלון ששוחרר חזרה למצב רגיל
                    Win32.ShowWindow(tab.Hwnd, Win32.SW_SHOWNORMAL);
                }
            }
            _tabs.Clear();
        }

        private void TabbedWorkspaceForm_FormClosing(object? sender, FormClosingEventArgs e)
        {
            // במקום לסגור את האפליקציה, אנחנו רק מחביאים את החלון הראשי ל-Tray
            if (e.CloseReason == CloseReason.UserClosing)
            {
                e.Cancel = true;
                this.Hide();
            }
            else
            {
                RestoreAllCapturedWindows();
            }
        }
    }

    public class OfficeTabItem
    {
        public IntPtr Hwnd { get; }
        public string Title { get; set; }
        public string ClassName { get; }
        public long OriginalStyle { get; }
        public long OriginalExStyle { get; }
        public Rectangle Bounds { get; set; }
        public Rectangle CloseButtonBounds { get; set; }

        public OfficeTabItem(IntPtr hwnd, string title, string className, long originalStyle, long originalExStyle)
        {
            Hwnd = hwnd;
            Title = title;
            ClassName = className;
            OriginalStyle = originalStyle;
            OriginalExStyle = originalExStyle;
        }
    }
}
