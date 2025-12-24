using System;
using System.Runtime.InteropServices;
using System.Threading;
using System.Windows.Forms;
using System.Drawing;
using Microsoft.Win32;
using System.Net; // for log on detection

class SmoothLabel : Label
{
    protected override void OnPaint(PaintEventArgs e)
    {
        e.Graphics.TextRenderingHint = System.Drawing.Text.TextRenderingHint.AntiAlias;
        base.OnPaint(e);
    }
}

class Program
{
    static string status = "EnterUsingEyes";
    static int UsingEyeMinutes = 1;
    static int restMinutes = 1;
    static DateTime EndOfUsingEyes = DateTime.Now;      //placeholder
    static DateTime EndOfRestingEyes = DateTime.Now;    //placeholder
    // Import Windows user32.dll
    [DllImport("user32.dll")]
    public static extern bool LockWorkStation();
    [DllImport("user32.dll")]
    public static extern bool SetWindowPos(IntPtr hWnd, IntPtr hWndInsertAfter, int X, int Y, int cx, int cy, uint uFlags);
    [DllImport("winmm.dll")]
    private static extern bool PlaySound(string pszSound, IntPtr hmod, uint fdwSound);
    private const uint SWP_NOMOVE = 0x0002;
    private const uint SWP_NOSIZE = 0x0001;
    private const uint SWP_SHOWWINDOW = 0x0040;
    private const uint SND_FILENAME = 0x00020000;
    private const uint SND_ASYNC = 0x0001;
    private static IntPtr HWND_TOPMOST = new IntPtr(-1);

    [STAThread]
    static void Main()
    {
        Application.EnableVisualStyles();
        Application.SetCompatibleTextRenderingDefault(false);
        SystemEvents.SessionSwitch += SystemEvents_SessionSwitch;

        Form form = new Form();
        SmoothLabel label = new SmoothLabel();
        label.Font = new Font("Arial", 36, FontStyle.Bold);
        label.ForeColor = Color.White;
        label.Cursor = Cursors.Hand;
        label.AutoSize = true;
        label.BackColor = Color.Transparent; // 使标签背景透明
        form.Controls.Add(label);
        form.TopMost = true;
        form.FormBorderStyle = FormBorderStyle.None;
        form.BackColor = Color.Black;
        form.TransparencyKey = Color.Black;
        form.ShowInTaskbar = false;
        form.StartPosition = FormStartPosition.Manual;
        var screen = Screen.PrimaryScreen;
        if (screen != null)
        {
            form.Location = new Point(screen.WorkingArea.Width / 2 - 50, 10);
        }
        form.Size = new Size(250, 50);

        // Make draggable
        bool dragging = false;
        Point dragCursorPoint = Point.Empty;
        Point dragFormPoint = Point.Empty;
        form.MouseDown += (s, e) =>
        {
            dragging = true;
            dragCursorPoint = Cursor.Position;
            dragFormPoint = form.Location;
        };
        form.MouseMove += (s, e) =>
        {
            if (dragging)
            {
                Point diff = Point.Subtract(Cursor.Position, new Size(dragCursorPoint));
                form.Location = Point.Add(dragFormPoint, new Size(diff));
            }
        };
        form.MouseUp += (s, e) =>
        {
            dragging = false;
        };


        // Enable Label drag to move the form
        label.MouseDown += (s, e) =>
        {
            dragging = true;
            dragCursorPoint = Cursor.Position;
            dragFormPoint = form.Location;
        };

        label.MouseMove += (s, e) =>
        {
            if (dragging)
            {
                Point diff = Point.Subtract(Cursor.Position, new Size(dragCursorPoint));
                form.Location = Point.Add(dragFormPoint, new Size(diff));
            }
        };

        label.MouseUp += (s, e) =>
        {
            dragging = false;
        };


        System.Windows.Forms.Timer timer = new System.Windows.Forms.Timer();
        timer.Interval = 1000;
        double remaining = UsingEyeMinutes;
        bool soundPlayed = false;
        timer.Tick += (s, e) =>
        {
            DateTime now = DateTime.Now;
            if (status == "EnterUsingEyes")
            {
                EndOfUsingEyes = now.AddMinutes(UsingEyeMinutes);
                remaining = (EndOfUsingEyes - now).TotalMinutes;
                status = "UsingEyes";
                soundPlayed = true;
                label.ForeColor = Color.White;
            }
            else if (status == "UsingEyes")
            {
                remaining = (EndOfUsingEyes - now).TotalMinutes;
                if (now >= EndOfUsingEyes)
                {
                    status = "EnterRestingEyes";
                }
                else if( remaining <= 1 && soundPlayed)
                {
                    PlaySound("sound material/cuckoo-9-94258 (mp3cut.net).wav", IntPtr.Zero, SND_FILENAME | SND_ASYNC);
                    label.ForeColor = Color.Red;
                    soundPlayed = false;
                }
//                SystemSounds.Beep.Play();
            }
            else if (status == "EnterRestingEyes")
            {
                EndOfRestingEyes = now.AddMinutes(restMinutes);
                status = "RestingEyes";
                LockWorkStation();
                label.ForeColor = Color.White;
            }
            else if (status == "RestingEyes")
            {
                remaining = (EndOfRestingEyes - now).TotalMinutes;
                if (now >= EndOfRestingEyes)
                {
                    status = "Idle";
                    PlaySound("sound material/10Minutes.wav", IntPtr.Zero, SND_FILENAME | SND_ASYNC);
                }
            }
            label.Text = $"{remaining:F0}";
        };
        timer.Start();

        // Ensure the form is always on top
        SetWindowPos(form.Handle, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE | SWP_SHOWWINDOW);

        Application.Run(form);
    }

    // The Event Handler
    private static void SystemEvents_SessionSwitch(object sender, SessionSwitchEventArgs e)
    {
        DateTime now = DateTime.Now;
        if (e.Reason == SessionSwitchReason.SessionUnlock)
        {
            if( status == "Idle" || now >= EndOfUsingEyes.AddMinutes(restMinutes) || now >= EndOfRestingEyes )
            {
                // Play sound when rest period ends and user unlocks
                status = "EnterUsingEyes";
            }
            else if( status == "RestingEyes" )
            {
                LockWorkStation();
            }
            else if( now >= EndOfUsingEyes.AddMinutes(restMinutes) )
            {
                status = "EnterUsingEyes";
            }
        }
        else if (e.Reason == SessionSwitchReason.SessionLock)
        {
            EndOfRestingEyes = now.AddMinutes(restMinutes);
        }
    }
 
}

//publish command:
//dotnet publish -c Release -r win-x64 --self-contained true -p:PublishSingleFile=true -p:IncludeAllContentForSelfExtract=true