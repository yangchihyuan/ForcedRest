using System;
using System.Runtime.InteropServices;
using System.Threading;
using System.Windows.Forms;
using System.Drawing;
using Microsoft.Win32;
using System.Net; // for log on detection

class Program
{
    static string status = "EnterUsingEyes";
    static int UsingEyeMinutes = 30;
    static int restMinutes = 10;
    static DateTime EndOfUsingEyes = DateTime.Now;      //placeholder
    static DateTime EndOfRestingEyes = DateTime.Now;    //placeholder
    // Import Windows user32.dll
    [DllImport("user32.dll")]
    public static extern bool LockWorkStation();

    [STAThread]
    static void Main()
    {
        Application.EnableVisualStyles();
        Application.SetCompatibleTextRenderingDefault(false);
        SystemEvents.SessionSwitch += SystemEvents_SessionSwitch;

        Form form = new Form();
        Label label = new Label();
        label.Font = new Font("Arial", 24, FontStyle.Bold);
        label.ForeColor = Color.White;
        label.AutoSize = true;
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
        form.Size = new Size(300, 50);

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
        timer.Tick += (s, e) =>
        {
            if (status == "EnterUsingEyes")
            {
                EndOfUsingEyes = DateTime.Now.AddMinutes(UsingEyeMinutes);
                remaining = (EndOfUsingEyes - DateTime.Now).TotalMinutes;
                status = "UsingEyes";
            }
            else if (status == "UsingEyes")
            {
                remaining = (EndOfUsingEyes - DateTime.Now).TotalMinutes;
                if (DateTime.Now >= EndOfUsingEyes)
                {
                    status = "EnterRestingEyes";
                }
            }
            else if (status == "EnterRestingEyes")
            {
                EndOfRestingEyes = DateTime.Now.AddMinutes(restMinutes);
                status = "RestingEyes";
                LockWorkStation();
            }
            else if (status == "RestingEyes")
            {
                remaining = (EndOfRestingEyes - DateTime.Now).TotalMinutes;
                if (DateTime.Now >= EndOfRestingEyes)
                {
                    status = "Idle";
                }
            }
            label.Text = $"{status} {remaining:F0}";
        };
        timer.Start();

        Application.Run(form);
    }

    // The Event Handler
    private static void SystemEvents_SessionSwitch(object sender, SessionSwitchEventArgs e)
    {
        if (e.Reason == SessionSwitchReason.SessionUnlock)
        {
            if( status == "Idle" )
            {
                status = "EnterUsingEyes";
            }
            else if( status == "RestingEyes" )
            {
                LockWorkStation();
            }
        }
    }
 
}
