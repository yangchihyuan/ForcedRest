using System;
using System.Runtime.InteropServices;
using System.Threading;
using System.Windows.Forms;
using System.Drawing;

class Program
{
    // 匯入 Windows user32.dll 函式庫
    [DllImport("user32.dll")]
    public static extern bool LockWorkStation();

    [STAThread]
    static void Main()
    {
        Application.EnableVisualStyles();
        Application.SetCompatibleTextRenderingDefault(false);

        Form form = new Form();
        Label label = new Label();
        label.Font = new Font("Arial", 24, FontStyle.Bold);
        label.ForeColor = Color.White;
        label.AutoSize = true;
        label.Text = "Initializing...";
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
        form.Size = new Size(200, 50);

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

        // Timer logic
        DateTime ExpectRestTime = DateTime.Now.AddMinutes(30);
        DateTime ExpectUnlockTime = DateTime.Now;       //placeholder
        string status = "UsingEyes";

        //Timer timer = new Timer();
        System.Windows.Forms.Timer timer = new System.Windows.Forms.Timer();
        timer.Interval = 1000;
        timer.Tick += (s, e) =>
        {
            double remaining;
            if (status == "UsingEyes")
            {
                remaining = (ExpectRestTime - DateTime.Now).TotalMinutes;
                if (DateTime.Now >= ExpectRestTime)
                {
                    status = "RestingEyes";
                    LockWorkStation();
                    ExpectUnlockTime = DateTime.Now.AddMinutes(10);
                }
            }
            else
            {
                remaining = (ExpectUnlockTime - DateTime.Now).TotalMinutes;
                if (DateTime.Now >= ExpectUnlockTime)
                {
                    ExpectRestTime = DateTime.Now.AddMinutes(30);
                    status = "UsingEyes";
                }
                else
                {
                    LockWorkStation();
                    remaining = (ExpectUnlockTime - DateTime.Now).TotalMinutes;
                }
            }
            label.Text = $"{remaining:F0}";
        };
        timer.Start();

        // 讓 Label 也能觸發拖動
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
        Application.Run(form);
    }
}
