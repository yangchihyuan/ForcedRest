using System;
using System.Runtime.InteropServices;
using System.Threading;
using System.Windows.Forms;
using System.Drawing;
using Microsoft.Win32;
using System.Net; // for log on detection
using System.Globalization;
using CsvHelper;
using CsvHelper.Configuration.Attributes;
using NAudio.Wave;
using NAudio.CoreAudioApi;
using System.Speech.Synthesis;

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
    static int UsingEyeMinutes = 30;
    static int restMinutes = 10;
    static DateTime StartOfUsingEyes = DateTime.Now;      //placeholder
    static DateTime EndOfUsingEyes = DateTime.Now;      //placeholder
    static DateTime EndOfRestingEyes = DateTime.Now;    //placeholder
    static DateTime StartOfRestingEyes = DateTime.Now;    //placeholder
    static List<exceptionTime> exceptionTimes = new List<exceptionTime>();
    // Import Windows user32.dll
    [DllImport("user32.dll")]
    public static extern bool LockWorkStation();
    [DllImport("user32.dll")]
    public static extern bool SetWindowPos(IntPtr hWnd, IntPtr hWndInsertAfter, int X, int Y, int cx, int cy, uint uFlags);
    private const uint SWP_NOMOVE = 0x0002;
    private const uint SWP_NOSIZE = 0x0001;
    private const uint SWP_SHOWWINDOW = 0x0040;
    private static IntPtr HWND_TOPMOST = new IntPtr(-1);

    static double remaining = UsingEyeMinutes;
    static SmoothLabel? label;
    static bool soundPlayed = true;

    private static void PlaySoundFile(string filePath)
    {
        try
        {
            var audioFile = new AudioFileReader(filePath);
            var volumeStream = new WaveChannel32(audioFile);
            volumeStream.Volume = 3.0f;  // Increase volume (1.0f = original, 2.0f = double, etc.)
            var outputDevice = new WasapiOut(AudioClientShareMode.Shared, 100);
            outputDevice.Init(volumeStream);
            outputDevice.Play();
            outputDevice.PlaybackStopped += (sender, args) =>
            {
                outputDevice.Dispose();
                volumeStream.Dispose();
                audioFile.Dispose();
            };
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error playing sound: {ex.Message}");
        }
    }

    private static void SpeakText(string text)
    {
        try
        {
            using (var synthesizer = new SpeechSynthesizer())
            {
                synthesizer.Volume = 100;  // 0-100
                synthesizer.Rate = 0;      // -10 to 10 (speed)
                synthesizer.Speak(text);
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error speaking text: {ex.Message}");
        }
    }

    [STAThread]
    static void Main()
    {
        try
        {
            Console.WriteLine("Starting app");
            var version = System.Reflection.Assembly.GetExecutingAssembly().GetName().Version;
            Console.WriteLine($"Version: {version}");
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            SystemEvents.SessionSwitch += SystemEvents_SessionSwitch;
            SystemEvents.PowerModeChanged += SystemEvents_PowerModeChanged;
            SystemEvents.SessionEnding += SystemEvents_SessionEnding;
                    
            Form form = new Form();
            label = new SmoothLabel();
            label.Font = new Font("Arial", 36, FontStyle.Bold);
            label.ForeColor = Color.White;
            label.Cursor = Cursors.Hand;
            label.AutoSize = true;
            label.BackColor = Color.Transparent;
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


            //Load the csv file of exception times
            // Reading the file
            Console.WriteLine("Reading CSV");
            string exceptionFile = "exception_times.csv";
            if (File.Exists(exceptionFile))
            {
                using (var reader = new StreamReader(exceptionFile))
                using (var csv = new CsvReader(reader, CultureInfo.InvariantCulture))
                {
                    var records = csv.GetRecords<exceptionTime>().ToList();
                    exceptionTimes.AddRange(records);

                    foreach (var exception_time in records)
                    {
                        Console.WriteLine($"{exception_time.DayOfWeek}: {exception_time.Start} - {exception_time.End}");
                    }
                }
            }
            else
            {
                Console.WriteLine("No exception times file found.");
            }

            System.Windows.Forms.Timer timer = new System.Windows.Forms.Timer();
            timer.Interval = 1000;
            ResetCountDown();   
            timer.Tick += (s, e) =>
            {
                DateTime now = DateTime.Now;

                // [Debug Tip] Uncomment the line below to force a breakpoint here
                // if (System.Diagnostics.Debugger.IsAttached) System.Diagnostics.Debugger.Break();

                // Check for exception times
                foreach( var exTime in exceptionTimes )
                {
                    if( now.DayOfWeek == (DayOfWeek)exTime.DayOfWeek )
                    {
                        DateTime startTime = DateTime.ParseExact(exTime.Start!, "HH:mm", CultureInfo.InvariantCulture);
                        DateTime endTime = DateTime.ParseExact(exTime.End!, "HH:mm", CultureInfo.InvariantCulture);
                        DateTime startDateTime = new DateTime(now.Year, now.Month, now.Day, startTime.Hour, startTime.Minute, 0);
                        DateTime endDateTime = new DateTime(now.Year, now.Month, now.Day, endTime.Hour, endTime.Minute, 0);
                        if( now >= startDateTime && now <= endDateTime )
                        {
                            // In exception time
                            status = "Idle";
                            label.Text = "EX";
                            label.ForeColor = Color.Black;
                            return;     //return mean exit the lambda function
                        }
                    }
                }

                if (status == "UsingEyes")
                {
                    remaining = (EndOfUsingEyes - now).TotalMinutes;  
                    if (now >= EndOfUsingEyes)
                    {
                        EndOfRestingEyes = now.AddMinutes(restMinutes);
                        status = "RestingEyes";
                        LockWorkStation();
                        label.ForeColor = Color.White;

                    }
                    else if( remaining <= 1 && soundPlayed)
                    {
                        PlaySoundFile("sound material/cuckoo-9-94258 (mp3cut.net).wav");
//                         SpeakText("One minute remaining. Time to rest your eyes.");
                        label.ForeColor = Color.Red;
                        soundPlayed = false;
                    }
                }
                else if (status == "RestingEyes")
                {
                    remaining = (EndOfRestingEyes - now).TotalMinutes;
                    if (now > EndOfRestingEyes.AddSeconds(5))       //to prevent the round to 0 issue.// In Main(), add this subscription:
                    {
                        status = "Idle";
                        SpeakText("休息時間結束");
                    }
                }
                label.Text = $"{remaining:F0}";
            };
            timer.Start();

            // Ensure the form is always on top
            SetWindowPos(form.Handle, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE | SWP_SHOWWINDOW);

            Application.Run(form);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Exception: {ex.Message}");
            Console.WriteLine(ex.StackTrace);
        }
    }

    // The Event Handler
    private static void SystemEvents_SessionSwitch(object sender, SessionSwitchEventArgs e)
    {
        DateTime now = DateTime.Now;
        if (e.Reason == SessionSwitchReason.SessionUnlock)
        {
            if( status == "Idle" || now >= EndOfUsingEyes.AddMinutes(restMinutes) || now >= StartOfRestingEyes.AddMinutes(restMinutes) )
            {
                ResetCountDown();
            }
            else if( status == "RestingEyes" )
            {
                LockWorkStation();
            }
        }
        else if (e.Reason == SessionSwitchReason.SessionLock)
        {
            CalculateRestTime();
            // System is locked
            Console.WriteLine("System locked");
        }
    }

    private static void SystemEvents_PowerModeChanged(object sender, PowerModeChangedEventArgs e)
    {
        if (e.Mode == PowerModes.Suspend)
        {
            CalculateRestTime();
            // System is entering hibernation/sleep
            Console.WriteLine("System suspending (hibernation/sleep)");
        }
        else if (e.Mode == PowerModes.Resume)
        {
            // System is resuming from hibernation/sleep
            Console.WriteLine("System resuming from hibernation/sleep");
            // Optionally reset or adjust timers if needed
            //ResetCountDown();
        }
    }

    // Add this method to the Program class:
    private static void SystemEvents_SessionEnding(object sender, SessionEndingEventArgs e)
    {
        if (e.Reason == SessionEndReasons.SystemShutdown)
        {
            // The system is shutting down or rebooting.
            // Perform cleanup or save state here.
            try
            {
                using (RegistryKey key = Registry.CurrentUser.CreateSubKey(@"Software\ForcedRest"))
                {
                    key.SetValue("LastShutdownTime", DateTime.Now.ToString(CultureInfo.InvariantCulture));
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Registry error: {ex.Message}");
            }
            Console.WriteLine("System is shutting down or rebooting.");
        }
        else if (e.Reason == SessionEndReasons.Logoff)
        {
            // The user is logging off.
            Console.WriteLine("User is logging off.");
        }
    }

    private static void ResetCountDown()
    {
        DateTime now = DateTime.Now;
        StartOfUsingEyes = now;
        EndOfUsingEyes = now.AddMinutes(UsingEyeMinutes);
        status = "UsingEyes";
        soundPlayed = true;
        label!.ForeColor = Color.White;
    }

    private static void CalculateRestTime()
    {
        status = "RestingEyes";
        DateTime now = DateTime.Now;
        TimeSpan usedTime = now - StartOfUsingEyes;
        double usedMinutes = usedTime.TotalMinutes;     //bug: here may round to 0 if less than 0.5 minutes
        double expectedRestMinutes = (usedMinutes / UsingEyeMinutes) * restMinutes;
        StartOfRestingEyes = now;
        EndOfRestingEyes = now.AddMinutes(expectedRestMinutes);
    }
 
}

public class exceptionTime
{
    public int DayOfWeek { get; set; }
    [Name("Start")]
    public string? Start { get; set; }
    [Name("End")]
    public string? End { get; set; }
}

//publish command:
//dotnet publish -c Release -r win-x64 --self-contained true -p:PublishSingleFile=true -p:IncludeAllContentForSelfExtract=true
//debug run command:
//dotnet run -c Debug