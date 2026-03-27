using System;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;
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
    public SmoothLabel()
    {
        this.SetStyle(ControlStyles.StandardDoubleClick, true);
    }

    protected override void OnPaint(PaintEventArgs e)
    {
        e.Graphics.TextRenderingHint = System.Drawing.Text.TextRenderingHint.AntiAlias;
        base.OnPaint(e);
    }
}

class Program
{
    static string status = "UsingEyes";
    static int UsingEyeMinutes = 30;
    static int restMinutes = 10;
    static DateTime StartOfUsingEyes = DateTime.Now;      //placeholder
    static DateTime EndOfUsingEyes = StartOfUsingEyes.AddMinutes(UsingEyeMinutes);
    static TimeSpan UsingEyeTimeSpan = TimeSpan.FromMinutes(0); //placeholder
    static DateTime StartOfRestingEyes = DateTime.MinValue;    //placeholder
    static DateTime EndOfRestingEyes = DateTime.MinValue;    //placeholder
    static DateTime LogoutTime = DateTime.MinValue;    //placeholder
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

    static int NumberOfExtensions = 0;

    private static readonly string LogDir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "ForcedRest");
    private static readonly string LogFilePath = Path.Combine(LogDir, "forcedrest.log");

    static void WriteLog(string message)
    {
        Console.WriteLine(message);
        try
        {
            Directory.CreateDirectory(LogDir);
            File.AppendAllText(LogFilePath, $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] {message}{Environment.NewLine}");
        }
        catch
        {
            // Ignore logging errors; app should continue.
        }
    }


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
        Task.Run(async () =>
        {
            try
            {
                using (var stream = new System.IO.MemoryStream())
                {
                    using (var synthesizer = new SpeechSynthesizer())
                    {
                        synthesizer.SetOutputToWaveStream(stream);
                        synthesizer.Volume = 100;
                        synthesizer.Rate = 0;
                        synthesizer.Speak(text);
                    }
                    stream.Position = 0;
                    using (var waveReader = new WaveFileReader(stream))
                    using (var volumeStream = new WaveChannel32(waveReader))
                    using (var outputDevice = new WasapiOut(AudioClientShareMode.Shared, 100))
                    {
                        volumeStream.Volume = 3.0f; // Increase volume (3.0f = 300%)
                        var tcs = new TaskCompletionSource<bool>();
                        outputDevice.PlaybackStopped += (s, e) => tcs.TrySetResult(true);
                        outputDevice.Init(volumeStream);
                        outputDevice.Play();
                        await tcs.Task;
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error speaking text: {ex.Message}");
            }
        });
    }

    [STAThread]
    static void Main()
    {
        try
        {
            WriteLog("Application started");

            var version = System.Reflection.Assembly.GetExecutingAssembly().GetName().Version;
            WriteLog($"Version: {version}");
//            Console.WriteLine($"Version: {version}");
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            SystemEvents.SessionSwitch += SystemEvents_SessionSwitch;
            SystemEvents.PowerModeChanged += SystemEvents_PowerModeChanged;
            SystemEvents.SessionEnding += SystemEvents_SessionEnding;

            //load the Registry to get the previous EndOfRestingEyes
            using (RegistryKey? key = Registry.CurrentUser.OpenSubKey(@"Software\ForcedRest"))
            {
                if (key != null)
                {
                    object? value = key.GetValue("EndOfRestingEyes");
                    if (value != null)
                    {
                        string endOfRestingEyesStr = value.ToString()!;
                        EndOfRestingEyes = DateTime.Parse(endOfRestingEyesStr, CultureInfo.InvariantCulture);
                    }

                    value = null;
                    value = key.GetValue("EndOfUsingEyes");
                    if (value != null)
                    {
                        string endOfUsingEyesStr = value.ToString()!;
                        EndOfUsingEyes = DateTime.Parse(endOfUsingEyesStr, CultureInfo.InvariantCulture);
                        WriteLog($"Loaded EndOfUsingEyes from registry: {EndOfUsingEyes}");
                    }

                    value = null;
                    value = key.GetValue("LogoutTime");
                    if (value != null)
                    {
                        string logoutTimeStr = value.ToString()!;
                        LogoutTime = DateTime.Parse(logoutTimeStr, CultureInfo.InvariantCulture);
                    }

                }
            }

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
            form.Size = new Size(50, 50);

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

            label.DoubleClick += (s, e) =>
            {
                if (remaining < 1 && status == "UsingEyes" && NumberOfExtensions < 12)
                {
                    EndOfUsingEyes += TimeSpan.FromMinutes(1);
                    NumberOfExtensions += 1;
                    soundPlayed = true;
                    label.ForeColor = Color.White;
                }
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
                        exception_time.startTime = DateTime.ParseExact(exception_time.Start!, "HH:mm", CultureInfo.InvariantCulture);
                        exception_time.endTime = DateTime.ParseExact(exception_time.End!, "HH:mm", CultureInfo.InvariantCulture);
                    }
                }
            }
            else
            {
                Console.WriteLine("No exception times file found.");
            }

            System.Windows.Forms.Timer timer = new System.Windows.Forms.Timer();
            timer.Interval = 1000;

            //Judge if need to reset the countdown
            DateTime now = DateTime.Now;
            if( now >= EndOfRestingEyes )
            {
                ResetCountDown();
            }
            else
            {
                status = "RestingEyes";
                LockWorkStation();
            }


            timer.Tick += (s, e) =>
            {
                DateTime now = DateTime.Now;

                // Check for exception times
                foreach( var exTime in exceptionTimes )
                {
                    if( now.DayOfWeek == (DayOfWeek)exTime.DayOfWeek )
                    {
                        // Compare using TimeOfDay to ignore Date components
                        if (now.TimeOfDay >= exTime.startTime.TimeOfDay && now.TimeOfDay < exTime.endTime.TimeOfDay)
                        {
                            // In exception time
                            status = "ExceptionTime";
                            label.Text = "EX";
                            label.ForeColor = Color.Black;
                            return;     //return mean exit the lambda function
                        }
                        else if( status == "ExceptionTime" )
                        {
                            // Exiting exception time
                            ResetCountDown();
                        }
                    }
                }

                if (status == "UsingEyes")
                {
                    remaining= (EndOfUsingEyes - now).TotalMinutes;  
                    if (now >= EndOfUsingEyes)
                    {
                        LockWorkStation();
                    }
                    else if( remaining <= 1 )
                    {
                        if(soundPlayed)
                        {
                            PlaySoundFile("sound material/cuckoo-9-94258 (mp3cut.net).wav");
                            label.ForeColor = Color.Red;
                            soundPlayed = false;
                        }
                        TimeSpan ts = EndOfUsingEyes - now;
                        if (NumberOfExtensions > 0)
                            label.Text = $"{ts.Seconds:D2}+{NumberOfExtensions}";
                        else
                            label.Text = $"{ts.Seconds:D2}";
                    }
                    else{
                        if (NumberOfExtensions > 0)
                            label.Text = $"{remaining:F0}+{NumberOfExtensions}";
                        else
                            label.Text = $"{remaining:F0}";
                    }
                }
                else if (status == "RestingEyes")
                {
                    label.ForeColor = Color.Yellow;
                    if (now > EndOfRestingEyes.AddSeconds(1))       //to prevent the round to 0 issue.// In Main(), add this subscription:
                    {
                        status = "Idle";
                        Console.WriteLine(now.ToString() + " Change to Idle");
                        SpeakText("休息時間結束");
                    }
                    TimeSpan ts = EndOfRestingEyes - now;
                    if (ts.TotalSeconds < 0) ts = TimeSpan.Zero;
                    label.Text = $"{(int)ts.TotalMinutes}:{ts.Seconds:D2}";
                }
                //Console.WriteLine(now.ToString() + " Status: " + status + $" Remaining: {remaining:F2} minutes");
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
            if( status == "Idle" || now >= EndOfRestingEyes )
            {
                ResetCountDown();
                Console.WriteLine(now.ToString() + "System Unlocked");
            }
            else if( status == "RestingEyes" )
            {
                LockWorkStation();
            }
        }
        else if (e.Reason == SessionSwitchReason.SessionLock)
        {
            //When I call LockWorkStation(), it will trigger this event.
            if( status == "UsingEyes" )
            {
                UsingEyeTimeSpan = now - StartOfUsingEyes;
                EnterRestMode();
            }
            else if( status == "RestingEyes" )
            {
                // Do nothing, continue resting
            }
            else if( status == "Idle" )
            {
                // Do nothing
            }
            // System is locked
            // If Suspend the system, will this event be triggered first?
            Console.WriteLine("System locked" + now.ToString());
        }
    }

    private static void SystemEvents_PowerModeChanged(object sender, PowerModeChangedEventArgs e)
    {
        if (e.Mode == PowerModes.Suspend)
        {
            if( status == "UsingEyes" )
            {
                UsingEyeTimeSpan = DateTime.Now - StartOfUsingEyes;
                //If I still have time remaining, save them into the registry, and when system resume, I can decide whether to continue the countdown or reset it.
                try
                {
                    using (RegistryKey key = Registry.CurrentUser.CreateSubKey(@"Software\ForcedRest"))
                    {
                        key.SetValue("StartOfUsingEyes", StartOfUsingEyes.ToString(CultureInfo.InvariantCulture));
                        WriteLog($"Suspendt: Saved EndOfRestingEyes to registry: {EndOfRestingEyes}");
                        key.SetValue("SuspendTime", DateTime.Now.ToString(CultureInfo.InvariantCulture));
                        key.SetValue("UsingEyeTimeSpan", UsingEyeTimeSpan.TotalSeconds.ToString(CultureInfo.InvariantCulture));
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Registry error: {ex.Message}");
                }
            }
            else if( status == "RestingEyes" )
            {
                // Do nothing, continue resting
            }
            else if( status == "Idle" )
            {
                // Do nothing
            }
            // System is entering hibernation/sleep
            DateTime now = DateTime.Now;
            WriteLog("System suspending (hibernation/sleep)");
        }
        else if (e.Mode == PowerModes.Resume)
        {
            // System is resuming from hibernation/sleep
            WriteLog("System resuming from hibernation/sleep");
            using (RegistryKey? key = Registry.CurrentUser.OpenSubKey(@"Software\ForcedRest"))
            {
                if (key != null)
                {
                    object? value = key.GetValue("UsingEyeTimeSpan");
                    if (value != null)
                    {
                        string usingEyeTimeSpanStr = value.ToString()!;
                        double usingEyeTimeSpanSeconds = double.Parse(usingEyeTimeSpanStr, CultureInfo.InvariantCulture);
                        UsingEyeTimeSpan = TimeSpan.FromSeconds(usingEyeTimeSpanSeconds);
                    }
                }
            }
            //Take the rest time into account when resume
            //There are 2 variables:
            //1. The EyeUsingTime, which could be 0 to 35.
            //2. The HibernationDuration, which could be any value.
            //3. The RestTime, it is determined by the NumberOfExtensions.
            //The rules:
            //1. if the HibernationDuration is greater than the RestTime, then directly reset the countdown.
            //2. if there are remaining, restore it accordingly.
        }
    }

    // Add this method to the Program class:
    private static void SystemEvents_SessionEnding(object sender, SessionEndingEventArgs e)
    {
        if (e.Reason == SessionEndReasons.SystemShutdown || 
        e.Reason == SessionEndReasons.Logoff )
        {
            UsingEyeTimeSpan = DateTime.Now - StartOfUsingEyes;
            EnterRestMode();            
            // The system is shutting down or rebooting.
            // Perform cleanup or save state here.
            try
            {
                using (RegistryKey key = Registry.CurrentUser.CreateSubKey(@"Software\ForcedRest"))
                {
                    key.SetValue("EndOfRestingEyes", EndOfRestingEyes.ToString(CultureInfo.InvariantCulture));
                    key.SetValue("EndOfUsingEyes", EndOfUsingEyes.ToString(CultureInfo.InvariantCulture));
                    WriteLog($"Shutdown or Logout: Saved EndOfRestingEyes to registry: {EndOfRestingEyes}");
                    key.SetValue("LogoutTime", DateTime.Now.ToString(CultureInfo.InvariantCulture));
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Registry error: {ex.Message}");
            }
            Console.WriteLine("System is shutting down or rebooting.");
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
        NumberOfExtensions = 0;
    }

    private static void EnterRestMode()
    {
        status = "RestingEyes";
        DateTime now = DateTime.Now;
        double usedMinutes = UsingEyeTimeSpan.TotalMinutes;     //bug: here may round to 0 if less than 0.5 minutes
        double expectedRestMinutes = (usedMinutes / UsingEyeMinutes) * restMinutes;
        StartOfRestingEyes = now;
        EndOfRestingEyes = now.AddMinutes(expectedRestMinutes) + TimeSpan.FromSeconds(NumberOfExtensions * 20);
        Console.WriteLine(now.ToString() + " expectedRestMinutes: " + expectedRestMinutes.ToString() + "    EndOfRestingEyes: " + EndOfRestingEyes.ToString());
    }
 
}

public class exceptionTime
{
    public int DayOfWeek { get; set; }
    [Name("Start")]
    public string? Start { get; set; }
    [Name("End")]
    public string? End { get; set; }

    public DateTime startTime;
    public DateTime endTime;
    public DateTime startDateTime;
    public DateTime endDateTime;
}

//publish command:
//dotnet publish -c Release -r win-x64 --self-contained true -p:PublishSingleFile=true -p:IncludeAllContentForSelfExtract=true
//debug run command:
//dotnet run -c Debug