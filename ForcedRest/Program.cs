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
    static TimeSpan UsingEyeTimeSpan = TimeSpan.FromMinutes(0); //placeholder
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

    //static double remaining = UsingEyeMinutes;
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
            Console.WriteLine("Starting app");
            var version = System.Reflection.Assembly.GetExecutingAssembly().GetName().Version;
            Console.WriteLine($"Version: {version}");
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
                        DateTime now = DateTime.Now;
                        if( now < EndOfRestingEyes )
                        {
                            status = "RestingEyes";
                            Console.WriteLine($"Loaded EndOfRestingEyes from registry: {EndOfRestingEyes}");
                        }
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
            ResetCountDown();   
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
                    double remaining= (EndOfUsingEyes - now).TotalMinutes;  
                    //label.ForeColor = Color.White;
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
                        label.Text = $"{ts.Seconds:D2}";

                    }
                    else{
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
                CalculateRestTime();
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
                CalculateRestTime();
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
            Console.WriteLine(now.ToString());
            Console.WriteLine("System suspending (hibernation/sleep)");
        }
        else if (e.Mode == PowerModes.Resume)
        {
            // System is resuming from hibernation/sleep
            DateTime now = DateTime.Now;
            Console.WriteLine(now.ToString());
            Console.WriteLine("System resuming from hibernation/sleep");
            // Optionally reset or adjust timers if needed
            //ResetCountDown();
        }
    }

    // Add this method to the Program class:
    private static void SystemEvents_SessionEnding(object sender, SessionEndingEventArgs e)
    {
        if (e.Reason == SessionEndReasons.SystemShutdown || 
        e.Reason == SessionEndReasons.Logoff )
        {
            // The system is shutting down or rebooting.
            // Perform cleanup or save state here.
            try
            {
                using (RegistryKey key = Registry.CurrentUser.CreateSubKey(@"Software\ForcedRest"))
                {
                    key.SetValue("EndOfRestingEyes", EndOfRestingEyes.ToString(CultureInfo.InvariantCulture));
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
    }

    private static void CalculateRestTime()
    {
        status = "RestingEyes";
        DateTime now = DateTime.Now;
        double usedMinutes = UsingEyeTimeSpan.TotalMinutes;     //bug: here may round to 0 if less than 0.5 minutes
        double expectedRestMinutes = (usedMinutes / UsingEyeMinutes) * restMinutes;
        StartOfRestingEyes = now;
        EndOfRestingEyes = now.AddMinutes(expectedRestMinutes);
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