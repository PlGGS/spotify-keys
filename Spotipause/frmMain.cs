using System;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Gma.System.MouseKeyHook;

namespace Spotipause
{
    public partial class frmMain : Form
    {
        [DllImport("user32.dll")]
        private static extern IntPtr SendMessage(IntPtr hWnd, int Msg, IntPtr wParam, IntPtr lParam);
        [DllImport("user32.dll", SetLastError = true)]
        internal static extern bool MoveWindow(IntPtr hWnd, int X, int Y, int nWidth, int nHeight, bool bRepaint);
        [DllImport("user32.dll")]
        public static extern int SetForegroundWindow(IntPtr hwnd);
        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        static extern bool ShowWindow(IntPtr hWnd, ShowWindowEnum flags);
        private enum ShowWindowEnum
        {
            Hide = 0,
            ShowNormal = 1, ShowMinimized = 2, ShowMaximized = 3,
            Maximize = 3, ShowNormalNoActivate = 4, Show = 5,
            Minimize = 6, ShowMinNoActivate = 7, ShowNoActivate = 8,
            Restore = 9, ShowDefault = 10, ForceMinimized = 11
        };
        private IKeyboardMouseEvents m_GlobalHook;
        Process spotifyProcess;
        IntPtr spotifyHWnd;
        private const int WM_APPCOMMAND = 0x319;
        private const int MEDIA_PREVIOUS_TRACK = 0xC0000;
        private const int MEDIA_NEXT_TRACK = 0xB0000;
        bool pressingControl = false;
        bool pressingShift = false;
        int processEntry = 0;

        /// <summary>
        /// Hides application from alt+tab menu
        /// </summary>
        protected override CreateParams CreateParams
        {
            get
            {
                CreateParams cp = base.CreateParams;
                cp.ExStyle |= 0x80;
                return cp;
            }
        }

        /// <summary>
        /// Subscribes to MouseKeyHook global events
        /// </summary>
        public void Subscribe()
        {
            // Note: for the application hook, use the Hook.AppEvents() instead
            m_GlobalHook = Hook.GlobalEvents();
            m_GlobalHook.KeyDown += GlobalHookKeyDown;
            m_GlobalHook.KeyUp += GlobalHookKeyUp;
        }

        /// <summary>
        /// Detects KeyUp events while outside of application
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void GlobalHookKeyUp(object sender, KeyEventArgs e)
        {
            if (pressingControl)
            {
                if (pressingShift)
                {
                    if (e.KeyCode == Keys.Space)
                    {
                        //Last song
                        Console.WriteLine("Skipping back");
                        SendMessage(spotifyHWnd, WM_APPCOMMAND, (IntPtr)0, (IntPtr)MEDIA_PREVIOUS_TRACK);
                    }
                }
                else if (e.KeyCode == Keys.Space)
                {
                    //Next song
                    Console.WriteLine("Skipping forward");
                    SendMessage(spotifyHWnd, WM_APPCOMMAND, (IntPtr)0, (IntPtr)MEDIA_NEXT_TRACK);
                }
            }

            if (e.KeyCode == Keys.LShiftKey)
            {
                pressingShift = false;
            }

            if (e.KeyCode == Keys.LControlKey)
            {
                pressingControl = false;
            }
            
            Console.WriteLine("shift: " + pressingShift + " | control: " + pressingControl);
        }

        /// <summary>
        /// Detects KeyDown events while outside of application
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void GlobalHookKeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.LControlKey)
            {
                pressingControl = true;
            }

            if (e.KeyCode == Keys.LShiftKey)
            {
                pressingShift = true;
            }

            Console.WriteLine("shift: " + pressingShift + " | control: " + pressingControl);
        }

        /// <summary>
        /// Unsubscribes from MouseKeyHook events
        /// </summary>
        public void Unsubscribe()
        {
            m_GlobalHook.KeyDown -= GlobalHookKeyDown;
            m_GlobalHook.KeyUp -= GlobalHookKeyUp;
            m_GlobalHook.Dispose();
        }

        /// <summary>
        /// Application Initialization
        /// </summary>
        public frmMain()
        {
            InitializeComponent();
            WriteProcesses("Spotify");
            GetMainSpotifyProcess();
            Console.WriteLine("used: " + processEntry);
            Subscribe();
        }

        /// <summary>
        /// Finds the Spotfiy process that has a usable hWnd to hook into
        /// </summary>
        private void GetMainSpotifyProcess()
        {
            if (Process.GetProcessesByName("Spotify") != null)
            {
                spotifyProcess = Process.GetProcessesByName("Spotify")[processEntry];

                for (int i = processEntry; i < Process.GetProcessesByName("Spotify").Length; i++)
                {
                    try
                    {
                        spotifyHWnd = spotifyProcess.MainWindowHandle;
                    }
                    catch (Exception)
                    {
                        Console.WriteLine("caught: " + processEntry);
                        processEntry++;
                        GetMainSpotifyProcess();
                    }
                }
            }
            else
            {
                MessageBox.Show("Please start Spotify before starting Spotipause. Thank you.", "Spotipause");
                Application.Exit();
            }
        }

        /// <summary>
        /// Writes running process names that contain a specific set of characters to the console for easy viewing
        /// </summary>
        private void WriteProcesses(string procNameContains)
        {
            Process[] processes = Process.GetProcesses();
            foreach (Process proc in processes)
            {
                if (proc.ProcessName.Contains("procNameContains"))
                {
                    Console.WriteLine(proc.ProcessName + " | " + proc.Id);
                }
            }
        }

        /// <summary>
        /// Brings the Spotify application window into view
        /// </summary>
        public void BringWindowToFront()
        {
            if (spotifyProcess != null)
            {
                if (spotifyHWnd == IntPtr.Zero)
                {
                    ShowWindow(spotifyProcess.Handle, ShowWindowEnum.Restore);
                }

                SetForegroundWindow((IntPtr)spotifyProcess.MainWindowHandle);
            }
            else
            {
                Process.Start(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + "\\Spotify\\Spotify.exe");
                Application.Restart();
            }
        }

        /// <summary>
        /// Run on the initialization of the form
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void frmMain_Load(object sender, EventArgs e)
        {
            this.Visible = false;
        }
    }
}
