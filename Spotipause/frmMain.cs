using System;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Gma.System.MouseKeyHook;
using System.IO;
using System.Reflection;
using System.Threading;

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
        [DllImport("user32.dll")]
        private static extern bool RegisterHotKey(IntPtr hWnd, int id, uint fsModifiers, uint vk);
        [DllImport("user32.dll")]
        private static extern bool UnregisterHotKey(IntPtr hWnd, int id);
        private enum ShowWindowEnum
        {
            Hide = 0,
            ShowNormal = 1, ShowMinimized = 2, ShowMaximized = 3,
            Maximize = 3, ShowNormalNoActivate = 4, Show = 5,
            Minimize = 6, ShowMinNoActivate = 7, ShowNoActivate = 8,
            Restore = 9, ShowDefault = 10, ForceMinimized = 11
        };
        private IKeyboardMouseEvents m_GlobalHook;
        Process[] spotifyProcesses;
        IntPtr[] spotifyHWnds;
        const int WM_APPCOMMAND = 0x319;
        const int MEDIA_PREVIOUS_TRACK = 0xC0000;
        const int MEDIA_NEXT_TRACK = 0xB0000;
        const int APPCOMMAND_PLAY_PAUSE = 0xE0000;
        bool pressingControl = false;
        bool pressingShift = false;
        bool pressingAlt = false;

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
#if DEBUG
            Console.ForegroundColor = ConsoleColor.Blue;
#endif

            if (pressingControl)
            {
                if (pressingAlt)
                {
                    if (e.KeyCode == Keys.Space)
                    {
                        //Play and pause
#if DEBUG
                        Console.WriteLine("Playing/Pausing");
#endif
                        for (int i = 0; i < Process.GetProcessesByName("Spotify").Length; i++)
                        {
                            SendMessage(spotifyHWnds[i], WM_APPCOMMAND, (IntPtr)0, (IntPtr)APPCOMMAND_PLAY_PAUSE);
                        }
                    }
                }
                else if (pressingShift)
                {
                    if (e.KeyCode == Keys.Space)
                    {
                        //Last song
#if DEBUG
                        Console.WriteLine("Skipping back");
#endif
                        for (int i = 0; i < Process.GetProcessesByName("Spotify").Length; i++)
                        {
                            SendMessage(spotifyHWnds[i], WM_APPCOMMAND, (IntPtr)0, (IntPtr)MEDIA_PREVIOUS_TRACK);
                        }
                    }
                }
                else if (e.KeyCode == Keys.Space)
                {
                    //Next song
#if DEBUG
                    Console.WriteLine("Skipping forward");
#endif
                    for (int i = 0; i < Process.GetProcessesByName("Spotify").Length; i++)
                    {
                        SendMessage(spotifyHWnds[i], WM_APPCOMMAND, (IntPtr)0, (IntPtr)MEDIA_NEXT_TRACK);
                    }
                }
            }

            switch (e.KeyCode)
            {
                case Keys.LControlKey:
                    pressingControl = false;
                    break;
                case Keys.LShiftKey:
                    pressingShift = false;
                    break;
                case Keys.LMenu:
                    pressingAlt = false;
                    break;
            }

#if DEBUG
            Console.ForegroundColor = ConsoleColor.Green;
            Console.WriteLine($"UP ({e.KeyCode.ToString()}) :: shift: " + pressingShift + " | control: " + pressingControl + " | alt: " + pressingAlt);
#endif
        }

        /// <summary>
        /// Detects KeyDown events while outside of application
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void GlobalHookKeyDown(object sender, KeyEventArgs e)
        {
            switch (e.KeyCode)
            {
                case Keys.LControlKey:
                    pressingControl = true;
                    break;
                case Keys.LShiftKey:
                    pressingShift = true;
                    break;
                case Keys.LMenu:
                    pressingAlt = true;
                    break;
            }

#if DEBUG
            Console.ForegroundColor = ConsoleColor.Red;
            Console.WriteLine($"DOWN ({e.KeyCode.ToString()}) :: shift: " + pressingShift + " | control: " + pressingControl + " | alt: " + pressingAlt);
#endif
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
#if DEBUG
            WriteProcesses("Spotify");
#endif
            GetSpotifyProcesses();
        }

        /// <summary>
        /// Finds and stores all Spotfiy processes that have usable hWnds to hook into
        /// </summary>
        private void GetSpotifyProcesses()
        {
            if (Process.GetProcessesByName("Spotify").Length != 0)
            {
                spotifyProcesses = new Process[Process.GetProcessesByName("Spotify").Length];
                spotifyHWnds = new IntPtr[spotifyProcesses.Length];

                for (int i = 0; i < Process.GetProcessesByName("Spotify").Length; i++)
                {
                    try
                    {
                        spotifyProcesses[i] = Process.GetProcessesByName("Spotify")[i];
                        spotifyHWnds[i] = spotifyProcesses[i].MainWindowHandle;
                    }
                    catch (Exception)
                    {
#if DEBUG
                        Console.WriteLine("caught: " + Process.GetProcessesByName("Spotify")[i]);
#endif
                    }
                }

                Subscribe();
            }
            else
            {
#if DEBUG
                Console.WriteLine("Starting Spotify");
#endif

                try
                {
#if DEBUG
                    Debug.WriteLine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().GetName().CodeBase).Substring(6) + "\\Spotify.lnk");
#endif
                    Process.Start(Path.GetDirectoryName(Assembly.GetExecutingAssembly().GetName().CodeBase).Substring(6) + "\\Spotify.lnk");
                    Thread.Sleep(5000);
                    Application.Restart();
                }
                catch (Exception)
                {
                    MessageBox.Show("Failed to start the Spotify process. Please either place a shortcut to Spotify in your Spotipause directory, or start Spotify manually before running.", "Spotipause");
                    throw;
                }
            }
        }

        /// <summary>
        /// Writes running process names that contain a specific set of characters to the console for easy viewing
        /// </summary>
        private void WriteProcesses(string procName)
        {
            Process[] processes = Process.GetProcesses();
            foreach (Process proc in processes)
            {
                if (proc.ProcessName.Contains(procName))
                {
                    Console.WriteLine(proc.ProcessName + " | " + proc.Id);
                }
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
