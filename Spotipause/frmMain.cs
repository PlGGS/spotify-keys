using System;
using System.Diagnostics;
using System.Linq;
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

        /// <summary>
        /// Hides application from alt+tab menu
        /// </summary>
        protected override CreateParams CreateParams
        {
            get
            {
                // Turn on WS_EX_TOOLWINDOW style bit
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
            if (!ModifierKeys.HasFlag(Keys.Shift))
            {
                pressingShift = false;
            }

            if (!ModifierKeys.HasFlag(Keys.Control))
            {
                pressingControl = false;
            }

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

            //Console.WriteLine("shift: " + pressingShift + " | control: " + pressingControl);
        }

        /// <summary>
        /// Detects KeyDown events while outside of application
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void GlobalHookKeyDown(object sender, KeyEventArgs e)
        {
            if (ModifierKeys.HasFlag(Keys.Shift))
            {
                pressingShift = true;
            }

            if (ModifierKeys.HasFlag(Keys.Control))
            {
                pressingControl = true;
            }
        }
        
        /// <summary>
        /// Unsubscribes from MouseKeyHook events
        /// </summary>
        public void Unsubscribe()
        {
            m_GlobalHook.KeyDown -= GlobalHookKeyDown;
            m_GlobalHook.KeyUp -= GlobalHookKeyUp;

            //It is recommeneded to dispose it
            m_GlobalHook.Dispose();
        }

        /// <summary>
        /// Application Initialization
        /// </summary>
        public frmMain()
        {
            InitializeComponent();
            Subscribe();

            spotifyProcess = Process.GetProcessesByName("Spotify").FirstOrDefault();
            
            if (spotifyProcess != null)
            {
                //get the hWnd of the process
                spotifyHWnd = spotifyProcess.MainWindowHandle;
            }
            else
            {
                //Only works with defualt installation folder for now...
                Process.Start(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + "\\Spotify\\Spotify.exe");
                Application.Restart();
            }
        }

        /// <summary>
        /// Writes all running process names to the console for easy viewing
        /// </summary>
        private void WriteProcesses()
        {
            Process[] processes = Process.GetProcesses();
            foreach (Process proc in processes)
            {
                Console.WriteLine(proc.ProcessName);
            }
        }

        /// <summary>
        /// Brings the Spotify application window into view
        /// </summary>
        public void BringWindowToFront()
        {
            //WriteProcesses();

            if (spotifyProcess != null)
            {
                if (spotifyHWnd == IntPtr.Zero)
                {
                    //the window is hidden so try to restore it before setting focus.
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
