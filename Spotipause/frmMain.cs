using System;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Gma.System.MouseKeyHook;
using System.IO;
using System.Reflection;
using System.Threading;
using System.Data.SqlClient;
using System.Configuration;
using Windows.UI.Notifications;
//using Microsoft.Toolkit.Uwp.Notifications; // Notifications library
using Microsoft.QueryStringDotNET; // QueryString.NET
using Windows.Data.Xml.Dom;

namespace Spotipause
{
    public partial class frmMain : Form
    {
        // Sebastian
        [DllImport("user32.dll")]
        static extern IntPtr GetForegroundWindow();
        [DllImport("user32.dll")]
        private static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint ProcessId);

        // Blake
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
        const int WM_HOTKEY = 0x0312;
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

        private void Notification(String text)
        {

            // In a real app, these would be initialized with actual data
            string title = "Andrew sent you a picture";
            string content = "Check this out, Happy Canyon in Utah!";
            string image = "http://blogs.msdn.com/cfs-filesystemfile.ashx/__key/communityserver-blogs-components-weblogfiles/00-00-01-71-81-permanent/2727.happycanyon1_5B00_1_5D00_.jpg";
            string logo = "ms-appdata:///local/Andrew.jpg";

            // TODO: all values need to be XML escaped

            // Construct the visuals of the toast
            string toastVisual =
            $@"<visual>
  <binding template='ToastGeneric'>
    <text>{title}</text>
    <text>{content}</text>
    <image src='{image}'/>
    <image src='{logo}' placement='appLogoOverride' hint-crop='circle'/>
  </binding>
</visual>";

            // In a real app, these would be initialized with actual data
            int conversationId = 384928;

            // Generate the arguments we'll be passing in the toast
            string argsReply = $"action=reply&conversationId={conversationId}";
            string argsLike = $"action=like&conversationId={conversationId}";
            string argsView = $"action=viewImage&imageUrl={Uri.EscapeDataString(image)}";

            // TODO: all args need to be XML escaped

            string toastActions =
            $@"<actions>
 
  <input
      type='text'
      id='tbReply'
      placeHolderContent='Type a response'/>
 
  <action
      content='Reply'
      arguments='{argsReply}'
      activationType='background'
      imageUri='Assets/Reply.png'
      hint-inputId='tbReply'/>
 
  <action
      content='Like'
      arguments='{argsLike}'
      activationType='background'/>
 
  <action
      content='View'
      arguments='{argsView}'/>
 
</actions>";

            // Now we can construct the final toast content
            string argsLaunch = $"action=viewConversation&conversationId={conversationId}";

            // TODO: all args need to be XML escaped

            string toastXmlString =
            $@"<toast launch='{argsLaunch}'>
    {toastVisual}
    {toastActions}
</toast>";

            // Parse to XML
            XmlDocument toastXml = new XmlDocument();
            toastXml.LoadXml(toastXmlString);

            // Generate toast
            var toast = new ToastNotification(toastXml);
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

            #region SEBASTIAN

            var blacklisted = false;
            var activeWindowPath = "";

            IntPtr activeWindow = GetForegroundWindow();
            uint processId;
            GetWindowThreadProcessId(activeWindow, out processId);
            try
            {
                var activeProcess = Process.GetProcessById((int)processId);
#if DEBUG
                Debug.WriteLine(activeProcess.MainModule.FileName);
#endif
                var path = activeProcess.MainModule.FileName;
                activeWindowPath = path;

                SqlConnection connection = new SqlConnection(Properties.Settings.Default.blacklistConnectionString);
                SqlCommand command = new SqlCommand("select * from blacklist where path LIKE @path", connection);
                command.Parameters.Add("@path", System.Data.SqlDbType.NText).Value = path;
                connection.Open();
                var isBlocked = command.ExecuteReader().Read();
                blacklisted = isBlocked;
            }
            catch (Exception err)
            {
#if DEBUG
                Debug.WriteLine(err);
                Debug.WriteLine("couldn't find process");
#endif
            }

            var sql = "";

            if(pressingControl && pressingAlt && pressingShift && e.KeyCode == Keys.Insert && !blacklisted)
            {
                sql = "INSERT INTO blacklist (path) VALUES (@path);";
                this.Notification("s");
            }
            if(pressingControl && pressingAlt && pressingShift && e.KeyCode == Keys.Delete)
            {
#if DEBUG
                Debug.WriteLine("delete");
#endif
                sql = "DELETE FROM blacklist WHERE path LIKE @path;";
            }

            if (sql != "" && activeWindowPath != "")
            {
                try
                {
                    SqlConnection connection = new SqlConnection(Properties.Settings.Default.blacklistConnectionString);
                    SqlCommand command = new SqlCommand(sql, connection);
                    command.Parameters.Add("@path", System.Data.SqlDbType.NText).Value = activeWindowPath;
                    connection.Open();
                    command.ExecuteNonQuery();
                } catch(Exception) { }
            }

            if (blacklisted)
            {
                return;
            }

            #endregion

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

            if (e.KeyCode == Keys.LControlKey)
            {
                pressingControl = false;
            }

            if (e.KeyCode == Keys.LShiftKey)
            {
                pressingShift = false;
            }

            if (e.KeyCode == Keys.LMenu)
            {
                pressingAlt= false;
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
            if (e.KeyCode == Keys.LControlKey)
            {
                pressingControl = true;
            }

            if (e.KeyCode == Keys.LShiftKey)
            {
                pressingShift = true;
            }

            if (e.KeyCode == Keys.LMenu)
            {
                pressingAlt = true;
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
            WriteProcesses("Spotify");
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
                    MessageBox.Show("Failed to start the Spotify process. Please either fix the shortcut in your Spotipause directory, or start Spotify manually.", "Spotipause");
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
#if DEBUG
                    Console.WriteLine(proc.ProcessName + " | " + proc.Id);
#endif
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
