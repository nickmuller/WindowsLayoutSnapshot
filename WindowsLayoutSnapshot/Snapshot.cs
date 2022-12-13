using System;
using System.Collections.Generic;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using System.Diagnostics;
using static WindowsLayoutSnapshot.Native;
using Newtonsoft.Json;

namespace WindowsLayoutSnapshot
{

    internal class Snapshot {

        private Dictionary<int, WinInfo> m_infos = new Dictionary<int, WinInfo>();
        private List<IntPtr> m_windowsBackToTop = new List<IntPtr>();

        private Snapshot(bool userInitiated) {
#if DEBUG
            Debug.WriteLine("*** NEW SNAPSHOT ***");
#endif
            EnumWindows(EvalWindow, 0);

            TimeTaken = DateTime.UtcNow;
            UserInitiated = userInitiated;

            var pixels = new List<long>();
            foreach (var screen in Screen.AllScreens)
                pixels.Add(screen.Bounds.Width * screen.Bounds.Height);
            MonitorPixelCounts = pixels.ToArray();
            NumMonitors = pixels.Count;
        }

        private Snapshot(bool userInitiated, string snapshotName)
        {
#if DEBUG
            Debug.WriteLine("*** NEW SNAPSHOT ***");
#endif

            this.snapshotName = snapshotName;

            EnumWindows(EvalWindow, 0);

            TimeTaken = DateTime.UtcNow;
            UserInitiated = userInitiated;

            var pixels = new List<long>();
            foreach (var screen in Screen.AllScreens)
                pixels.Add(screen.Bounds.Width * screen.Bounds.Height);
            MonitorPixelCounts = pixels.ToArray();
            NumMonitors = pixels.Count;
        }

        private Snapshot(bool userInitiated, string snapshotName, Dictionary<int, WinInfo> processList)
        {
#if DEBUG
            Debug.WriteLine("*** RESTORE SNAPSHOT ***");
#endif

            this.snapshotName = snapshotName;

            //EnumWindows(EvalWindow, 0);

            //m_infos = processList;
            foreach (var item in processList)
            {
                m_infos.Add(item.Key, item.Value);
            }

            TimeTaken = DateTime.UtcNow;
            UserInitiated = userInitiated;

            var pixels = new List<long>();
            foreach (var screen in Screen.AllScreens)
                pixels.Add(screen.Bounds.Width * screen.Bounds.Height);
            MonitorPixelCounts = pixels.ToArray();
            NumMonitors = pixels.Count;
        }

        internal static Snapshot TakeSnapshot(bool userInitiated) {
            return new Snapshot(userInitiated);
        }

        internal static Snapshot TakeSnapshot(bool userInitiated, string snapshotName)
        {
            return new Snapshot(userInitiated, snapshotName);
        }

        internal static Snapshot TakeSnapshot(bool userInitiated, string snapshotName, Dictionary<int, WinInfo>  processList)
        {
            return new Snapshot(userInitiated, snapshotName, processList);
        }

        private bool EvalWindow(int hwndInt, int lParam) {
            var hwnd = new IntPtr(hwndInt);

            if (!IsAltTabWindow(hwnd))
                return true;

            // EnumWindows returns windows in Z order from back to front
            m_windowsBackToTop.Add(hwnd);

            WinInfo win = GetWindowInfo(hwnd);
            m_infos.Add((int)hwnd.ToInt64(), win);

#if DEBUG
            // For debugging purpose, output window title with handle
            int textLength = 256;
            System.Text.StringBuilder outText = new System.Text.StringBuilder(textLength + 1);
            int a = GetWindowText(hwnd, outText, outText.Capacity);
            Debug.WriteLine(hwnd + " " + win.position + " " + outText);
#endif

            return true;
        }

        internal DateTime TimeTaken { get; private set; }
        internal bool UserInitiated { get; private set; }
        internal string snapshotName { get; private set; }
        internal long[] MonitorPixelCounts { get; private set; }
        internal int NumMonitors { get; private set; }

        public string GetDisplayString() {
            DateTime dt = TimeTaken.ToLocalTime();
            return snapshotName != null ? snapshotName : dt.ToString("M") + ", " + dt.ToString("T");
        }

        internal TimeSpan Age {
            get { return DateTime.UtcNow.Subtract(TimeTaken); }
        }

        internal void RestoreAndPreserveMenu(object sender, EventArgs e) { // ignore extra params
            // We save and restore the current foreground window because it's our tray menu
            // I couldn't find a way to get this handle straight from the tray menu's properties;
            //   the ContextMenuStrip.Handle isn't the right one, so I'm using win32
            // More info RE the restore is below, where we do it
            var currentForegroundWindow = GetForegroundWindow();

            try {
                Restore(sender, e);
            } finally {
                // A combination of SetForegroundWindow + SetWindowPos (via set_Visible) seems to be needed
                // This was determined by trying a bunch of stuff
                // This prevents the tray menu from closing, and makes sure it's still on top
                SetForegroundWindow(currentForegroundWindow);
                TrayIconForm.me.Visible = true;
            }
        }

        internal void Restore(object sender, EventArgs e) { // ignore extra params
                                                            // first, restore the window rectangles and normal/maximized/minimized states
            foreach (var placement in m_infos) {

                var processId = placement.Key;
                var processPath = placement.Value.processPath;

                // If PID is not existing try to start process
                try
                {
                    uint pid = 0;
                    GetWindowThreadProcessId((IntPtr)processId, out pid);
                    Process proc = Process.GetProcessById((int)pid);
                    proc.MainModule.FileName.ToString();
                }
                catch
                {
                    try
                    {
                        var processFound = false;

                        // Check if process path is already started
                        foreach (var item in Process.GetProcesses())
                        {
                           try
                            {
                                if (item.MainModule.FileName.ToString() == processPath && processFound == false)
                                {
                                    Debug.WriteLine("Existing process");
                                    int count = 0;
                                    while (count < 6)
                                    {
                                        var process = item;
                                        process.Refresh();
                                        processId = (int)process.MainWindowHandle.ToInt64();
                                        if(processId == 0)
                                        {
                                            System.Threading.Thread.Sleep(250);
                                            Debug.WriteLine("Count =" + count);
                                            count++;
                                        }else
                                        {
                                            processFound = true;
                                            count = 6;
                                        }
                                    }
                                    if (processFound)
                                    {
                                        break;
                                    }
                                }
                            }
                            catch
                            {
                                // Don't trigger an error if access denied
                            }
                        }

                        if (processFound == false)
                        {
                            Debug.WriteLine("New process");
                            // Start process because app can't be started
                            var process = Process.Start(processPath);
                            int count = 0;
                            while (count < 12)
                            {
                                process.Refresh();
                                processId = (int)process.MainWindowHandle.ToInt64();
                                if (processId == 0)
                                {
                                    System.Threading.Thread.Sleep(250);
                                    Debug.WriteLine("New Count =" + count);
                                    count++;
                                }
                                else
                                {
                                    processFound = true;
                                    count = 12;
                                }
                            }
                        }
                    }
                    catch (Exception errorToCheck)
                    {
                        Debug.WriteLine(errorToCheck);
                        // Don't trigger an error if app is not existing anymore
                    }
                }


                // make sure window will be inside a monitor
                Rectangle newpos = GetRectInsideNearestMonitor(placement.Value);
                Debug.WriteLine(newpos);
                try
                {
                    if (!SetWindowPos((IntPtr)processId, 0, newpos.Left, newpos.Top, newpos.Width, newpos.Height, 0x0004 /*NOZORDER*/))
                    {
                        var err = GetLastError();
                        Debug.WriteLine("Can't move window " + placement.Key + ": Error" + GetLastError());
                    }else
                    {
                        Debug.WriteLine(processId);
                    }
                }
                catch(Exception errrr)
                {
                    Debug.WriteLine(placement.Key);
                    Debug.WriteLine(errrr);
                }
            }


            // now update the z-orders
            m_windowsBackToTop = m_windowsBackToTop.FindAll(IsWindowVisible);
            IntPtr positionStructure = BeginDeferWindowPos(m_windowsBackToTop.Count);
            for (int i = 0; i < m_windowsBackToTop.Count; i++) {
                positionStructure = DeferWindowPos(positionStructure, m_windowsBackToTop[i], i == 0 ? IntPtr.Zero : m_windowsBackToTop[i - 1],
                    0, 0, 0, 0, DeferWindowPosCommands.SWP_NOMOVE | DeferWindowPosCommands.SWP_NOSIZE | DeferWindowPosCommands.SWP_NOACTIVATE);
            }
            EndDeferWindowPos(positionStructure);
        }

        private static Rectangle GetRectInsideNearestMonitor(WinInfo win) {
            Rectangle real = win.position;
            Rectangle rect = win.visible;
            Rectangle monitorRect = Screen.GetWorkingArea(rect); // use workspace coordinates
            Rectangle y = new Rectangle(
                Math.Max(monitorRect.Left, Math.Min(monitorRect.Right - rect.Width, rect.Left)),
                Math.Max(monitorRect.Top, Math.Min(monitorRect.Bottom - rect.Height, rect.Top)),
                Math.Min(monitorRect.Width, rect.Width),
                Math.Min(monitorRect.Height, rect.Height)
            );
            if (rect != real) // support different real and visible position
                y = new Rectangle(
                    y.Left - rect.Left + real.Left,
                    y.Top - rect.Top + real.Top,
                    y.Width - rect.Width + real.Width,
                    y.Height - rect.Height + real.Height
                );
#if DEBUG
            if (y != real)
                Debug.WriteLine("Moving " + real + "→" + y + " in monitor " + monitorRect);
#endif
            return y;
        }

        private static bool IsAltTabWindow(IntPtr hwnd) {
            if (!IsWindowVisible(hwnd))
                return false;

            IntPtr extendedStyles = GetWindowLongPtr(hwnd, (-20)); // GWL_EXSTYLE
            if ((extendedStyles.ToInt64() & WS_EX_APPWINDOW) > 0)
                return true;
            if ((extendedStyles.ToInt64() & WS_EX_TOOLWINDOW) > 0)
                return false;

            IntPtr hwndTry = GetAncestor(hwnd, GetAncestor_Flags.GetRootOwner);
            IntPtr hwndWalk = IntPtr.Zero;
            while (hwndTry != hwndWalk) {
                hwndWalk = hwndTry;
                hwndTry = GetLastActivePopup(hwndWalk);
                if (IsWindowVisible(hwndTry))
                    break;
            }
            if (hwndWalk != hwnd)
                return false;

            return true;
        }

        public class WinInfo {
            public Rectangle position; // real window border, we use this to move it
            public Rectangle visible; // visible window borders, we use this to force inside a screen
            public string title;
            public string processPath;
        }

        public class SnapshotJSON
        {
            public string name; 
            public Dictionary<int, WinInfo> processList;
        }

        public class SnapshotBackJSON
        {
            public string name;
            public Dictionary<int, WinInfo> processList;
        }

        private static WinInfo GetWindowInfo(IntPtr hwnd) {
            WinInfo win = new WinInfo();
            RECT pos;
            if (!GetWindowRect(hwnd, out pos))
                throw new Exception("Error getting window rectangle");
            win.position = win.visible = pos.ToRectangle();
            if (Environment.OSVersion.Version.Major >= 6)
                if (DwmGetWindowAttribute(hwnd, 9 /*DwmwaExtendedFrameBounds*/, out pos, Marshal.SizeOf(typeof(Native.RECT))) == 0)
                    win.visible = pos.ToRectangle();

            // Get process title
            int textLength = 256;
            System.Text.StringBuilder outText = new System.Text.StringBuilder(textLength + 1);
            int a = GetWindowText(hwnd, outText, outText.Capacity);
            win.title = outText.ToString();

            // Get process path
            try
            {
                uint pid = 0;
                GetWindowThreadProcessId(hwnd, out pid);
                Process proc = Process.GetProcessById((int)pid); //Gets the process by ID.
                win.processPath = proc.MainModule.FileName.ToString();    //Returns the path. 
            }
            catch
            {
                win.processPath = null;
            }

            return win;
        }

        public string getJSON() 
        {
            string jsonString = JsonConvert.SerializeObject(this.m_infos);

            SnapshotJSON snapshotJSON = new SnapshotJSON();
            snapshotJSON.name = this.GetDisplayString();
            snapshotJSON.processList = this.m_infos;

            return JsonConvert.SerializeObject(snapshotJSON);
        }
    }
}
