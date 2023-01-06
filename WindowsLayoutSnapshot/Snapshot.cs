using System;
using System.Collections.Generic;
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
            const int textLength = 256;
            var outText = new System.Text.StringBuilder(textLength + 1);
            GetWindowText(hwnd, outText, outText.Capacity);
            Debug.WriteLine(hwnd + " " + outText);
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

        internal void Restore(object sender, EventArgs e) // ignore extra params
        {
            // first, restore the window rectangles and normal/maximized/minimized states
            foreach (var placement in m_infos)
            {
                var processId = placement.Key;

                try
                {
                    GetWindowThreadProcessId((IntPtr)processId, out var pid);
                    var proc = Process.GetProcessById((int)pid);
                    proc.MainModule.FileName.ToString();
                }
                catch
                {
                    // Skip process if PID does not exist
                    continue;
                }

                if (!SetWindowPlacement((IntPtr)processId, ref placement.Value.windowPlacement))
                {
                    var err = GetLastError();
                    throw new Exception($"SetWindowPlacement of processId {processId} with title '{placement.Key}' failed with error: {err}");
                }
            }

            // now update the z-orders
            m_windowsBackToTop = m_windowsBackToTop.FindAll(IsWindowVisible);
            var positionStructure = BeginDeferWindowPos(m_windowsBackToTop.Count);
            for (var i = 0; i < m_windowsBackToTop.Count; i++)
            {
                positionStructure = DeferWindowPos(positionStructure, m_windowsBackToTop[i],
                    i == 0 ? IntPtr.Zero : m_windowsBackToTop[i - 1], 0, 0, 0, 0,
                    DeferWindowPosCommands.SWP_NOMOVE | DeferWindowPosCommands.SWP_NOSIZE | DeferWindowPosCommands.SWP_NOACTIVATE);
            }
            EndDeferWindowPos(positionStructure);
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
            public string title;
            public WINDOWPLACEMENT windowPlacement;
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
            var win = new WinInfo();

            if (GetWindowPlacement(hwnd, out var windowPlacement))
            {
                win.windowPlacement = windowPlacement;
            }
            else
            {
                var err = GetLastError();
                throw new Exception($"GetWindowPlacement of hwnd {hwnd} failed with error: {err}");
            }
            
            // Get process title
            const int textLength = 256;
            var outText = new System.Text.StringBuilder(textLength + 1);
            GetWindowText(hwnd, outText, outText.Capacity);
            win.title = outText.ToString();
            
            return win;
        }

        public string GetJson() 
        {
            var snapshotJson = new SnapshotJSON
            {
                name = GetDisplayString(),
                processList = m_infos
            };
            return JsonConvert.SerializeObject(snapshotJson);
        }
    }
}
