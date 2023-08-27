using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text.Json;
using System.Text.Json.Serialization;
using System.Windows.Forms;
namespace WindowsLayoutSnapshot
{

    internal class Snapshot
    {
        private readonly List<Window> windows = new List<Window>();
        private DateTime timeTaken;

        internal bool UserInitiated { get; }
        internal string SnapshotName { get; }
        internal long[] MonitorPixelCounts { get; }
        internal int NumMonitors { get; }

        internal static Snapshot TakeSnapshot(bool userInitiated)
        {
            return new Snapshot(userInitiated, null);
        }

        internal static Snapshot TakeSnapshot(bool userInitiated, string snapshotName)
        {
            return new Snapshot(userInitiated, snapshotName);
        }

        internal static Snapshot TakeSnapshot(bool userInitiated, string snapshotName, List<WinInfo> windows)
        {
            return new Snapshot(userInitiated, snapshotName, windows);
        }

        private Snapshot(bool userInitiated, string snapshotName, List<WinInfo> windows = null)
        {
            UserInitiated = userInitiated;
            SnapshotName = snapshotName;

            if (windows == null)
            {
                Native.EnumWindows(EvalWindow, 0);
            }
            else
            {
                foreach (var window in windows)
                {
                    windows.Add(window);
                }
            }

            timeTaken = DateTime.UtcNow;

            var pixels = Screen.AllScreens.Select(s => (long)s.Bounds.Width * s.Bounds.Height).ToList();
            MonitorPixelCounts = pixels.ToArray();
            NumMonitors = pixels.Count;
        }

        private bool EvalWindow(IntPtr hwnd, IntPtr lParam)
        {
            var processName = GetProcessName(hwnd);
            var windowPlacement = GetWindowPlacement(hwnd);
            var newWindow = new Window(hwnd, processName, windowPlacement);

            if (newWindow.IsRestorableWindow())
                windows.Add(newWindow);

            return true;
        }

        private static string GetProcessName(IntPtr hwnd)
        {
            Native.GetWindowThreadProcessId(hwnd, out var processId);

            var handleNotValid = processId == 0;
            if (handleNotValid)
                return string.Empty;

            var processes = Process.GetProcesses();
            if (!processes.Any(p => p.Id == processId))
                return string.Empty;

            var process = processes.Single(p => p.Id == processId);
            return process.ProcessName;
        }

        private Native.WINDOWPLACEMENT GetWindowPlacement(IntPtr hwnd)
        {
            if (Native.GetWindowPlacement(hwnd, out var windowPlacement))
                return windowPlacement;

            var ex = new Win32Exception(Marshal.GetLastWin32Error());
            throw new Exception($"{nameof(GetWindowPlacement)} of hwnd {hwnd} failed with error: {ex.NativeErrorCode} ({ex.Message})");
        }

        internal void RestoreAndPreserveMenu(object sender, EventArgs e)
        {
            // We save and restore the current foreground window because it's our tray menu
            // I couldn't find a way to get this handle straight from the tray menu's properties;
            //   the ContextMenuStrip.Handle isn't the right one, so I'm using win32
            // More info RE the restore is below, where we do it
            var currentForegroundWindow = Native.GetForegroundWindow();

            try
            {
                Restore(sender, e);
            }
            finally
            {
                // A combination of SetForegroundWindow + SetWindowPos (via set_Visible) seems to be needed
                // This was determined by trying a bunch of stuff
                // This prevents the tray menu from closing, and makes sure it's still on top
                Native.SetForegroundWindow(currentForegroundWindow);
                TrayIconForm.me.Visible = true;
            }
        }

        internal void Restore(object sender, EventArgs e)
        {
            var processes = Process.GetProcesses();

            // first restore the window rectangles and normal/maximized/minimized states
            foreach (var window in windows)
            {
                Native.GetWindowThreadProcessId(window.Handle, out var processId);

                var handleNotValid = processId == 0;
                if (handleNotValid || !processes.Any(p => p.Id == processId))
                    continue;

                var wp = window.WindowPlacement;
                if (!Native.SetWindowPlacement(window.Handle, ref wp))
                {
                    var ex = new Win32Exception(Marshal.GetLastWin32Error());
                    throw new Exception($"{nameof(Native.SetWindowPlacement)} of window with " +
                        $"hwnd {window.Handle} and process name '{window.ProcessName}' " +
                        $"failed with error: {ex.NativeErrorCode} ({ex.Message})");
                }
            }

            // then update the z-orders
            var visibleWindows = windows.Where(w => Native.IsWindowVisible(w.Handle)).ToList();
            var positionHandle = Native.BeginDeferWindowPos(visibleWindows.Count);
            var previousWindowHandle = IntPtr.Zero;
            foreach (var window in visibleWindows)
            {
                positionHandle = Native.DeferWindowPos(positionHandle, window.Handle, previousWindowHandle, 0, 0, 0, 0,
                    Native.DeferWindowPosCommands.SWP_NOMOVE | Native.DeferWindowPosCommands.SWP_NOSIZE |
                    Native.DeferWindowPosCommands.SWP_NOACTIVATE);
                previousWindowHandle = window.Handle;
            }
            Native.EndDeferWindowPos(positionHandle);
        }

        [Serializable]
        public class SnapshotJson
        {
            [JsonPropertyName("name")]
            public string Name { get; set; }
            [JsonPropertyName("windows")]
            public List<WinInfo> Windows { get; set; }
        }

        [Serializable]
        public class WinInfo
        {
            [JsonPropertyName("windowHandle")]
            public long WindowHandle { get; set; }
            [JsonPropertyName("processName")]
            public string ProcessName { get; set; }
            [JsonPropertyName("windowPlacement")]
            public Native.WINDOWPLACEMENT WindowPlacement { get; set; }
        }

        public string GetJson()
        {
            var snapshotJson = new SnapshotJson
            {
                Name = GetDisplayString(),
                Windows = windows.Select(w => new WinInfo
                {
                    WindowHandle = w.Handle.ToInt64(),
                    ProcessName = w.ProcessName,
                    WindowPlacement = w.WindowPlacement
                }).ToList()
            };
            return JsonSerializer.Serialize(snapshotJson);
        }

        public string GetDisplayString() => SnapshotName ?? timeTaken.ToLocalTime().ToString("d MMMM, HH:mm:ss");
    }
}
