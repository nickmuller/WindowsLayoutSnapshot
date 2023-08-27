using System;

namespace WindowsLayoutSnapshot
{
    internal class Window
    {
        public Window(IntPtr hwnd, string processName, Native.WINDOWPLACEMENT windowPlacement)
        {
            Handle = hwnd;
            ProcessName = processName;
            WindowPlacement = windowPlacement;
        }

        public IntPtr Handle { get; }
        public string ProcessName { get; }
        public Native.WINDOWPLACEMENT WindowPlacement { get; }

        private bool IsWindowVisible() => Native.IsWindowVisible(Handle);
        private bool IsAppWindow() => (Native.GetWindowLong(Handle, Native.GWL_EXSTYLE) & (uint)Native.WS_EX_APPWINDOW) == (uint)Native.WS_EX_APPWINDOW;
        private bool IsToolWindow() => (Native.GetWindowLong(Handle, Native.GWL_EXSTYLE) & (uint)Native.WS_EX_TOOLWINDOW) == (uint)Native.WS_EX_TOOLWINDOW;
        private bool IsOwner() => Native.GetWindow(Handle, Native.GetWindowCmd.GW_OWNER) == IntPtr.Zero;

        public bool IsRestorableWindow() => IsWindowVisible() && IsOwner() && (IsAppWindow() || !IsToolWindow());
    }
}
