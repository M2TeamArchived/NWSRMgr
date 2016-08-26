using System;
using System.Runtime.InteropServices;

namespace M2
{
    class DPIWrapper
    {
        // 不可见Win32API定义
        #region PrivateWin32APIs

        [DllImport("user32.dll", CharSet = CharSet.Unicode,
            EntryPoint = "EnableChildWindowDpiMessage")]
        private static extern bool _EnableChildWindowDpiMessage(
            IntPtr hWnd,
            bool bEnable);

        [DllImport("win32u.dll", CharSet = CharSet.Unicode,
            EntryPoint = "NtUserEnableChildWindowDpiMessage")]
        private static extern bool _NtUserEnableChildWindowDpiMessage(
            IntPtr hWnd,
            bool bEnable);

        [DllImport("SHCore.dll", CharSet = CharSet.Unicode,
            EntryPoint = "GetDpiForMonitor")]
        private static extern long _GetDpiForMonitor(
            IntPtr hmonitor,
            MONITOR_DPI_TYPE dpiType,
            ref uint dpiX,
            ref uint dpiY);

        #endregion

        // 通知系统，让系统自动缩放非客户区(至少Win10)
        public static bool EnableChildWindowDpiMessage(
            IntPtr hWnd, bool bEnable)
        {
            bool bRet = false;

            try
            {
                bRet = _EnableChildWindowDpiMessage(hWnd, bEnable);
            }
            catch (DllNotFoundException) { }
            catch (EntryPointNotFoundException) { }

            if (!bRet)
            {
                try
                {
                    bRet = _NtUserEnableChildWindowDpiMessage(hWnd, bEnable);
                }
                catch (DllNotFoundException) { }
                catch (EntryPointNotFoundException) { }
            }

            return bRet;
        }

        public const long E_NOINTERFACE = 0x80004002L;
        public const long S_OK = 0x00000000L;

        public enum MONITOR_DPI_TYPE
        {
            MDT_EFFECTIVE_DPI = 0,
            MDT_ANGULAR_DPI = 1,
            MDT_RAW_DPI = 2,
            MDT_DEFAULT = MDT_EFFECTIVE_DPI
        }

        // Windows 8.1 新增的DPI获取API
        public static long GetDpiForMonitor(
            IntPtr hmonitor,
            MONITOR_DPI_TYPE dpiType,
            ref uint dpiX,
            ref uint dpiY)
        {
            long hr = E_NOINTERFACE;

            try
            {
                hr = _GetDpiForMonitor(
                    hmonitor, dpiType, ref dpiX, ref dpiY);
            }
            catch (DllNotFoundException) { }
            catch (EntryPointNotFoundException) { }

            return hr;
        }

        // 调用GetDpiForMonitor获取窗口的DPI
        public static long GetWindowDpi(
            IntPtr hWnd,
            ref uint dpiX,
            ref uint dpiY)
        {
            IntPtr hMonitor = MonitorFromWindow(
                hWnd, MonitorFlags.MONITOR_DEFAULTTONEAREST);

            return GetDpiForMonitor(
                hMonitor,
                MONITOR_DPI_TYPE.MDT_EFFECTIVE_DPI,
                ref dpiX,
                ref dpiY);
        }

        public enum MonitorFlags
        {
            MONITOR_DEFAULTTONULL = 0x00000000,
            MONITOR_DEFAULTTOPRIMARY = 0x00000001,
            MONITOR_DEFAULTTONEAREST = 0x00000002,
        }

        // 依据窗口句柄获取显示器句柄
        [DllImport("User32.dll")]
        public static extern IntPtr MonitorFromWindow(
            IntPtr hwnd, MonitorFlags dwFlags);

        public const int WM_DPICHANGED = 0x02E0;

        [StructLayout(LayoutKind.Sequential)]
        public struct RECT
        {
            public int left;
            public int top;
            public int right;
            public int bottom;
        }

        public static ushort LOWORD(uint value)
        {
            return (ushort)(value & 0xFFFF);
        }
        public static ushort HIWORD(uint value)
        {
            return (ushort)(value >> 16);
        }
    }
}
