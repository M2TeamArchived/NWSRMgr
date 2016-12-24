/**************************************************************************
描述：为wpf窗口实现Per-Monitor DPI Aware支持
维护者：Mouri_Naruto (M2-Team)
版本：1.1 (2016-08-29)
基于项目：无
协议：The MIT License
用法：在wpf窗口类的构造函数添加 new WpfPerMonitorDPIAwareSupport(this);
建议的.Net Framework版本：4.0及以后
***************************************************************************/

using System;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Interop;
using System.Windows.Media;

namespace M2
{
    class WpfPerMonitorDPIAwareSupport
    {
        // 本类私有Win32API定义
        #region PrivateWin32

        private const int WM_DPICHANGED = 0x02E0;

        private const long S_OK = 0x00000000L;
        private const long E_NOINTERFACE = 0x80004002L;
        
        [StructLayout(LayoutKind.Sequential)]
        private struct RECT
        {
            public int left;
            public int top;
            public int right;
            public int bottom;
        }

        public enum MONITOR_FLAGS
        {
            MONITOR_DEFAULTTONULL = 0x00000000,
            MONITOR_DEFAULTTOPRIMARY = 0x00000001,
            MONITOR_DEFAULTTONEAREST = 0x00000002,
        }

        public enum MONITOR_DPI_TYPE
        {
            MDT_EFFECTIVE_DPI = 0,
            MDT_ANGULAR_DPI = 1,
            MDT_RAW_DPI = 2,
            MDT_DEFAULT = MDT_EFFECTIVE_DPI
        }

        private ushort LOWORD(uint value)
        {
            return (ushort)(value & 0xFFFF);
        }
        private ushort HIWORD(uint value)
        {
            return (ushort)(value >> 16);
        }

        private T PtrToStructure<T>(IntPtr Ptr)
        {
            return (T)Marshal.PtrToStructure(Ptr, typeof(T));
        }

        [DllImport("user32.dll", CharSet = CharSet.Unicode)]
        private static extern bool EnableChildWindowDpiMessage(
            IntPtr hWnd,
            bool bEnable);

        [DllImport("win32u.dll", CharSet = CharSet.Unicode)]
        private static extern bool NtUserEnableChildWindowDpiMessage(
            IntPtr hWnd,
            bool bEnable);

        [DllImport("SHCore.dll", CharSet = CharSet.Unicode,
            EntryPoint = "GetDpiForMonitor")]
        private static extern long GetDpiForMonitor(
            IntPtr hmonitor,
            MONITOR_DPI_TYPE dpiType,
            ref uint dpiX,
            ref uint dpiY);

        // 依据窗口句柄获取显示器句柄
        [DllImport("User32.dll")]
        private static extern IntPtr MonitorFromWindow(
            IntPtr hwnd, MONITOR_FLAGS dwFlags);

        #endregion

        // WPF窗口对象
        private Window m_WpfWindow;

        // WPF窗口缩放
        private void ScaleWpfWindow(ScaleTransform Scale)
        {
            VisualTreeHelper.GetChild(m_WpfWindow, 0).SetValue(
                FrameworkElement.LayoutTransformProperty, Scale);
        }

        // 处理WM_DPICHANGED消息
        private IntPtr OnDPIChanged(
            IntPtr hWnd,
            int message,
            IntPtr wParam,
            IntPtr lParam,
            ref bool handled)
        {
            if (WM_DPICHANGED == message)
            {
                uint _wParam = Convert.ToUInt32(wParam.ToInt32());

                RECT prcNewWindow = PtrToStructure<RECT>(lParam);

                ScaleTransform Scale = new ScaleTransform(
                    LOWORD(_wParam) / 96.0, HIWORD(_wParam) / 96.0);

                ScaleWpfWindow(Scale);

                m_WpfWindow.Left = prcNewWindow.left;
                m_WpfWindow.Top = prcNewWindow.top;

                m_WpfWindow.Width = prcNewWindow.right - prcNewWindow.left;
                m_WpfWindow.Height = prcNewWindow.bottom - prcNewWindow.top;
            }

            return IntPtr.Zero;
        }

        // WPF窗口Loaded事件处理
        private void OnLoaded(object sender, RoutedEventArgs e)
        {
            IntPtr hWnd = new WindowInteropHelper(m_WpfWindow).Handle;

            // 调用Windows 8.1新增的GetDpiForMonitor获取窗口的DPI
            long hr = E_NOINTERFACE;
            uint dpiX = 96; uint dpiY = 96;
            try
            {
                hr = GetDpiForMonitor(
                    MonitorFromWindow(
                        hWnd, 
                        MONITOR_FLAGS.MONITOR_DEFAULTTONEAREST), 
                    MONITOR_DPI_TYPE.MDT_EFFECTIVE_DPI,
                    ref dpiX,
                    ref dpiY);
            }
            catch (DllNotFoundException) { }
            catch (EntryPointNotFoundException) { }

            // 如获取DPI失败即视本环境不支持Per-Monitor DPI Aware，执行返回
            if (S_OK != hr) return;

            // 先通知系统，让系统自动缩放非客户区(至少Win10)
            bool bRet = false;
            try
            {
                bRet = EnableChildWindowDpiMessage(hWnd, true);
            }
            catch (DllNotFoundException) { }
            catch (EntryPointNotFoundException) { }

            try
            {
                if (!bRet) bRet= NtUserEnableChildWindowDpiMessage(hWnd, true);
            }
            catch (DllNotFoundException) { }
            catch (EntryPointNotFoundException) { }

            // 再执行WPF窗口缩放
            ScaleTransform Scale = new ScaleTransform(
                dpiX / 96.0, dpiY / 96.0);

            ScaleWpfWindow(Scale);

            m_WpfWindow.Width *= Scale.ScaleX;
            m_WpfWindow.Height *= Scale.ScaleY;

            // 最后加入WM_DPICHANGED消息处理函数
            HwndSource hwndSource =
                (HwndSource)PresentationSource.FromVisual(m_WpfWindow);
            if (null != hwndSource)
            {
                hwndSource.AddHook(new HwndSourceHook(OnDPIChanged));
            }
        }

        //构造函数
        public WpfPerMonitorDPIAwareSupport(Window WpfWindow)
        {
            m_WpfWindow = WpfWindow;
            m_WpfWindow.Loaded += new RoutedEventHandler(OnLoaded);
        }    
    }
}
