using Microsoft.Win32;
using System;
using System.Management;
using System.Runtime.InteropServices;

namespace NWSRMgr
{
    public class SystemRestorePoint
    {
        public int SequenceNumber { get; set; }
        public string CreationTime { get; set; }
        public string Description { get; set; }
        public string RestorePointType { get; set; }
        public string DeviceObject { get; set; }
    }

    public class SystemRestore
    {
        [StructLayout(LayoutKind.Sequential)]
        struct RESTOREPOINTINFO
        {
            public int dwEventType;
            public int dwRestorePtType;
            public long llSequenceNumber;
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 257)]
            public string szDescription;
        }

        [StructLayout(LayoutKind.Sequential)]
        struct STATEMGRSTATUS
        {
            public uint nStatus;
            public long llSequenceNumber;
        }

        [DllImport("SrClient.dll")]
        static extern bool SRSetRestorePoint(ref RESTOREPOINTINFO SRPInfo, ref STATEMGRSTATUS SRPStatus);

        [DllImport("SrClient.dll")]
        static extern int SRRemoveRestorePoint(int dwRPNum);

        [DllImport("SrClient.dll")]
        static extern int DisableSR([MarshalAs(UnmanagedType.LPWStr)]string Drive);

        [DllImport("SrClient.dll")]
        static extern int EnableSR([MarshalAs(UnmanagedType.LPWStr)]string Drive);

        public ManagementObjectSearcher SRObject = new ManagementObjectSearcher("root/default", "SELECT * FROM SystemRestore");
        public ManagementObjectSearcher VSSObject = new ManagementObjectSearcher("root/CIMV2", "SELECT * FROM Win32_ShadowStorage");
        public ManagementObjectSearcher VSSCopyObject = new ManagementObjectSearcher("root/CIMV2", "SELECT * FROM Win32_ShadowCopy");

        public bool Enable(string DriveName)
        {
            try
            {
                return (EnableSR(DriveName) == 0);
            }
            catch
            {
                return false;
            }
        }
        public bool Disable(string DriveName)
        {
            try
            {
                return (DisableSR(DriveName) == 0);
            }
            catch
            {
                return false;
            }
        }

        public bool IsSREnabled()
        {
            RegistryKey SPPClientKey = Registry.LocalMachine.OpenSubKey("Software\\Microsoft\\Windows NT\\CurrentVersion\\SPP\\Clients", true);
            return (SPPClientKey.ValueCount != 0);
        }

        public bool DeleteRestorePoint(int dwRPNum)
        {
            try
            {
                return (SRRemoveRestorePoint(dwRPNum) == 0);
            }
            catch
            {
                return false;
            }
        }

        public bool DeleteRestorePoints()
        {
            try
            {
                foreach (ManagementObject SRInfo in SRObject.Get())
                {
                    SRRemoveRestorePoint(Convert.ToInt32(SRInfo["SequenceNumber"]));
                }
                return true;
            }
            catch
            {
                return false;
            }
        }

        public bool CreateRestorePoint(string RPName)
        {
            try
            {
                RegistryKey SystemRestoreKey = Registry.LocalMachine.OpenSubKey("Software\\Microsoft\\Windows NT\\CurrentVersion\\SystemRestore", true);
                SystemRestoreKey.SetValue("SystemRestorePointCreationFrequency", 0, RegistryValueKind.DWord);

                bool result = false;
                RESTOREPOINTINFO RPInfo = new RESTOREPOINTINFO();
                STATEMGRSTATUS RPStatus = new STATEMGRSTATUS();

                RPInfo.dwEventType = 100;
                RPInfo.dwRestorePtType = 16;
                RPInfo.llSequenceNumber = 0;
                RPInfo.szDescription = RPName;

                result = SRSetRestorePoint(ref RPInfo, ref RPStatus);

                SystemRestoreKey.DeleteValue("SystemRestorePointCreationFrequency");
                return result;
            }
            catch
            {
                return false;
            }
        }

        public long GetUsedSize()
        {
            long SRSize = 0;
            try
            {
                foreach (ManagementObject VSSInfo in VSSObject.Get())
                {
                    SRSize += Convert.ToInt64(VSSInfo["AllocatedSpace"]);
                }
            }
            catch { }
            return SRSize;
        }

        public string GetRestorePointType(int RestorePointType)
        {
            switch (RestorePointType)
            {
                case 0:
                    return "应用安装";
                case 1:
                    return "应用卸载";
                case 6:
                    return "撤销";
                case 7:
                    return "检查点";
                case 10:
                    return "设备安装";
                case 11:
                    return "首次运行";
                case 12:
                    return "更改设置";
                case 13:
                    return "取消操作";
                case 14:
                    return "备份恢复";
                case 15:
                    return "备份";
                case 16:
                    return "手动";
                case 17:
                    return "Windows Update";
                case 18:
                    return "关键更新";
                default:
                    return "未知";
            }
        }

        public DateTime ConvertTime(string TimeStr)
        {
            int year = Convert.ToInt32(new string(TimeStr.ToCharArray(0, 4)));
            int month = Convert.ToInt32(new string(TimeStr.ToCharArray(4, 2)));
            int day = Convert.ToInt32(new string(TimeStr.ToCharArray(6, 2)));
            int hour = Convert.ToInt32(new string(TimeStr.ToCharArray(8, 2)));
            int minute = Convert.ToInt32(new string(TimeStr.ToCharArray(10, 2)));
            int second = Convert.ToInt32(new string(TimeStr.ToCharArray(12, 2)));
            return new DateTime(year, month, day, hour, minute, second, DateTimeKind.Utc);
        }

        public bool RestoreFromRestorePoint(int RPNum)
        {
            ManagementClass SRClass = new ManagementClass("//./root/default:SystemRestore");

            try
            {
                object[] SRArgs = { RPNum };
                SRClass.InvokeMethod("Restore", SRArgs);
                return true;
            }
            catch
            {
                return false;
            }
        }
    }
}

class ExitWindows
{
    [DllImport("ntdll.dll")]
    private static extern void RtlAdjustPrivilege(int Privilege, int Enable, int Thread, int OldValue);

    [DllImport("user32.dll", ExactSpelling = true, SetLastError = true)]
    private static extern bool ExitWindowsEx(int flg, int rea);

    public const int SE_SHUTDOWN_PRIVILEGE = 19;

    private const int EWX_LOGOFF = 0x00000000;
    private const int EWX_SHUTDOWN = 0x00000001;
    private const int EWX_REBOOT = 0x00000002;
    private const int EWX_FORCE = 0x00000004;
    private const int EWX_POWEROFF = 0x00000008;
    private const int EWX_FORCEIFHUNG = 0x00000010;

    public static void DoExitWin(int flg)
    {

        //give current process SeShutdownPrivilege
        RtlAdjustPrivilege(SE_SHUTDOWN_PRIVILEGE, 1, 0, 0);

        //Exit windows
        if (!ExitWindowsEx(flg, 0))
        {
            throw new Exception("Exit Windows fail");
        }
    }

    /// <summary>
    /// Reboot computer
    /// </summary>
    /// <param name="force">force reboot</param>
    public static void Reboot(bool force = false)
    {
        if (force)
        {
            DoExitWin(EWX_REBOOT | EWX_FORCE);
        }
        else
        {
            DoExitWin(EWX_REBOOT | EWX_FORCEIFHUNG);
        }
    }

    /// <summary>
    /// Shut down computer
    /// </summary>
    /// <param name="force">force shut down</param>
    public static void Shutdown(bool force = false)
    {
        if (force)
        {
            DoExitWin(EWX_SHUTDOWN | EWX_FORCE);
        }
        else
        {
            DoExitWin(EWX_SHUTDOWN | EWX_FORCEIFHUNG);
        }
    }

    /// <summary>
    /// Log off
    /// </summary>
    /// <param name="force">force logoff</param>
    public static void Logoff(bool force = false)
    {
        if (force)
        {
            DoExitWin(EWX_LOGOFF | EWX_FORCE);
        }
        else
        {
            DoExitWin(EWX_LOGOFF | EWX_FORCEIFHUNG);
        }
    }
}