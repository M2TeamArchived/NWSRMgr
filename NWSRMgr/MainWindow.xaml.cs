using System;
using System.Diagnostics;
using System.IO;
using System.Management;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Interop;
using System.Windows.Media;
using M2;

namespace NWSRMgr
{
    /// <summary>
    /// MainWindow.xaml 的交互逻辑
    /// </summary>
    public partial class MainWindow : Window
    {
        SystemRestore SR = new SystemRestore();

        [DllImport("kernel32.dll", EntryPoint = "CreateSymbolicLinkW", CharSet = CharSet.Unicode)]
        static extern int CreateSymbolicLink([In] string lpSymlinkFileName, [In] string lpTargetFileName, int dwFlags);

        public MainWindow()
        {
            new WpfPerMonitorDPIAwareSupport(this);

            InitializeComponent();
        }
        
        public void RefreshList()
        {

            /*PS C:\windows\system32> get-wmiObject -class SystemRestore -namespace "root\default"

            CreationTime           Description                    SequenceNumber    EventType         RestorePointType
            ------------           -----------                    --------------    ---------         ----------------
            2014/5/10 9:26:29      Good                           18                BEGIN_SYSTEM_C... 9*/

            //get-wmiObject -class Win32_ShadowCopy -namespace "root/CIMV2"

            listView.Items.Clear();
       
            foreach (ManagementObject SRInfo in SR.SRObject.Get())
            {
                SystemRestorePoint Item = new SystemRestorePoint();

                Item.SequenceNumber = Convert.ToInt32(SRInfo["SequenceNumber"]);

                DateTime CreationTime = SR.ConvertTime(SRInfo["CreationTime"].ToString()).ToLocalTime();
                Item.CreationTime = CreationTime.ToString();
                try
                {
                    Item.Description = SRInfo["Description"].ToString();
                }
                catch
                {
                    Item.Description = "未命名";
                }
                Item.RestorePointType = SR.GetRestorePointType(
                    Convert.ToInt32(SRInfo["RestorePointType"]));

                try
                {
                    long minvalue = -1;

                    foreach (ManagementObject VSSCopyInfo in SR.VSSCopyObject.Get())
                    {
                        DateTime InstallTime = SR.ConvertTime(VSSCopyInfo["InstallDate"].ToString());

                        long value =  Math.Abs(CreationTime.Ticks - InstallTime.Ticks);

                        if (minvalue == -1 || value < minvalue)
                        {
                            Item.DeviceObject = VSSCopyInfo["DeviceObject"].ToString() + "\\";
                            minvalue = value;
                        }
                    }
                }
                catch{ }

                listView.Items.Add(Item);
            }


            StatusLabel.Content = "已使用：";

            try
            {
                long SRSize = SR.GetUsedSize();

                string result;
                double num = SRSize >> 20;
                if (num > 1024)
                {
                    result = string.Format("{0:0.0}", num / 1024) + " GB";
                }
                else
                {
                    result = string.Format("{0:0.0}", num) + " MB";
                }

                StatusLabel.Content += result;
            }
            catch
            {
                StatusLabel.Content += "0 MB";
            }

            StatusLabel.Content += "\r\n系统还原状态：";

            if (SR.IsSREnabled())
            {
                StatusLabel.Content += "开启";
                SetStatusButton.Content = "禁用系统还原";
            }
            else
            {
                StatusLabel.Content += "关闭";
                SetStatusButton.Content = "启用系统还原";
            }
        }

        private void RefreshList_Click(object sender, RoutedEventArgs e)
        {
            RefreshList();
        }

        private void About_Click(object sender, RoutedEventArgs e)
        {
            MessageBox.Show(
                "NWSRMgr 2.1.1612.0\n© 2016 M2-Team. All rights reserved.",
                "关于 NWSRMgr",
                MessageBoxButton.OK,
                MessageBoxImage.Information);
        }

        private void Window_Initialized(object sender, EventArgs e)
        {
            RefreshList();
        }

        private void DeleteAll_Click(object sender, RoutedEventArgs e)
        {
            MessageBoxResult result;
            result = MessageBox.Show(
                "是否删除全部系统还原点",
                "NWSRMgr", 
                MessageBoxButton.YesNo,
                MessageBoxImage.Question,
                MessageBoxResult.No,
                MessageBoxOptions.DefaultDesktopOnly);
            if (result == MessageBoxResult.Yes)
            {
                if (SR.DeleteRestorePoints())
                {
                    MessageBox.Show("删除全部系统还原点成功", "NWSRMgr");
                }
                else
                {
                    MessageBox.Show("删除全部系统还原点失败", "NWSRMgr");
                }
                RefreshList();
            }
        }

        private void SetStatus_Click(object sender, RoutedEventArgs e)
        {
            MessageBoxResult result;
            if (SR.IsSREnabled())
            {
                result = MessageBox.Show(
                    "是否禁用系统还原",
                    "NWSRMgr",
                    MessageBoxButton.YesNo,
                    MessageBoxImage.Question,
                    MessageBoxResult.No,
                    MessageBoxOptions.DefaultDesktopOnly);
                if (result == MessageBoxResult.Yes)
                {
                    if (SR.Disable(null))
                    {
                        MessageBox.Show("系统还原关闭成功", "NWSRMgr");
                    }
                    else
                    {
                        MessageBox.Show("系统还原关闭失败", "NWSRMgr");
                    }
                }         
            }
            else
            {
                result = MessageBox.Show(
                    "是否启用系统还原",
                    "NWSRMgr",
                    MessageBoxButton.YesNo,
                    MessageBoxImage.Question,
                    MessageBoxResult.No,
                    MessageBoxOptions.DefaultDesktopOnly);
                if (result == MessageBoxResult.Yes)
                {
                    if (SR.Enable(null))
                    {
                        MessageBox.Show("系统还原开启成功", "NWSRMgr");
                    }
                    else
                    {
                        MessageBox.Show("系统还原开启失败", "NWSRMgr");
                    }
                }
            }
            RefreshList();
        }

        private void CreateRP_Click(object sender, RoutedEventArgs e)
        {
            CreateRP_Grid.Visibility = Visibility.Visible;
        }

        private void CreateRP_CancelButton_Click(object sender, RoutedEventArgs e)
        {
            CreateRP_Grid.Visibility = Visibility.Hidden;
        }

        private void CreateRP_OKButton_Click(object sender, RoutedEventArgs e)
        {
            if (CreateRP_TextBox.Text == "")
            {
                MessageBox.Show("还原点名称不能为空", "NWSRMgr");
            }
            else
            {
                if (SR.CreateRestorePoint(CreateRP_TextBox.Text))
                {
                    MessageBox.Show("系统还原点创建成功", "NWSRMgr");
                }
                else
                {
                    MessageBox.Show("系统还原点创建失败", "NWSRMgr");
                }
                CreateRP_TextBox.Text = null;
                CreateRP_Grid.Visibility = Visibility.Hidden;
                RefreshList();
            }
        }

        private void DeleteRP_Click(object sender, RoutedEventArgs e)
        {
            MessageBoxResult result;
            SystemRestorePoint SelectedItem = 
                (SystemRestorePoint)listView.SelectedItem;
            if (SelectedItem == null)
            {
                MessageBox.Show("请选择一个系统还原点", "NWSRMgr");
            }
            else
            {
                int RPNum = SelectedItem.SequenceNumber;
                result = MessageBox.Show(
                    "确定删除还原点" + RPNum.ToString(),
                    "NWSRMgr",
                    MessageBoxButton.YesNo,
                    MessageBoxImage.Question,
                    MessageBoxResult.No,
                    MessageBoxOptions.DefaultDesktopOnly);
                if (result == MessageBoxResult.Yes)
                {
                    if (SR.DeleteRestorePoint(RPNum))
                    {
                        MessageBox.Show(
                            "还原点" + RPNum.ToString() + "删除成功",
                            "NWSRMgr");
                    }
                    else
                    {
                        MessageBox.Show(
                            "还原点" + RPNum.ToString() + "删除失败",
                            "NWSRMgr");
                    }
                }
            }
        }

        private void MountRP_Click(object sender, RoutedEventArgs e)
        {
            MessageBoxResult result;
            SystemRestorePoint SelectedItem =
                (SystemRestorePoint)listView.SelectedItem;
            if (SelectedItem == null)
            {
                MessageBox.Show("请选择一个系统还原点", "NWSRMgr");
            }
            else
            {
                int RPNum = SelectedItem.SequenceNumber;
                string[] array = SelectedItem.DeviceObject.Split(new char[] { '\\' });
                string LinkName = Environment.ExpandEnvironmentVariables("%SystemDrive%") + "\\" + array[array.Length - 2];

                if (!Directory.Exists(LinkName))
                {
                    result = MessageBox.Show(
                        "确定挂载还原点" + RPNum.ToString(),
                        "NWSRMgr",
                        MessageBoxButton.YesNo,
                        MessageBoxImage.Question,
                        MessageBoxResult.No,
                        MessageBoxOptions.DefaultDesktopOnly);
                    if (result == MessageBoxResult.Yes)
                    {
                        CreateSymbolicLink(LinkName, SelectedItem.DeviceObject, 1);
                        Process.Start(LinkName);
                    }
                }
                else
                {
                    result = MessageBox.Show(
                        "还原点" + RPNum.ToString() + "存在，是否卸载还原点",
                        "NWSRMgr",
                        MessageBoxButton.YesNo,
                        MessageBoxImage.Question,
                        MessageBoxResult.No,
                        MessageBoxOptions.DefaultDesktopOnly);
                    if (result == MessageBoxResult.Yes)
                    {
                        Directory.Delete(LinkName);
                    }
                }    
            }
        }

        private void RestoreRP_Click(object sender, RoutedEventArgs e)
        {
            MessageBoxResult result;
            SystemRestorePoint SelectedItem =
                (SystemRestorePoint)listView.SelectedItem;
            if (SelectedItem == null)
            {
                MessageBox.Show("请选择一个系统还原点", "NWSRMgr");
            }
            else
            {
                int RPNum = SelectedItem.SequenceNumber;
                result = MessageBox.Show(
                    "确定从还原点" + RPNum.ToString() + "还原",
                    "NWSRMgr",
                    MessageBoxButton.YesNo,
                    MessageBoxImage.Question,
                    MessageBoxResult.No,
                    MessageBoxOptions.DefaultDesktopOnly);
                if (result == MessageBoxResult.Yes)
                {
                    if (SR.RestoreFromRestorePoint(RPNum))
                    {
                        result = MessageBox.Show(
                            "系统还原成功,是否重启电脑以继续",
                            "NWSRMgr",
                            MessageBoxButton.YesNo,
                            MessageBoxImage.Question,
                            MessageBoxResult.No,
                            MessageBoxOptions.DefaultDesktopOnly);
                        if (result == MessageBoxResult.Yes)
                        {
                            ExitWindows.Reboot();
                        }
                    }
                    else
                    {
                        MessageBox.Show("系统还原失败", "NWSRMgr");
                    }
                }
            }
        }
    }
}