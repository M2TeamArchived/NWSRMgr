using System;
using System.Windows;

namespace NWSRMgr
{
    /// <summary>
    /// App.xaml 的交互逻辑
    /// </summary>
    public partial class App : Application
    {
        private void Application_Startup(object sender, StartupEventArgs e)
        {
            if (Environment.OSVersion.Version.Major >= 6)
            {
                StartupUri = new Uri("MainWindow.xaml", UriKind.RelativeOrAbsolute);
            }
            else
            {
                MessageBox.Show("本程序只支持Windows Vista以上操作系统", "NWSRMgr");
            }
        }
    }
}
