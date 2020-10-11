using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Interop;

namespace 收支统计
{
    /// <summary>
    /// App.xaml 的交互逻辑
    /// </summary>
    public partial class App : Application
    {
        protected override void OnStartup(StartupEventArgs e)
        {
            if(DateTime.Now> new DateTime(2020, 10, 20))
            {
                MessageBox.Show("请联系正式版！");
                Environment.Exit(0);
            }

            base.OnStartup(e);
        }
    }
}
