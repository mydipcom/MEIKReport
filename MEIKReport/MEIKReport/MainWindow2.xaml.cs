using MEIKReport.Common;
using MEIKReport.Views;
using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Interop;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.Xml;

namespace MEIKReport
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow2 : Window
    {        
        //定义将要运行的原始的MEIK程序的进程
        public Process AppProc { get; set; }

        private WindowListen windowListen = new WindowListen();

        //设置指定窗体的父窗体
        [DllImport("user32.dll", SetLastError = true)]
        private static extern long SetParent(IntPtr hWndChild, IntPtr hWndNewParent);
        //获取指定坐标处窗体句柄
        [DllImport("user32.dll", EntryPoint = "WindowFromPoint")]
        public static extern int WindowFromPoint(int xPoint, int yPoint);

        public MainWindow2()
        {
            InitializeComponent();
            string meikPath = OperateIniFile.ReadIniData("Base", "MEIK base", "C:\\Program Files (x86)\\MEIK 5.6", System.AppDomain.CurrentDomain.BaseDirectory + "Config.ini");
            meikPath += "\\MEIK.exe";
            if (File.Exists(meikPath))
            {
                try
                {
                    //启动外部程序
                    AppProc = Process.Start(meikPath);
                    if (AppProc != null)
                    {
                        AppProc.StartInfo.WindowStyle = ProcessWindowStyle.Hidden; //隐藏
                        //proc.WaitForExit();//等待外部程序退出后才能往下执行
                        AppProc.WaitForInputIdle();
                        /**这段代码是把其他程序的窗体嵌入到当前窗体中**/
                        IntPtr appWinHandle = AppProc.MainWindowHandle;
                        IntPtr mainWinHandle = new WindowInteropHelper(this).Handle;
                        SetParent(appWinHandle, mainWinHandle);
                        
                        //监视进程退出
                        AppProc.EnableRaisingEvents = true;
                        //指定退出事件方法
                        AppProc.Exited += new EventHandler(proc_Exited);
                       // BackButtonWindow backButtonWindow = new BackButtonWindow();
                        //backButtonWindow.Owner = this;
                        //backButtonWindow.Show();
                        //backButtonWindow.Hide();                        

                        //windowListen.AppProc = AppProc;
                        //windowListen.backWinHwnd = new WindowInteropHelper(backButtonWindow).Handle;
                        //侦听MEIK程序是否打开全屏窗体 
                        //Thread workerThread = new Thread(windowListen.Run);
                        //workerThread.Start();                        
                    }
                }
                catch (ArgumentException ex)
                {
                    this.Show();
                    MessageBox.Show("Failed to start MEIK software v. 5.6, Exception: " +ex.Message);
                }
            }
            else
            {
                this.Show();
                MessageBox.Show("The file " + meikPath + "\\MEIK.exe is not exist.");
            }          
        }

        private void equipmentList_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {            
           // userGrid.DataContext = equipmentCode.SelectedItem;
        }
        
        private void ShowMainWindow(object sender, EventArgs e)
        {
            this.Show();
        }

        private void Screen_Click(object sender, RoutedEventArgs e)
        {
            
            this.Hide();
            IntPtr splashWinHwnd = Win32Api.FindWindowEx(IntPtr.Zero, IntPtr.Zero, "TfmSplash", null);
            IntPtr screeningBtnHwnd = Win32Api.FindWindowEx(splashWinHwnd, IntPtr.Zero, null, "Screening");

            IntPtr hwnd = Win32Api.GetWindow(splashWinHwnd, 5);
            while (hwnd.ToInt32() > 0)
            {
                
                StringBuilder text = new StringBuilder(512);
                Win32Api.GetWindowText(hwnd, text, text.Capacity);
                MessageBox.Show("" + text.ToString());
                
                hwnd = Win32Api.GetWindow(hwnd, 2);
            }

            //Win32Api.SendMessage(screeningBtnHwnd, Win32Api.WM_CLICK, 0, 0);
            //IntPtr screeningBtnHwnd = Win32Api.FindWindowEx(splashWinHwnd, IntPtr.Zero, null, "Choose an operating mode");
            //Win32Api.ShowWindow(screeningBtnHwnd, 0);
            //Win32Api.SendMessage(screeningBtnHwnd, Win32Api.WM_CLICK, 0, 0);
        }

        private void Report_Click(object sender, RoutedEventArgs e)
        {
            this.Hide();
            UserList userList = new UserList();
            //summaryReportPage.closeWindowEvent += new CloseWindowHandler(ShowMainWindow);
            userList.Owner = this;
            userList.ShowDialog();
        }
        /// <summary>
        ///启动外部程序退出事件
        /// </summary>
        void proc_Exited(object sender, EventArgs e)
        {
            this.Dispatcher.Invoke(new Action(delegate {this.Show();}));
        }

        private void Window_Closed(object sender, EventArgs e)
        {
            windowListen.Stop();
            AppProc.CloseMainWindow();
        }
    }
}
