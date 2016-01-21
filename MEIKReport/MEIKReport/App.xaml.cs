using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Diagnostics;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;

namespace MEIKReport
{
    /// <summary>
    /// Interaction logic for App.xaml
    /// </summary>
    public partial class App : Application
    {
        private const int MINIMUM_SPLASH_TIME = 500; // Miliseconds
        private const int SPLASH_FADE_TIME = 500;     // Miliseconds
        public static IntPtr splashWinHwnd = IntPtr.Zero;
        public static Window opendWin = null;
        public static string dataFolder = null;

        protected override void OnStartup(StartupEventArgs e)
        {
            //// Step 1 - Load the splash screen
            //SplashScreen splash = new SplashScreen("splash.png");
            //splash.Show(false, true);

            //// Step 2 - Start a stop watch
            //Stopwatch timer = new Stopwatch();
            //timer.Start();

            //// Step 3 - Load your windows but don't show it yet
            //base.OnStartup(e);
            //MainWindow1 main = new MainWindow1();

            //// Step 4 - Make sure that the splash screen lasts at least two seconds
            //timer.Stop();
            //int remainingTimeToShowSplash = MINIMUM_SPLASH_TIME - (int)timer.ElapsedMilliseconds;
            //if (remainingTimeToShowSplash > 0)
            //    Thread.Sleep(remainingTimeToShowSplash);

            //// Step 5 - show the page
            //splash.Close(TimeSpan.FromMilliseconds(SPLASH_FADE_TIME));
            //main.Show();
        }

        /// <summary>
        /// 获取当前操作系统版本
        /// </summary>
        private void GetOsVersion()
        {
            // Get OperatingSystem information from the system namespace.
            System.OperatingSystem osInfo = System.Environment.OSVersion;

            // Determine the platform.
            switch (osInfo.Platform)
            {
                // Platform is Windows 95, Windows 98, 
                // Windows 98 Second Edition, or Windows Me.
                case System.PlatformID.Win32Windows:

                    switch (osInfo.Version.Minor)
                    {
                        case 0:
                            Console.WriteLine("Windows 95");
                            break;
                        case 10:
                            if (osInfo.Version.Revision.ToString() == "2222A")
                                Console.WriteLine("Windows 98 Second Edition");
                            else
                                Console.WriteLine("Windows 98");
                            break;
                        case 90:
                            Console.WriteLine("Windows Me");
                            break;
                    }
                    break;

                // Platform is Windows NT 3.51, Windows NT 4.0, Windows 2000,
                // or Windows XP.
                case System.PlatformID.Win32NT:

                    switch (osInfo.Version.Major)
                    {
                        case 3:
                            Console.WriteLine("Windows NT 3.51");
                            break;
                        case 4:
                            Console.WriteLine("Windows NT 4.0");
                            break;
                        case 5:
                            if (osInfo.Version.Minor == 0)
                                Console.WriteLine("Windows 2000");
                            else
                                Console.WriteLine("Windows XP");
                            break;
                    } break;
            }
        }
    }
}
