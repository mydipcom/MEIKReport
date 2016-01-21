using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using System.IO;
using System.Diagnostics;

namespace MeikControlLibrary
{
    public partial class MeikUserControl : UserControl
    {
        [DllImport("user32.dll")]
        public static extern int GetWindowLong(IntPtr hWnd, int nlndex);

        //该函数改变指定窗口的属性。函数也将在指定偏移地址的一个32 位值存入窗口的额外窗口存
        [DllImport("user32.dll", EntryPoint = "SetWindowLong")]
        public static extern int SetWindowLong(IntPtr hWnd, int nlndex, int dwNewLong);
        [DllImport("user32.dll", EntryPoint = "FindWindowEx")]
        public static extern IntPtr FindWindowEx(IntPtr hwndParent, IntPtr hwndChildAfter, string lpszClass, string lpszWindow);
        //设置指定窗体的父窗体
        [DllImport("user32.dll", SetLastError = true)]
        public static extern long SetParent(IntPtr hWndChild, IntPtr hWndNewParent);
        [DllImport("user32.dll")]
        public static extern bool MoveWindow(IntPtr hWnd, int x, int y, int nWidth, int nHeight, bool BRePaint);
        [DllImport("user32.dll")]
        public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);
        [DllImport("user32.dll")]
        public static extern bool IsWindow(IntPtr hWnd);
        public string AppPath { get; set; }

        

        public MeikUserControl()
        {
            InitializeComponent();
            
        }
        
        public void StartApp(){            
            IntPtr splashWinHwnd = FindWindowEx(IntPtr.Zero, IntPtr.Zero, "TfmSplash", null);
            if (!IsWindow(splashWinHwnd))
            {
                string meikPath = "C:\\Program Files (x86)\\MEIK 5.6\\MEIK.exe";
                if (File.Exists(meikPath))
                {
                    try
                    {
                        //启动外部程序
                        Process AppProc = Process.Start(meikPath);
                        if (AppProc != null)
                        {
                            AppProc.StartInfo.WindowStyle = ProcessWindowStyle.Hidden; //隐藏
                            //proc.WaitForExit();//等待外部程序退出后才能往下执行
                            AppProc.WaitForInputIdle();
                            /**这段代码是把其他程序的窗体嵌入到当前窗体中**/
                            IntPtr appWinHandle = AppProc.MainWindowHandle;
                            splashWinHwnd = FindWindowEx(IntPtr.Zero, IntPtr.Zero, "TfmSplash", null);
                            //监视进程退出
                            AppProc.EnableRaisingEvents = true;
                            //指定退出事件方法
                            AppProc.Exited += new EventHandler(proc_Exited);
                            
                        }
                    }
                    catch (ArgumentException ex)
                    {
                        MessageBox.Show("Failed to start MEIK software v. 5.6");
                    }
                }
                else
                {
                    MessageBox.Show("The file " + meikPath + "\\MEIK.exe is not exist.");
                }
            }
             
            
            MoveWindow(splashWinHwnd, 0, 0, 720, 576, true);
            // Set new process's parent to this window   
            SetParent(splashWinHwnd, this.Handle);
            //Add WS_CHILD window style to child window 
            const int GWL_STYLE = -16;
            const int WS_CHILD = 0x40000000;
            int style = GetWindowLong(splashWinHwnd, GWL_STYLE);
            style = style | WS_CHILD;
            SetWindowLong(splashWinHwnd, GWL_STYLE, style);            
        }

        //控件透明背景绘制 
        protected override void OnPaintBackground(PaintEventArgs e)
        {
            base.OnPaintBackground(e);
            Graphics g = e.Graphics;
            //以下代码提取控件原来应有的背景（视觉上的位图，可能会包括其他空间的部分图形）到位图，进行变换后绘制 
            if (Parent != null)
            {
                BackColor = Color.Transparent;
                //本控件在父控件的子集中的index 
                int index = Parent.Controls.GetChildIndex(this);
                for (int i = Parent.Controls.Count - 1; i > index; i--)
                {
                    //对每一个父控件的可视子控件，进行绘制区域交集运算，得到应该绘制的区域 
                    Control c = Parent.Controls[i];
                    //如果有交集且可见 
                    if (c.Bounds.IntersectsWith(Bounds) && c.Visible)
                    {                        
                        Bitmap bmp = new Bitmap(c.Width, c.Height, g);
                        c.DrawToBitmap(bmp, c.ClientRectangle);
                        g.TranslateTransform(c.Left - Left, c.Top - Top);
                        //画图 
                        g.DrawImageUnscaled(bmp, Point.Empty);
                        g.TranslateTransform(Left - c.Left, Top - c.Top);
                        bmp.Dispose();
                    }
                }
            }
            else
            {
                g.Clear(Parent.BackColor);
                g.FillRectangle(new SolidBrush(Color.FromArgb(255, Color.Transparent)), this.ClientRectangle);
            }
        }


        /// <summary>
        ///启动外部程序退出事件
        /// </summary>
        void proc_Exited(object sender, EventArgs e)
        {
            ControlEventArgs ce = new ControlEventArgs(this);
            this.OnControlRemoved(ce);
        }

        private void MeikUserControl_Load(object sender, EventArgs e)
        {
            StartApp();
        }
    }
}
