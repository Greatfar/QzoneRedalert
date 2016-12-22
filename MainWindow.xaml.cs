﻿using System;
using System.Windows;
using System.Windows.Navigation;
using System.Configuration;
using System.Runtime.InteropServices;       //清空session时用到
using Forms = System.Windows.Forms;         //最小化到系统托盘
using System.Threading;

namespace RedAlert
{
    /// <summary>
    /// MainWindow.xaml 的交互逻辑
    /// </summary>
    public partial class MainWindow : Window
    {
        //类成员变量
        string game_url;                                        //免登陆URL
        bool isLogin = false;                                   //登录状态。不能直接在复选框xml绑定的函数上关闭定时器，因为复选框初始化时，定时器还没创建，所以先记录下来
        System.Windows.Threading.DispatcherTimer timer;         //定时器变量
        bool isFresh = false;                                   //刷新状态。如果是刷新状态，则不再向配置文件写入数据
        bool isAuto = false;                                    //自动脚本标记
        Thread threadZZ = null;                                 //自动征战线程


        private Forms.NotifyIcon notifyIcon;                   //最小化到系统托盘变量

        //-------------------------模拟鼠标事件常量------------------------
        const int MOUSEEVENTF_MOVE = 0x0001;        //移动鼠标
        const int MOUSEEVENTF_LEFTDOWN = 0x0002;    //模拟鼠标左键按下
        const int MOUSEEVENTF_LEFTUP = 0x0004;      //模拟鼠标左键抬起
        const int MOUSEEVENTF_RIGHTDOWN = 0x0008;   //模拟鼠标右键按下
        const int MOUSEEVENTF_RIGHTUP = 0x0010;     //模拟鼠标右键抬起
        const int MOUSEEVENTF_MIDDLEDOWN = 0x0020;  //模拟鼠标中键按下
        const int MOUSEEVENTF_MIDDLEUP = 0x0040;    //模拟鼠标中键抬起
        const int MOUSEEVENTF_ABSOLUTE = 0x8000;    //标示是否采用绝对坐标


        //-------------------------导入windows API库函数用于模拟鼠标事件----------------------------------
        //声明库中的鼠标事件函数
        [System.Runtime.InteropServices.DllImport("user32")]
        private static extern int mouse_event(int dwFlags, int dx, int dy, int cButtons, int dwExtraInfo);
        //导入设置鼠标位置外部函数
        [DllImport("user32.dll")]
        static extern bool SetCursorPos(int X, int Y);
        //------------------------------------------------------------------------------------------------


        //------------------------导入清空session用到的函数----------------------------------------------------------------
        private const int INTERNET_OPTION_END_BROWSER_SESSION = 42;
        [DllImport("wininet.dll", SetLastError = true)]
        private static extern bool InternetSetOption(IntPtr hInternet, int dwOption, IntPtr lpBuffer, int lpdwBufferLength);
        //------------------------------------------------------------------------------------------------------------------




    /// <summary>
    /// ========================================主窗体函数=================================================
    /// </summary>
    public MainWindow()
        {
            InitializeComponent();      //初始化组件

            //--------------------------------------------程序驻留系统托盘----------------------------------------------------------
            this.notifyIcon = new Forms.NotifyIcon();
            this.notifyIcon.BalloonTipText = "辅佐程序已经最小到系统托盘";
            this.notifyIcon.ShowBalloonTip(2000);
            this.notifyIcon.Text = "红警大战";
            this.notifyIcon.Icon = System.Drawing.Icon.ExtractAssociatedIcon(System.Windows.Forms.Application.ExecutablePath);      //读取程序图标作为托盘图标
            this.notifyIcon.Visible = true;
            //添加系统托盘图标的-打开菜单项
            System.Windows.Forms.MenuItem open = new System.Windows.Forms.MenuItem("打开");
            open.Click += new EventHandler(Show);
            //添加系统托盘图标的-退出菜单项
            System.Windows.Forms.MenuItem exit = new System.Windows.Forms.MenuItem("退出");
            exit.Click += new EventHandler(Close);
            //关联托盘控件菜单
            System.Windows.Forms.MenuItem[] childen = new System.Windows.Forms.MenuItem[] { open, exit };
            notifyIcon.ContextMenu = new System.Windows.Forms.ContextMenu(childen);
            //添加双击打开程序图标
            this.notifyIcon.MouseDoubleClick += new System.Windows.Forms.MouseEventHandler((o, e) =>
            {
                if (e.Button == Forms.MouseButtons.Left) this.Show(o, e);
            });
            //-----------------------------------------------------------------------------------------------------------------------


            //指定窗口显示位置
            WindowStartupLocation = WindowStartupLocation.Manual;       //自定义窗口显示位置
            this.Top = 0;                       //距离顶部
            this.Left = 188;                    //距离左边

            string strFreshT = ConfigurationManager.AppSettings["FreshT"];          //从配置文件读取刷新周期FreshT
            int freshT = Convert.ToInt32(strFreshT);                                //转换成整形，赋值给定时器周期变量

            //设置定时器
            timer = new System.Windows.Threading.DispatcherTimer();
            timer.Tick += new EventHandler(timer_Refresh);                          //为定时器时间绑定调用函数
            timer.Interval = new TimeSpan(0, freshT, 0);                            //设置定时值：TimeSpan（时, 分， 秒）。freshT分钟刷新一次

            //获取取当前日期
            string currentTime = System.DateTime.Now.ToString("yyyy-MM-dd");
            //MessageBox.Show(currentTime);

            //读取配置文件中的时间
            string lastTime = ConfigurationManager.AppSettings["Date"];
            //MessageBox.Show(lastTime);

            //判断当前时间和上次登录是否是同一天
            if (lastTime == currentTime)
            {

                //从配置文件读取登录状态，Status
                string myStatus = ConfigurationManager.AppSettings["Status"];

                //判断是否第一次登录
                if (myStatus == "secondtime")
                {
                    game_url = ConfigurationManager.AppSettings["Game"];            //从配置配置文件读取免登陆URL
                    Uri uri = new Uri(game_url);                                    //URL转换为URI
                    web1.Navigate(uri);                                             //web控件载入URL
                    timer.Start();                                                  //启动定时器
                }
                else
                {
                    //打开红警登录页面
                    Uri uri = new Uri("http://qqapp.qq.com/app/100616028.html");
                    web1.Navigate(uri);
                }
            }
            else          //如果不是同一天。因为免登陆URL一天后会失效。
            {

                //把当前时间写入配置文件
                Configuration cfa = ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None);
                cfa.AppSettings.Settings["Date"].Value = currentTime;
                cfa.Save();

                //打开红警登录页面
                Uri uri = new Uri("http://qqapp.qq.com/app/100616028.html");
                web1.Navigate(uri);
            }

            //绑定文档加载完成的回调函数，到web控件中,onLoadDocCompleted
            web1.LoadCompleted += new LoadCompletedEventHandler(onLoadDocCompleted);

        }




        //定时器回调函数
        private void timer_Refresh(object sender, EventArgs e)
        {
            //判断登录状态，如果没有登录就不执行刷新
            if (isLogin && !isAuto)
            { 
                Uri uri = new Uri(game_url);        //web控件，加载上一次登录的URL
                web1.Navigate(uri);
            }
        }




        //网页加载完成，回调函数
        private void onLoadDocCompleted(object sender, NavigationEventArgs e)
        {
            if (!isFresh)       //如果刷新状态不执行任何动作
            {
                //获取文档对象
                mshtml.IHTMLDocument2 doc2 = (mshtml.IHTMLDocument2)web1.Document;

                //判断：登录状态。登录状态通过网页标题判断。
                if (doc2.title == "红警之坦克风暴 - 应用中心")
                {
                    //获取游戏框架的iframe节点
                    mshtml.IHTMLElement m_iframe = (mshtml.IHTMLElement)doc2.all.item("app_frame", 0);
                    //获取该节点的src属性。即为游戏的免登陆URL
                    game_url = (string)m_iframe.getAttribute("src");
                    //把提取到的免登陆URL写入配置文件
                    Configuration cfa = ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None);
                    cfa.AppSettings.Settings["Game"].Value = game_url;               //写入游戏URL
                    cfa.AppSettings.Settings["Status"].Value = "secondtime";         //写入登录次数标记，用于第二次免登陆
                    cfa.Save();
                    //web控件重新载入免登陆URL。可达到去除腾讯QQ空间网页应用的外部框架，去除广告的目的目的。
                    Uri uri = new Uri(game_url);
                    web1.Navigate(uri);
                    //启动定时器，用于刷新计时
                    timer.Start();
                    //把刷新状态改为真，表示当前是刷新状态
                    isFresh = true;
                }
            }
        }




        //复选框选中时自动调用，在xaml中绑定
        private void CheckBox_Checked(object sender, RoutedEventArgs e)
        {
            isLogin = true;
        }




        //复选框取消选中时自动调用，在xaml中绑定
        private void CheckBox_Unchecked(object sender, RoutedEventArgs e)
        {
            isLogin = false;
        }




        //切换账号，按钮点击事件函数
        private void button_Click(object sender, RoutedEventArgs e)
        {
            //改变配置文件文件中的登录状态、免登陆URL。达到再次启动应用时自动打开登录页面
            Configuration cfa = ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None);
            cfa.AppSettings.Settings["Game"].Value = "reLogin"; 
            cfa.AppSettings.Settings["Status"].Value = "exit"; 
            cfa.Save();

            //清空session
            InternetSetOption(IntPtr.Zero, INTERNET_OPTION_END_BROWSER_SESSION, IntPtr.Zero, 0);

            //重新跳转到登录页面
            Uri uri = new Uri("http://qqapp.qq.com/app/100616028.html");
            web1.Navigate(uri);

            //把刷新状态改为未刷新，保证切换账号后，依然会自动获取免登陆URL，并重新载入
            isFresh = false;

            //消息提醒框
            // MessageBox.Show("登录要切换的QQ，可快速切换游戏", "提示");
        }




        //软件托盘菜单项响应函数，显示窗口
        private void Show(object sender, EventArgs e)
        {
            this.Visibility = System.Windows.Visibility.Visible;
            this.ShowInTaskbar = true;
            this.Activate();
        }




        //软件托盘菜单项响应函数，隐藏窗口
        private void Hide(object sender, EventArgs e)
        {
            this.ShowInTaskbar = false;
            this.Visibility = System.Windows.Visibility.Hidden;
        }




        //软件托盘菜单项响应函数，退出程序
        private void Close(object sender, EventArgs e)
        {
            System.Windows.Application.Current.Shutdown();
        }




        //重写关闭按钮（窗口右上角关闭按钮）实现方法。（已经通过xaml绑定）
        private void MainWindow_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            MessageBoxResult result = MessageBox.Show("是否最小化到系统托盘，后台运行？", "关闭选项", MessageBoxButton.YesNoCancel, MessageBoxImage.Question);
            if (result == MessageBoxResult.Yes)
            {
                Hide();             //调用隐藏窗口方法
                e.Cancel = true;    //并向窗口传递取消关闭
            }
            else if (result == MessageBoxResult.No)
            {
                e.Cancel = false;   //向窗口传递不取消关闭窗口消息
            }
            else
            {
                e.Cancel = true;        //取消关闭窗口
            }
        }




        //自动征战-按钮点击事件响应函数
        private void zhengzhan_Click(object sender, RoutedEventArgs e)
        {
            //开启一条新的线程
            threadZZ = new Thread(AutoZZ);
            threadZZ.Start();      
        }




        //自动征战-线程函数
        private void AutoZZ()
        {
            //设置自动脚本标记，防止刷新
            isAuto = true;

            //------------通过点击垂直箭头，复位初始状态-------------------------------------------
            //设置光标的坐标，垂直滚动条向上箭头
            SetCursorPos(1148, 35);
            //鼠标左键按下并弹起（单击一次），重复4次
            mouse_event(MOUSEEVENTF_LEFTDOWN | MOUSEEVENTF_LEFTUP, 0, 0, 0, 0);
            mouse_event(MOUSEEVENTF_LEFTDOWN | MOUSEEVENTF_LEFTUP, 0, 0, 0, 0);
            mouse_event(MOUSEEVENTF_LEFTDOWN | MOUSEEVENTF_LEFTUP, 0, 0, 0, 0);
            mouse_event(MOUSEEVENTF_LEFTDOWN | MOUSEEVENTF_LEFTUP, 0, 0, 0, 0);


            //设置鼠标的坐标，垂直滚动条向下箭头
            SetCursorPos(1148, 678);
            //鼠标左键按下并弹起（单击一次）
            mouse_event(MOUSEEVENTF_LEFTDOWN | MOUSEEVENTF_LEFTUP, 0, 0, 0, 0);

            //等待2秒
            Thread.Sleep(2000);

            //设置鼠标的坐标，征战图标
            SetCursorPos(995, 669);
            //鼠标左键按下并弹起（单击一次）
            mouse_event(MOUSEEVENTF_LEFTDOWN | MOUSEEVENTF_LEFTUP, 0, 0, 0, 0);

            //等待5秒
            Thread.Sleep(5000);

            //第二次征战标记
            bool isSecondZZ = false;
            //读取配置文件中的征战次数
            string strZZTime = ConfigurationManager.AppSettings["ZZTime"];
            int zzTime = Convert.ToInt32(strZZTime);        //转换成整型

            for (int i=0; i< zzTime; i++)
            {
                
                //设置鼠标的坐标，进攻
                SetCursorPos(669, 611);
                //鼠标左键按下并弹起（单击一次）
                mouse_event(MOUSEEVENTF_LEFTDOWN | MOUSEEVENTF_LEFTUP, 0, 0, 0, 0);

                //等待3秒
                Thread.Sleep(3000);

                //设置鼠标的坐标，跳过
                SetCursorPos(1000, 675);
                //鼠标左键按下并弹起（单击一次）
                mouse_event(MOUSEEVENTF_LEFTDOWN | MOUSEEVENTF_LEFTUP, 0, 0, 0, 0);

                //等待2秒
                Thread.Sleep(2000);

                //设置鼠标的坐标，确定
                SetCursorPos(667, 576);
                //鼠标左键按下并弹起（单击一次）
                mouse_event(MOUSEEVENTF_LEFTDOWN | MOUSEEVENTF_LEFTUP, 0, 0, 0, 0);

                //等待6秒
                Thread.Sleep(6000);

                //设置鼠标的坐标，奖品领取
                SetCursorPos(751, 475);
                //鼠标左键按下并弹起（单击一次）
                mouse_event(MOUSEEVENTF_LEFTDOWN | MOUSEEVENTF_LEFTUP, 0, 0, 0, 0);

                //等待8秒
                Thread.Sleep(8000);

                //启用重新征战（每天可征战两次）
                if (i == (zzTime -2) && !isSecondZZ)
                {
                    isSecondZZ = true;      //第二次征战标记
                    i = 0;                  //重置循环变量

                    //设置鼠标的坐标，重新征战
                    SetCursorPos(780, 611);
                    //鼠标左键按下并弹起（单击一次）
                    mouse_event(MOUSEEVENTF_LEFTDOWN | MOUSEEVENTF_LEFTUP, 0, 0, 0, 0);

                    //等待6秒
                    Thread.Sleep(6000);
                }
            }

            //关闭自动脚本标记，进行刷新
            isAuto = false;
        }



        //停止脚本-按钮点击事件响应函数
        private void stopAuto_Click(object sender, RoutedEventArgs e)
        {
            if(threadZZ != null)
            { 
                threadZZ.Abort();   //关闭自动征战线程
            }

			//关闭自动脚本标记，进行刷新
            isAuto = false;
        }
    }
}
