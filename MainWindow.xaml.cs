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
        Thread threadGet = null;                                //自动收集线程

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
        //-------------------------------------------------------------------


        //------------------------导入屏幕取色相关win32 api函数------------------------------------------
        [DllImport("user32")]
        private static extern int GetWindowDC(int hwnd);        //获取句柄
        [DllImport("user32")]
        private static extern int ReleaseDC(int hWnd, int hDC); //释放句柄
        [DllImport("gdi32")]
        private static extern int GetPixel(int hdc, int nXPos, int nYPos);  //获取指定点的颜色
        //-----------------------------------------------------------------------------------------------


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
            this.notifyIcon.BalloonTipText = "红警大战已最小化到后台运行";
            this.notifyIcon.ShowBalloonTip(2000);
            this.notifyIcon.Text = "红警大战";
            //读取程序图标作为托盘图标
            this.notifyIcon.Icon = System.Drawing.Icon.ExtractAssociatedIcon(System.Windows.Forms.Application.ExecutablePath);
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
            if (isAuto)     //如果有其他动作线程在执行，先关闭掉,防止多个动作相互干扰
            {
                MessageBoxResult result = MessageBox.Show("正在执行其他自动操作，是否停止它！", "警告", MessageBoxButton.YesNoCancel, MessageBoxImage.Question);
                if (result == MessageBoxResult.Yes)
                {
                    //关闭所有线程
                    ShutdAllThread();
                    //开启一条新的线程
                    threadZZ = new Thread(AutoZZ);
                    threadZZ.Start();
                }
            }else
            {
                //已经创建,则关闭当前线程
                if (threadZZ != null)
                {
                    threadZZ.Abort();
                }
                //开启一条新的线程
                threadZZ = new Thread(AutoZZ);
                threadZZ.Start();
            }
        }




        //自动征战-线程函数
        private void AutoZZ()
        {
            //调用征战逻辑
            ZZAction();
        }




        //征战动作逻辑函数
        private void ZZAction()
        {
            //设置自动脚本标记，防止页面被定时器刷新
            isAuto = true;

            //------------通过点击垂直箭头，复位初始状态-----------------------
            //单击4次，垂直滚动条向上箭头
            SetCursorPos(1148, 35);
            mouse_event(MOUSEEVENTF_LEFTDOWN | MOUSEEVENTF_LEFTUP, 0, 0, 0, 0);
            mouse_event(MOUSEEVENTF_LEFTDOWN | MOUSEEVENTF_LEFTUP, 0, 0, 0, 0);
            mouse_event(MOUSEEVENTF_LEFTDOWN | MOUSEEVENTF_LEFTUP, 0, 0, 0, 0);
            mouse_event(MOUSEEVENTF_LEFTDOWN | MOUSEEVENTF_LEFTUP, 0, 0, 0, 0);


            //单击1次，垂直滚动条向下箭头
            SetCursorPos(1148, 678);
            //鼠标左键按下并弹起（单击一次）
            mouse_event(MOUSEEVENTF_LEFTDOWN | MOUSEEVENTF_LEFTUP, 0, 0, 0, 0);

            //等待2秒
            Thread.Sleep(2000);

            //单击，征战图标
            SetCursorPos(995, 669);
            mouse_event(MOUSEEVENTF_LEFTDOWN | MOUSEEVENTF_LEFTUP, 0, 0, 0, 0);

            //等待5秒
            Thread.Sleep(5000);
            
            //定义3个用于屏幕取色的变量
            Point p;            //屏幕中的一个点
            int hdc;            //设备上下文句柄
            int c;              //颜色变量
            bool isReZZ = false;        //重新征战标记

            //死循环，只有“征战按钮为灰色时终止循环”自动终止循环
            while (true)
            {
                //单击，确定。重置机会已用完，或者每通过10关卡获得一个礼包
                SetCursorPos(751, 475);
                mouse_event(MOUSEEVENTF_LEFTDOWN | MOUSEEVENTF_LEFTUP, 0, 0, 0, 0);

                Thread.Sleep(8000);      //等待8秒

                //-----------------------判断“进攻”按钮的颜色---------------------------
                hdc = GetWindowDC(0);            //获取设备上下文句柄(0是屏幕的设备上下文) 
                c = GetPixel(hdc, 635, 611);     //获取指定点的颜色
                ReleaseDC(0, hdc);
                //如果按钮是灰色
                if(c == 9276813)
                {
                    break;      //终止循环
                }

                //单击，进攻
                SetCursorPos(669, 611);
                mouse_event(MOUSEEVENTF_LEFTDOWN | MOUSEEVENTF_LEFTUP, 0, 0, 0, 0);

                Thread.Sleep(3000);         //等待3秒

                //单击，跳过
                SetCursorPos(1000, 675);
                mouse_event(MOUSEEVENTF_LEFTDOWN | MOUSEEVENTF_LEFTUP, 0, 0, 0, 0);

                Thread.Sleep(2000);         //等待2秒

                //--------------判断第一颗星星的颜色--------------
                hdc = GetWindowDC(0);
                c = GetPixel(hdc, 451, 298);
                ReleaseDC(0, hdc);
                //如果第一颗星是灰色的
                if (c == 8355711)
                {
                    //单击，确定。跳过后的确定
                    SetCursorPos(667, 576);
                    mouse_event(MOUSEEVENTF_LEFTDOWN | MOUSEEVENTF_LEFTUP, 0, 0, 0, 0);
                    //判断是否第二次征战
                    if (!isReZZ)
                    {
                        Thread.Sleep(12000);
                        //单击，重新征战
                        SetCursorPos(780, 562);
                        mouse_event(MOUSEEVENTF_LEFTDOWN | MOUSEEVENTF_LEFTUP, 0, 0, 0, 0);
                        //把第二次重新征战标识设置为真
                        isReZZ = true;
                    }else
                    {
                        break;         //已经第二次征战，并且该关卡失败，直接退出循环
                    }
                }else
                {
                    //单击，确定。跳过后的确定
                    SetCursorPos(667, 576);
                    mouse_event(MOUSEEVENTF_LEFTDOWN | MOUSEEVENTF_LEFTUP, 0, 0, 0, 0);
                }

                Thread.Sleep(6000);      //等待6秒
            }

            //关闭由于点击“重新征战”弹出的重置机会已用完提醒框，否则无法正常关闭征战窗口
            SetCursorPos(848, 250);
            mouse_event(MOUSEEVENTF_LEFTDOWN | MOUSEEVENTF_LEFTUP, 0, 0, 0, 0);

            //关闭征战窗口
            Thread.Sleep(1000); 
            SetCursorPos(982, 192);
            mouse_event(MOUSEEVENTF_LEFTDOWN | MOUSEEVENTF_LEFTUP, 0, 0, 0, 0);

            //关闭自动脚本标记，进行刷新
            isAuto = false;
        }




        //停止脚本-按钮点击事件响应函数
        private void stopAuto_Click(object sender, RoutedEventArgs e)
        {
            ShutdAllThread();
        }



        //关闭所有已开启的自动线程函数
        private void ShutdAllThread()
        {
            //关闭自动征战线程
            if (threadZZ != null)
            {
                threadZZ.Abort();   
                threadZZ = null;
            }
            //关闭自动收集线程
            if (threadGet != null)
            {
                threadGet.Abort();
                threadGet = null;
            }

            //关闭自动脚本标记，进行刷新
            isAuto = false;
        }




        //自动收集-按钮点击事件响应函数
        private void btn_autoGet_Click(object sender, RoutedEventArgs e)
        {
            if(isAuto)     //如果有其他动作线程在执行，先关闭掉,防止多个动作相互干扰
            {
                MessageBoxResult result = MessageBox.Show("正在执行其他自动操作，是否停止它！", "警告", MessageBoxButton.YesNoCancel, MessageBoxImage.Question);
                if (result == MessageBoxResult.Yes)
                {
                    ShutdAllThread();   //关闭所有线程
                    //开启一条新的线程,用于自动收集
                    threadGet = new Thread(AutoGet);
                    threadGet.Start();
                }
            }else
            {
                //判断线程是否已经存在
                if (threadGet != null)
                {
                    threadGet.Abort();  //关闭线程
                }
                //开启一条新的线程,用于自动收集
                threadGet = new Thread(AutoGet);
                threadGet.Start();
            }
        }




        //自动收集-线程函数
        private void AutoGet()
        {
            //调用自动收集逻辑函数
            GetAction();
        }




        //自动收集逻辑函数
        private void GetAction()
        {
            //设置自动脚本标记，防止刷新
            isAuto = true;

            //------------通过调整垂直滚动条，复位初始状态-----------------------
            //单击4次，垂直滚动条向上箭头
            SetCursorPos(1148, 35);
            mouse_event(MOUSEEVENTF_LEFTDOWN | MOUSEEVENTF_LEFTUP, 0, 0, 0, 0);
            mouse_event(MOUSEEVENTF_LEFTDOWN | MOUSEEVENTF_LEFTUP, 0, 0, 0, 0);
            mouse_event(MOUSEEVENTF_LEFTDOWN | MOUSEEVENTF_LEFTUP, 0, 0, 0, 0);
            mouse_event(MOUSEEVENTF_LEFTDOWN | MOUSEEVENTF_LEFTUP, 0, 0, 0, 0);
            //单击1次，垂直滚动条向下箭头
            SetCursorPos(1148, 678);
            mouse_event(MOUSEEVENTF_LEFTDOWN | MOUSEEVENTF_LEFTUP, 0, 0, 0, 0);


            //单击，英雄图标
            Thread.Sleep(2000);            //等待2秒
            SetCursorPos(464, 668);
            mouse_event(MOUSEEVENTF_LEFTDOWN | MOUSEEVENTF_LEFTUP, 0, 0, 0, 0);
            //------单击---- --小酌一杯------4次(偶数次，不会出现没有关闭)-----
            //等待2秒
            Thread.Sleep(2000);
            SetCursorPos(466, 503);
            mouse_event(MOUSEEVENTF_LEFTDOWN | MOUSEEVENTF_LEFTUP, 0, 0, 0, 0);
            //等待1秒
            Thread.Sleep(1000);
            SetCursorPos(466, 503);
            mouse_event(MOUSEEVENTF_LEFTDOWN | MOUSEEVENTF_LEFTUP, 0, 0, 0, 0);
            //等待1秒
            Thread.Sleep(1000);
            SetCursorPos(466, 503);
            mouse_event(MOUSEEVENTF_LEFTDOWN | MOUSEEVENTF_LEFTUP, 0, 0, 0, 0);
			//确定一次，防止第3次单击，出现获得英雄提醒框，没有关闭
            Thread.Sleep(1000);
            SetCursorPos(466, 503);
            mouse_event(MOUSEEVENTF_LEFTDOWN | MOUSEEVENTF_LEFTUP, 0, 0, 0, 0);
            //等待2秒
            Thread.Sleep(2000);
            //关闭勋章不足-对话框
            SetCursorPos(850, 348);
            mouse_event(MOUSEEVENTF_LEFTDOWN | MOUSEEVENTF_LEFTUP, 0, 0, 0, 0);
            //--------单击----------畅饮一日-------------1次------------------
            Thread.Sleep(1000);            //等待1秒
            SetCursorPos(673, 500);
            mouse_event(MOUSEEVENTF_LEFTDOWN | MOUSEEVENTF_LEFTUP, 0, 0, 0, 0);
            //点击确定--恭喜你获得了该英雄
            Thread.Sleep(2000);
            SetCursorPos(666, 507);
            mouse_event(MOUSEEVENTF_LEFTDOWN | MOUSEEVENTF_LEFTUP, 0, 0, 0, 0);
            //等待2秒
            Thread.Sleep(2000);
            //关闭勋章不足-对话框
            SetCursorPos(850, 348);
            mouse_event(MOUSEEVENTF_LEFTDOWN | MOUSEEVENTF_LEFTUP, 0, 0, 0, 0);
            //--------单击----------豪饮3日-------------1次------------------
            Thread.Sleep(1000);            //等待1秒
            SetCursorPos(872, 500);
            mouse_event(MOUSEEVENTF_LEFTDOWN | MOUSEEVENTF_LEFTUP, 0, 0, 0, 0);
            //等待2秒
            Thread.Sleep(2000);
            //点击确定--恭喜你获得了该英雄
            SetCursorPos(892, 508);
            mouse_event(MOUSEEVENTF_LEFTDOWN | MOUSEEVENTF_LEFTUP, 0, 0, 0, 0);
            //等待2秒
            Thread.Sleep(2000);
            //关闭勋章不足-对话框
            SetCursorPos(850, 348);
            mouse_event(MOUSEEVENTF_LEFTDOWN | MOUSEEVENTF_LEFTUP, 0, 0, 0, 0);
            //--------单击----------关闭英雄对话框-------------1次------------------
            Thread.Sleep(1000);            //等待1秒
            SetCursorPos(996, 170);
            mouse_event(MOUSEEVENTF_LEFTDOWN | MOUSEEVENTF_LEFTUP, 0, 0, 0, 0);


            //单击，国家图标
            Thread.Sleep(2000);            //等待2秒
            SetCursorPos(568, 676);
            mouse_event(MOUSEEVENTF_LEFTDOWN | MOUSEEVENTF_LEFTUP, 0, 0, 0, 0);
            //--------单击----------国家宝箱-------------1次------------------
            Thread.Sleep(2000);            //等待2秒
            SetCursorPos(699, 405);
            mouse_event(MOUSEEVENTF_LEFTDOWN | MOUSEEVENTF_LEFTUP, 0, 0, 0, 0);
            //--------单击----------领取-------------1次------------------
            Thread.Sleep(5000);            //等待5秒
            SetCursorPos(558, 350);
            mouse_event(MOUSEEVENTF_LEFTDOWN | MOUSEEVENTF_LEFTUP, 0, 0, 0, 0);
            //--------单击----------关闭国家宝箱对话框-------------1次------------------
            Thread.Sleep(1000);            //等待1秒
            SetCursorPos(932, 170);
            mouse_event(MOUSEEVENTF_LEFTDOWN | MOUSEEVENTF_LEFTUP, 0, 0, 0, 0);
            //--------单击----------关闭国家对话框-------------1次------------------
            Thread.Sleep(1000);            //等待1秒
            SetCursorPos(985, 186);
            mouse_event(MOUSEEVENTF_LEFTDOWN | MOUSEEVENTF_LEFTUP, 0, 0, 0, 0);


            //单击，将领图标
            Thread.Sleep(2000);            //等待2秒
            SetCursorPos(680, 670);
            mouse_event(MOUSEEVENTF_LEFTDOWN | MOUSEEVENTF_LEFTUP, 0, 0, 0, 0);
            //--------单击----------低级冶炼-------------3次------------------
            Thread.Sleep(3000);            //等待3秒
            SetCursorPos(464, 493);
            mouse_event(MOUSEEVENTF_LEFTDOWN | MOUSEEVENTF_LEFTUP, 0, 0, 0, 0);
            //确定
            Thread.Sleep(3000);            //等待3秒
            SetCursorPos(453, 497);
            mouse_event(MOUSEEVENTF_LEFTDOWN | MOUSEEVENTF_LEFTUP, 0, 0, 0, 0);
            //第二次
            Thread.Sleep(3000);            //等待3秒
            SetCursorPos(464, 493);
            mouse_event(MOUSEEVENTF_LEFTDOWN | MOUSEEVENTF_LEFTUP, 0, 0, 0, 0);
            //确定
            Thread.Sleep(3000);            //等待3秒
            SetCursorPos(453, 497);
            mouse_event(MOUSEEVENTF_LEFTDOWN | MOUSEEVENTF_LEFTUP, 0, 0, 0, 0);
            //第三次
            Thread.Sleep(3000);            //等待3秒
            SetCursorPos(464, 493);
            mouse_event(MOUSEEVENTF_LEFTDOWN | MOUSEEVENTF_LEFTUP, 0, 0, 0, 0);
            //确定
            Thread.Sleep(3000);            //等待3秒
            SetCursorPos(453, 497);
            mouse_event(MOUSEEVENTF_LEFTDOWN | MOUSEEVENTF_LEFTUP, 0, 0, 0, 0);
            //关闭--勋章不足
            Thread.Sleep(1000);            //等待3秒
            SetCursorPos(848, 349);
            mouse_event(MOUSEEVENTF_LEFTDOWN | MOUSEEVENTF_LEFTUP, 0, 0, 0, 0);
            //--------单击----------中级冶炼-------------3次------------------
            Thread.Sleep(2000);            //等待2秒
            SetCursorPos(662, 490);
            mouse_event(MOUSEEVENTF_LEFTDOWN | MOUSEEVENTF_LEFTUP, 0, 0, 0, 0);
            Thread.Sleep(1000);            //等待1秒
            //确定
            SetCursorPos(662, 493);
            mouse_event(MOUSEEVENTF_LEFTDOWN | MOUSEEVENTF_LEFTUP, 0, 0, 0, 0);
            //关闭--勋章不足
            Thread.Sleep(1000);            //等待3秒
            SetCursorPos(848, 349);
            mouse_event(MOUSEEVENTF_LEFTDOWN | MOUSEEVENTF_LEFTUP, 0, 0, 0, 0);
            //--------单击----------高级冶炼-------------3次------------------
            Thread.Sleep(2000);            //等待2秒
            SetCursorPos(872, 493);
            mouse_event(MOUSEEVENTF_LEFTDOWN | MOUSEEVENTF_LEFTUP, 0, 0, 0, 0);
            Thread.Sleep(1000);            //等待1秒
            //确定
            SetCursorPos(872, 493);
            mouse_event(MOUSEEVENTF_LEFTDOWN | MOUSEEVENTF_LEFTUP, 0, 0, 0, 0);
            //关闭--勋章不足
            Thread.Sleep(1000);            //等待1秒
            SetCursorPos(848, 349);
            mouse_event(MOUSEEVENTF_LEFTDOWN | MOUSEEVENTF_LEFTUP, 0, 0, 0, 0);
            //-------------------------关闭--将领对话框-----------------------
            Thread.Sleep(2000);            //等待2秒
            SetCursorPos(990, 176);
            mouse_event(MOUSEEVENTF_LEFTDOWN | MOUSEEVENTF_LEFTUP, 0, 0, 0, 0);


            //单击，武器中心图标
            Thread.Sleep(3000);            //等待3秒
            SetCursorPos(886, 673);
            mouse_event(MOUSEEVENTF_LEFTDOWN | MOUSEEVENTF_LEFTUP, 0, 0, 0, 0);
            //--------单击----------野外-------------3次------------------
            Thread.Sleep(6000);            //等待8秒
            SetCursorPos(519, 263);
            mouse_event(MOUSEEVENTF_LEFTDOWN | MOUSEEVENTF_LEFTUP, 0, 0, 0, 0);
            Thread.Sleep(3000);            //等待3秒
            SetCursorPos(585, 516);
            mouse_event(MOUSEEVENTF_LEFTDOWN | MOUSEEVENTF_LEFTUP, 0, 0, 0, 0);
            Thread.Sleep(1000);            //等待1秒
            SetCursorPos(585, 516);
            mouse_event(MOUSEEVENTF_LEFTDOWN | MOUSEEVENTF_LEFTUP, 0, 0, 0, 0);
            Thread.Sleep(1000);            //等待1秒
            SetCursorPos(585, 516);
            mouse_event(MOUSEEVENTF_LEFTDOWN | MOUSEEVENTF_LEFTUP, 0, 0, 0, 0);
            //单击确定
            Thread.Sleep(1000);            //等待1秒
            SetCursorPos(761, 469);
            mouse_event(MOUSEEVENTF_LEFTDOWN | MOUSEEVENTF_LEFTUP, 0, 0, 0, 0);
            //--------单击-----------工厂--------------1次--------------------
            Thread.Sleep(2000);            //等待2秒
            SetCursorPos(592, 263);
            mouse_event(MOUSEEVENTF_LEFTDOWN | MOUSEEVENTF_LEFTUP, 0, 0, 0, 0);
            //点击--免费探索
            Thread.Sleep(2000);            //等待2秒
            SetCursorPos(453, 514);
            mouse_event(MOUSEEVENTF_LEFTDOWN | MOUSEEVENTF_LEFTUP, 0, 0, 0, 0);
            //关闭--探索卡已用完
            Thread.Sleep(2000);            //等待2秒
            SetCursorPos(847, 365);
            mouse_event(MOUSEEVENTF_LEFTDOWN | MOUSEEVENTF_LEFTUP, 0, 0, 0, 0);
            //--------单击-----------实验室--------------1次--------------------
            Thread.Sleep(2000);            //等待2秒
            SetCursorPos(662, 263);
            mouse_event(MOUSEEVENTF_LEFTDOWN | MOUSEEVENTF_LEFTUP, 0, 0, 0, 0);
            //点击--免费探索
            Thread.Sleep(2000);            //等待2秒
            SetCursorPos(456, 516);
            mouse_event(MOUSEEVENTF_LEFTDOWN | MOUSEEVENTF_LEFTUP, 0, 0, 0, 0);
            //关闭--探索卡已用完
            Thread.Sleep(2000);            //等待2秒
            SetCursorPos(852, 366);
            mouse_event(MOUSEEVENTF_LEFTDOWN | MOUSEEVENTF_LEFTUP, 0, 0, 0, 0);
            //关闭--武器中心
            Thread.Sleep(2000);            //等待2秒
            SetCursorPos(991, 188);
            mouse_event(MOUSEEVENTF_LEFTDOWN | MOUSEEVENTF_LEFTUP, 0, 0, 0, 0);


            //单击，战役图标
            Thread.Sleep(2000);            //等待2秒
            SetCursorPos(1044, 673);
            mouse_event(MOUSEEVENTF_LEFTDOWN | MOUSEEVENTF_LEFTUP, 0, 0, 0, 0);
            //--------单击-----------战场--------------1次--------------------
            Thread.Sleep(5000);            //等待5秒
            SetCursorPos(728, 362);
            mouse_event(MOUSEEVENTF_LEFTDOWN | MOUSEEVENTF_LEFTUP, 0, 0, 0, 0);
            //--------单击-----------扫荡--------------3次--------------------
            Thread.Sleep(3000);            //等待3秒
            SetCursorPos(756, 496);
            mouse_event(MOUSEEVENTF_LEFTDOWN | MOUSEEVENTF_LEFTUP, 0, 0, 0, 0);
            Thread.Sleep(1000);            //等待1秒
            SetCursorPos(756, 496);
            mouse_event(MOUSEEVENTF_LEFTDOWN | MOUSEEVENTF_LEFTUP, 0, 0, 0, 0);
            Thread.Sleep(1000);            //等待1秒
            SetCursorPos(756, 496);
            mouse_event(MOUSEEVENTF_LEFTDOWN | MOUSEEVENTF_LEFTUP, 0, 0, 0, 0);
            Thread.Sleep(1000);            //等待1秒
            SetCursorPos(864, 312);        //关闭消耗勋章获得行军令
            mouse_event(MOUSEEVENTF_LEFTDOWN | MOUSEEVENTF_LEFTUP, 0, 0, 0, 0);
            //关闭--扫荡
            Thread.Sleep(1000);            //等待1秒
            SetCursorPos(852, 277);
            mouse_event(MOUSEEVENTF_LEFTDOWN | MOUSEEVENTF_LEFTUP, 0, 0, 0, 0);
            //返回基地
            Thread.Sleep(1000);            //等待1秒
            SetCursorPos(1115, 677);
            mouse_event(MOUSEEVENTF_LEFTDOWN | MOUSEEVENTF_LEFTUP, 0, 0, 0, 0);


            //关闭自动脚本标记，进行刷新
            isAuto = false;
        }
    }
}