using System;
using System.Windows;
using System.Windows.Navigation;
using System.Configuration;
using System.Runtime.InteropServices;       //清空session时用到
using Forms = System.Windows.Forms;         //最小化到系统托盘
using System.Threading;
using System.Windows.Threading;
using System.Windows.Controls;

namespace RedAlert
{
    /// <summary>
    /// MainWindow.xaml 的交互逻辑
    /// </summary>
    public partial class MainWindow : Window
    {
        string game_url;                                        //免登陆URL
        bool isLogin = false;                                   //登录状态
        System.Windows.Threading.DispatcherTimer timer;         //定时器变量
        bool isFresh = false;                                   //刷新状态标识
        bool isAuto = false;                                    //自动脚本标记
        Thread threadZZ = null;                                 //自动征战线程
        Thread threadGet = null;                                //自动收集线程
        Thread threadLogin = null;                              //登录线程
        string str_qq = "";                                     //用户名
        string str_pwd = "";                                    //密码
        int accNum = 0;                                         //配置文件中保存的QQ账号个数
        int dfQQ = 0;                                           //默认登录的QQ序号
        int lastAccNum = 0;                                     //账号溢出指针

        private Forms.NotifyIcon notifyIcon;                    //系统托盘图标变量

        public delegate void delegate_login();                  //定义委托函数


        //----------------模拟鼠标事件常量--------------------
        const int MOUSEEVENTF_MOVE = 0x0001;        //移动鼠标
        const int MOUSEEVENTF_LEFTDOWN = 0x0002;    //左键按下
        const int MOUSEEVENTF_LEFTUP = 0x0004;      //左键抬起
        const int MOUSEEVENTF_RIGHTDOWN = 0x0008;   //右键按下
        const int MOUSEEVENTF_RIGHTUP = 0x0010;     //右键抬起
        const int MOUSEEVENTF_MIDDLEDOWN = 0x0020;  //中键按下
        const int MOUSEEVENTF_MIDDLEUP = 0x0040;    //中键抬起
        const int MOUSEEVENTF_ABSOLUTE = 0x8000;    //采用绝对坐标


        //-------------------------导入模拟鼠标事件的win32 api函数----------------------------------
        //鼠标事件函数
        [System.Runtime.InteropServices.DllImport("user32")]
        private static extern int mouse_event(int dwFlags, int dx, int dy, int cButtons, int dwExtraInfo);
        //设置光标位置函数
        [DllImport("user32.dll")]
        static extern bool SetCursorPos(int X, int Y);


        //-------------导入屏幕取色相关win32 api函数-------------------
        [DllImport("user32")]
        private static extern int GetWindowDC(int hwnd);        //获取窗口句柄函数
        [DllImport("user32")]
        private static extern int ReleaseDC(int hWnd, int hDC); //释放窗口句柄函数
        [DllImport("gdi32")]
        private static extern int GetPixel(int hdc, int nXPos, int nYPos);  //获取指定点的颜色函数

        
        //----------------------------------导入清空session用到的windows api函数------------------------------------------
        private const int INTERNET_OPTION_END_BROWSER_SESSION = 42;
        [DllImport("wininet.dll", SetLastError = true)]
        private static extern bool InternetSetOption(IntPtr hInternet, int dwOption, IntPtr lpBuffer, int lpdwBufferLength);




    /// <summary>
    /// =========主窗体类构造函数==================
    /// </summary>
    public MainWindow()
        {
            InitializeComponent();  //初始化组件（系统默认调用）

            InitDisplay();          //初始化显示状态

            InitRefresh();          //刷新初始化

            InitAccount();          //初始化账号显示

            CheckLoginStatus();     //检查登录状态
        }




        //显示初始化函数
        private void InitDisplay()
        {
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


            //---------------------------指定窗口显示位置------------------------------------
            WindowStartupLocation = WindowStartupLocation.Manual;       //自定义窗口显示位置
            this.Top = 0;                       //距离顶部
            this.Left = 188;                    //距离左边


            //绑定web控件的文档加载完成的回调函数，到web控件中,onLoadDocCompleted
            web1.LoadCompleted += new LoadCompletedEventHandler(onLoadDocCompleted);
        }




        //刷新初始化函数
        private void InitRefresh()
        {
            //读取刷新周期
            string strFreshT = ConfigurationManager.AppSettings["FreshT"];
            int freshT = Convert.ToInt32(strFreshT);

            //生成定时器对象，并设置好参数
            timer = new System.Windows.Threading.DispatcherTimer();
            timer.Tick += new EventHandler(timer_Refresh);      //绑定时间到达时的回调函数
            timer.Interval = new TimeSpan(0, freshT, 0);        //设置定时器时钟：TimeSpan（时, 分， 秒）。freshT分钟刷新一次
        }




        //账号初始化函数
        private void InitAccount()
        {
            //初始化账号、密码输入框
            string strNum = ConfigurationManager.AppSettings["Default_QQ"];
            dfQQ = Convert.ToInt16(strNum);
            string df_qq = "QQ" + dfQQ.ToString();
            string df_mm = "MM" + dfQQ.ToString();
            str_qq = qq_textbox.Text = ConfigurationManager.AppSettings[df_qq];
            str_pwd = passwordBox.Password = ConfigurationManager.AppSettings[df_mm];

            //获取账号个数
            string straccNum = ConfigurationManager.AppSettings["Account_Num"];
            accNum = Convert.ToInt16(straccNum);

            //初始化账号组合框
            for (int i = 0; i < accNum; i++)
            {
                int j = i + 1;
                string cfgqq = "QQ" + j.ToString();
                string qq = ConfigurationManager.AppSettings[cfgqq];
                ComboBoxItem qq_item = new ComboBoxItem();
                qq_item.Content = qq;
                comboBox.Items.Add(qq_item);
            }
            comboBox.SelectedIndex = dfQQ - 1;     //设置组合框默认显示项

            lastAccNum = Convert.ToInt16(ConfigurationManager.AppSettings["Last_accNum"]);  //读取上一次账号溢出值
            if (lastAccNum == 10) { lastAccNum = 0; }                                       //重置溢出值
        }




        //检查登录状态函数
        private void CheckLoginStatus()
        {
            //打开配置文件
            Configuration cfa = ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None);

            //获取当前时间的以100“毫微秒”为单位的值。
            long currentTime = System.DateTime.Now.Ticks;

            //读取配置文件中的时间
            string strLastTime = ConfigurationManager.AppSettings["Date"];
            long lastTime = long.Parse(strLastTime);

            //从配置文件读取登录状态
            string myStatus = ConfigurationManager.AppSettings["Status"];

            //判断是否第一次登录
            if (myStatus == "secondtime")
            {
                //10小时内使用免登陆URL
                if (((currentTime - lastTime) / 10000000) < 36000)
                {
                    game_url = ConfigurationManager.AppSettings["Game"];            //从配置配置文件读取免登陆URL
                    Uri uri = new Uri(game_url);                                    //URL转换为URI
                    web1.Navigate(uri);                                             //web控件载入URL
                    timer.Start();                                                  //启动定时器
                }
                else    //10小时后，免登陆URL失效
                {
                    cfa.AppSettings.Settings["Date"].Value = currentTime.ToString();    //把当前时间写入配置文件
                    GameLogin();      //调用登录函数，进行重新登录
                }
            }
            else    //第一次登录
            {
                //打开QQ空间登录跳转页面
                Uri uri = new Uri("http://i.qq.com/?s_url=http%3A%2F%2Fmy.qzone.qq.com%2Fapp%2F100616028.html#via=appcenter.info");
                web1.Navigate(uri);
                //把当前时间写入配置文件
                cfa.AppSettings.Settings["Date"].Value = currentTime.ToString();
            }
            cfa.Save();     //保存配置文件
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
                    cfa.AppSettings.Settings["Status"].Value = "secondtime";         //写入第二次登录标记
                    cfa.Save();
                    //重新载入免登陆URL。目的：去除腾讯QQ空间的外部框架，去除广告。
                    Uri uri = new Uri(game_url);
                    web1.Navigate(uri);
                    //启动定时器，用于刷新计时
                    timer.Start();
                    //把刷新状态改为真，进入刷新状态。网页加载完成不再执行提取操作。
                    isFresh = true;
                }
            }
        }




        //复选框选中时自动调用，（在xaml中绑定了该函数）
        private void CheckBox_Checked(object sender, RoutedEventArgs e)
        {
            isLogin = true;         //登录状态置为真。放行刷新动作。
        }




        //复选框取消选中时自动调用，（在xaml中绑定了该函数）
        private void CheckBox_Unchecked(object sender, RoutedEventArgs e)
        {
            isLogin = false;        //登录状态置为假。组织刷新动作
        }




        //退出按钮，点击事件回调函数
        private void button_Click(object sender, RoutedEventArgs e)
        {
            //改变配置文件文件中的登录状态、免登陆URL。达到再次启动应用时自动打开登录页面
            Configuration cfa = ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None);
            cfa.AppSettings.Settings["Game"].Value = "reLogin"; 
            cfa.AppSettings.Settings["Status"].Value = "exit"; 
            cfa.Save();

            //清空session
            InternetSetOption(IntPtr.Zero, INTERNET_OPTION_END_BROWSER_SESSION, IntPtr.Zero, 0);

            //跳转到登录页面
            Uri uri = new Uri("http://qqapp.qq.com/app/100616028.html");
            web1.Navigate(uri);

            //把刷新状态置为假。让文档加载完成时的提取免登陆url。
            isFresh = false;
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




        //重写关闭按钮回调方法。（在xaml中绑定了该函数）
        private void MainWindow_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            MessageBoxResult result = MessageBox.Show("是否最小化到系统托盘，后台运行？", "关闭选项", MessageBoxButton.YesNoCancel, MessageBoxImage.Question);
            if (result == MessageBoxResult.Yes)
            {
                Hide();             //调用隐藏窗口方法
                e.Cancel = true;    //并向窗口传递取消关闭消息
            }
            else if (result == MessageBoxResult.No)
            {
                e.Cancel = false;   //取消关闭，置为假
            }
            else
            {
                e.Cancel = true;    //取消关闭窗口
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
                //单击，确定。每通过10关卡获得一个礼包
                SetCursorPos(751, 475);
                mouse_event(MOUSEEVENTF_LEFTDOWN | MOUSEEVENTF_LEFTUP, 0, 0, 0, 0);

                Thread.Sleep(1000);      //等待1秒

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

                Thread.Sleep(4000);         //等待3秒

                //单击，跳过
                SetCursorPos(1000, 675);
                mouse_event(MOUSEEVENTF_LEFTDOWN | MOUSEEVENTF_LEFTUP, 0, 0, 0, 0);

                Thread.Sleep(3000);         //等待2秒

                //--------------判断第一颗星星的颜色--------------
                hdc = GetWindowDC(0);
                c = GetPixel(hdc, 451, 298);
                ReleaseDC(0, hdc);
                //如果第一颗星是灰色的
                if (c == 8355711)
                {
                    //单击，确定
                    SetCursorPos(667, 576);
                    mouse_event(MOUSEEVENTF_LEFTDOWN | MOUSEEVENTF_LEFTUP, 0, 0, 0, 0);
                    Thread.Sleep(13000);    //等待13秒
                    //判断是否第二次征战
                    if (!isReZZ)
                    {
                        //单击，重新征战
                        SetCursorPos(778, 613);
                        mouse_event(MOUSEEVENTF_LEFTDOWN | MOUSEEVENTF_LEFTUP, 0, 0, 0, 0);
                        Thread.Sleep(2000);
                        //单击，确定
                        SetCursorPos(562, 420);
                        mouse_event(MOUSEEVENTF_LEFTDOWN | MOUSEEVENTF_LEFTUP, 0, 0, 0, 0);
                        Thread.Sleep(1000);
                        //--------------判断是否已经用完重置机会--------------
                        hdc = GetWindowDC(0);
                        c = GetPixel(hdc, 590, 392);
                        ReleaseDC(0, hdc);
                        if(c == 988233)     //红色提示框
                        {
                            //关闭-勋章不足
                            SetCursorPos(848, 380);
                            mouse_event(MOUSEEVENTF_LEFTDOWN | MOUSEEVENTF_LEFTUP, 0, 0, 0, 0);
                            break;      //退出循环
                        }
                        //把第二次重新征战标识设置为真
                        isReZZ = true;
                    }else
                    {
                        break;                  //已经第二次征战，并且该关卡失败，直接退出循环
                    }
                }else
                {
                    //单击，确定。
                    SetCursorPos(667, 576);
                    mouse_event(MOUSEEVENTF_LEFTDOWN | MOUSEEVENTF_LEFTUP, 0, 0, 0, 0);
                }

                Thread.Sleep(15000);
            }

            //关闭征战窗口
            Thread.Sleep(1000); 
            SetCursorPos(982, 192);
            mouse_event(MOUSEEVENTF_LEFTDOWN | MOUSEEVENTF_LEFTUP, 0, 0, 0, 0);

            //关闭自动脚本标记，进行刷新
            isAuto = false;
        }




        //停止动作按钮，点击事件回调函数
        private void stopAuto_Click(object sender, RoutedEventArgs e)
        {
            ShutdAllThread();
        }




        //关闭所有自动线程函数
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
            isAuto = false;     //自动标识置为假，放行刷新。
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




        //登录按钮-点击事件回调函数
        private void login_btn_Click(object sender, RoutedEventArgs e)
        {
            SavePassport();         //保存用户名、密码
            isFresh = false;        //把刷新状态改为未刷新，保证登录后，自动获取免登陆URL
            GameLogin();            //调用登录函数
        }




        //保存账号、密码函数
        private void SavePassport()
        {
            //更新登录用的QQ、密码
            str_qq = qq_textbox.Text;
            str_pwd = passwordBox.Password;

            Configuration cfa = ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None); //生成配置文件对象
            bool isExist = false;           //账号是否存在标识
            int j = 0;                      //QQ在配置中的位置
            string cfg_qq, cfg_mm;          //账号、密码定位变量（配置文件）
            int acc_pointer;                //账号指针

            //遍历配置文件，判断账号是否已经存在
            for (int i = 0; i < accNum; i++)
            {
                j = i + 1;
                string the_qq = "QQ" + j.ToString();
                string cfgQQ = ConfigurationManager.AppSettings[the_qq];
                if (cfgQQ == qq_textbox.Text)
                {
                    isExist = true;
                    break;
                }
            }

            //账号已存在
            if (isExist)
            {
                //账号指针指向账号所在位置
                cfg_qq = "QQ" + j.ToString();
                cfg_mm = "MM" + j.ToString();
                dfQQ = j;   //默认QQ
            }else
            {
                accNum++;           //账号数量+1
                dfQQ = accNum;      //默认QQ
                if(accNum > 10)     //配置文件中只能存10个账号
                {
                    accNum = 10;
                    acc_pointer = lastAccNum += 1;        //指向上一次溢出账号的下一个，将之覆盖
                }else
                {
                    acc_pointer = accNum;
                }
                //账号指针指向新的地方
                cfg_qq = "QQ" + acc_pointer.ToString();
                cfg_mm = "MM" + acc_pointer.ToString();
                cfa.AppSettings.Settings["Account_Num"].Value = accNum.ToString();  //更新个账号个数
            }
            
            //账号、密码写入配置文件
            cfa.AppSettings.Settings[cfg_qq].Value = str_qq;
            cfa.AppSettings.Settings[cfg_mm].Value = str_pwd;
            //更新默认账号、账号覆盖指针
            cfa.AppSettings.Settings["Default_QQ"].Value = dfQQ.ToString();
            cfa.AppSettings.Settings["Last_accNum"].Value = lastAccNum.ToString();
            cfa.Save();
        }




        //登录函数
        private void GameLogin()
        {
            //清空session
            InternetSetOption(IntPtr.Zero, INTERNET_OPTION_END_BROWSER_SESSION, IntPtr.Zero, 0);
            //关闭之前的登录线程
            if (threadLogin != null) { threadLogin.Abort(); }
            //启动登录线程
            threadLogin = new Thread(login_acttion);
            threadLogin.Start();
        }




        //登录线程函数
        private void login_acttion()
        {
            //---------------------------使用委托函数操作父线程的控件-----------------------------------
            //打开登录页面
            Dispatcher.BeginInvoke(DispatcherPriority.Normal, new delegate_login(OpenQzone));
            //改变网页框架
            Thread.Sleep(2000);
            Dispatcher.BeginInvoke(DispatcherPriority.Normal, new delegate_login(SelectLoginiFrame));
            //切换--账号密码登录方式
            Thread.Sleep(2000);
            Dispatcher.BeginInvoke(DispatcherPriority.Normal, new delegate_login(SwitchToPlogin));
            //输入QQ账号
            Thread.Sleep(1000);
            Dispatcher.BeginInvoke(DispatcherPriority.Normal, new delegate_login(InputQQ));
            //触发QQ输入框的失去焦点js函数
            Thread.Sleep(1000);
            SetCursorPos(675, 343);
            mouse_event(MOUSEEVENTF_LEFTDOWN | MOUSEEVENTF_LEFTUP, 0, 0, 0, 0);
            Thread.Sleep(1000);
            SetCursorPos(675, 396);
            mouse_event(MOUSEEVENTF_LEFTDOWN | MOUSEEVENTF_LEFTUP, 0, 0, 0, 0);
            Thread.Sleep(1000);
            //输入密码并提交
            Dispatcher.BeginInvoke(DispatcherPriority.Normal, new delegate_login(InputPassword));
            Thread.Sleep(1000);
            //提交表单
            Dispatcher.BeginInvoke(DispatcherPriority.Normal, new delegate_login(Submit));
        }




        //切换账号按钮--点击事件回调函数
        private void change_btn_Click(object sender, RoutedEventArgs e)
        {
            int cbi = comboBox.SelectedIndex;       //获取组合框选中项的索引值
            cbi++;                                  //转化为QQ在配置文件中的序号
            string qq_index = "QQ" + cbi.ToString();
            string mm_index = "MM" + cbi.ToString();

            //更新登录用的账号、密码
            str_qq = ConfigurationManager.AppSettings[qq_index];
            str_pwd = ConfigurationManager.AppSettings[mm_index];

            //更新默认QQ
            Configuration cfa = ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None);
            cfa.AppSettings.Settings["Default_QQ"].Value = cbi.ToString();
            cfa.Save();
            
            isFresh = false;    //把刷新置为假，放行web控件的网页加载完成后，提取免登陆URL
            GameLogin();        //调用登录函数
        }




        //委托函数--转到框架
        private void OpenQzone()
        {
            Uri uri = new Uri("http://i.qq.com/?s_url=http%3A%2F%2Fmy.qzone.qq.com%2Fapp%2F100616028.html#via=appcenter.info");
            web1.Navigate(uri);
        }
        //委托函数--转到框架
        private void SelectLoginiFrame()
        {
            mshtml.IHTMLDocument2 doc = (mshtml.IHTMLDocument2)web1.Document;
            mshtml.IHTMLElement l_iframe = (mshtml.IHTMLElement)doc.all.item("login_frame", 0);
            string zoneurl = (string)l_iframe.getAttribute("src");
            Uri li_uri = new Uri(zoneurl);
            web1.Navigate(li_uri);
        }
        //委托函数--切换登录方式
        private void SwitchToPlogin()
        {
            mshtml.IHTMLDocument2 doc = (mshtml.IHTMLDocument2)web1.Document;
            mshtml.IHTMLElement plogin = (mshtml.IHTMLElement)doc.all.item("switcher_plogin", 0);
            plogin.click();
        }
        //委托函数--输入QQ号码
        private void InputQQ()
        {
            mshtml.IHTMLDocument2 doc = (mshtml.IHTMLDocument2)web1.Document;
            mshtml.IHTMLElement usr = (mshtml.IHTMLElement)doc.all.item("u", 0);
            usr.setAttribute("value", str_qq);
        }
        //委托函数--输入密码并点击登录
        private void InputPassword()
        {
            mshtml.IHTMLDocument2 doc = (mshtml.IHTMLDocument2)web1.Document;
            mshtml.IHTMLElement password = (mshtml.IHTMLElement)doc.all.item("p", 0);
            password.setAttribute("value", str_pwd);
        }
        //委托函数--提交表单
        private void Submit()
        {
            mshtml.IHTMLDocument2 doc = (mshtml.IHTMLDocument2)web1.Document;
            mshtml.IHTMLElement submit = (mshtml.IHTMLElement)doc.all.item("login_button", 0);
            submit.click();
        }




        //刷新按钮-点击事件回调函数
        private void Refresh_btn_Click(object sender, RoutedEventArgs e)
        {
            //载入免登陆URL
            Uri uri = new Uri(game_url);
            web1.Navigate(uri);
        }
    }
}