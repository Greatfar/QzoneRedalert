using System;
using System.Windows;
using System.Windows.Navigation;
using System.Configuration;
using System.Runtime.InteropServices;       //清空session时用到

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


        //------------------------清空session时用到------------------------        private const int INTERNET_OPTION_END_BROWSER_SESSION = 42;
        [DllImport("wininet.dll", SetLastError = true)]
        private static extern bool InternetSetOption(IntPtr hInternet, int dwOption, IntPtr lpBuffer, int lpdwBufferLength);
        //-----------------------------------------------------------------




    /// <summary>
    /// ========================================主窗体函数=================================================
    /// </summary>
    public MainWindow()
        {
            InitializeComponent();      //初始化组件

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

            string myStatus = ConfigurationManager.AppSettings["Status"];           //从配置文件读取登录状态，Status

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

            //绑定文档加载完成的回调函数，到web控件中,onLoadDocCompleted
            web1.LoadCompleted += new LoadCompletedEventHandler(onLoadDocCompleted);

        }




        //定时器回调函数
        private void timer_Refresh(object sender, EventArgs e)
        {
            //判断登录状态，如果没有登录就不执行刷新
            if (isLogin)
            { 
                Uri uri = new Uri(game_url);        //web控件，加载上一次登录的URL
                web1.Navigate(uri);
            }
        }




        //网页加载完成，回调函数
        private void onLoadDocCompleted(object sender, NavigationEventArgs e)
        {
            //获取文档对象
            mshtml.IHTMLDocument2 doc2 = (mshtml.IHTMLDocument2)web1.Document;

            //判断：登录状态、刷新状态。登录状态通过网页标题判断，刷新状态通过全局变量判断。
            if (doc2.title == "红警之坦克风暴 - 应用中心" && isFresh == false) {
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


    }
}
