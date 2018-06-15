using mshtml;
using SHDocVw;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;

namespace WindowsFormsApplication1
{
    public partial class Form1 : Form
    {
        [DllImport("user32.dll")]
        private static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);
        //设置前台窗口API（SetForegroundWindow）
        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool SetForegroundWindow(IntPtr hWnd);
        [DllImport("user32.dll")]
        private static extern IntPtr FindWindow(string lpClassName, string lpWindowName);
        [DllImport("user32.dll")]
        private static extern int SendMessage(IntPtr hWnd, int Msg, int wParam, int lParam);
        [DllImport("user32.dll")]
        private static extern bool SetCursorPos(int X, int Y);
        [DllImport("user32.dll")]
        private static extern void mouse_event(int dwFlags, int dx, int dy, int dwData, int dwExtraInfo);
        [DllImport("user32.dll")]
        private static extern void keybd_event(byte bVk, byte bScan, uint dwFlags, uint dwExtraInfo);
        [DllImport("user32.dll")]
        private static extern bool SetWindowPos(IntPtr hWnd, IntPtr hWndlnsertAfter, int X, int Y, int cx, int cy, uint Flags);
        //ShowWindow参数
        private const int SW_SHOWNORMAL = 1;
        private const int SW_RESTORE = 9;
        private const int SW_SHOWNOACTIVATE = 4;
        //SendMessage参数
        private const int WM_KEYDOWN = 0X100;
        private const int WM_KEYUP = 0X101;
        private const int WM_SYSCHAR = 0X106;
        private const int WM_SYSKEYUP = 0X105;
        private const int WM_SYSKEYDOWN = 0X104;
        private const int WM_CHAR = 0X102;
        //2)、遍历所有窗口得到句柄
        //1 定义委托方法CallBack，枚举窗口API（EnumWindows），得到窗口名API（GetWindowTextW）和得到窗口类名API（GetClassNameW）
        public delegate bool CallBack(int hwnd, int lParam);
        [DllImport("user32")]
        public static extern int EnumWindows(CallBack x, int y);
        [DllImport("user32.dll")]
        private static extern int GetWindowTextW(IntPtr hWnd, [MarshalAs(UnmanagedType.LPWStr)]StringBuilder lpString, int nMaxCount);
        [DllImport("user32.dll")]
        private static extern int GetClassNameW(IntPtr hWnd, [MarshalAs(UnmanagedType.LPWStr)]StringBuilder lpString, int nMaxCount);

        //3)、打开窗口得到句柄
        //1 定义设置活动窗口API（SetActiveWindow），设置前台窗口API（SetForegroundWindow）
        [DllImport("user32.dll")]
        static extern IntPtr SetActiveWindow(IntPtr hWnd);

        const int GW_HWNDFIRST = 0; //{同级别 Z 序最上}  
        const int GW_HWNDLAST = 1; //{同级别 Z 序最下}  
        const int GW_HWNDNEXT = 2; //{同级别 Z 序之下}  
        const int GW_HWNDPREV = 3; //{同级别 Z 序之上}  
        const int GW_OWNER = 4; //{属主窗口}  
        const int GW_CHILD = 5; //{子窗口中的最上}  

        [DllImport("user32.dll", EntryPoint = "GetWindow")]//获取窗体句柄，hwnd为源窗口句柄  
                                                           /*wCmd指定结果窗口与源窗口的关系，它们建立在下述常数基础上： 
                                                                 GW_CHILD 
                                                                 寻找源窗口的第一个子窗口 
                                                                 GW_HWNDFIRST 
                                                                 为一个源子窗口寻找第一个兄弟（同级）窗口，或寻找第一个顶级窗口 
                                                                 GW_HWNDLAST 
                                                                 为一个源子窗口寻找最后一个兄弟（同级）窗口，或寻找最后一个顶级窗口 
                                                                 GW_HWNDNEXT 
                                                                 为源窗口寻找下一个兄弟窗口 
                                                                 GW_HWNDPREV 
                                                                 为源窗口寻找前一个兄弟窗口 
                                                                 GW_OWNER 
                                                                 寻找窗口的所有者 
                                                            */
        public static extern int GetWindow(
            int hwnd,
            int wCmd
        );

        [DllImport("user32.dll", EntryPoint = "SetParent")]//设置父窗体  
        public static extern int SetParent(
            int hWndChild,
            int hWndNewParent
        );

        [DllImport("user32.dll", EntryPoint = "GetCursorPos")]//获取鼠标坐标  
        public static extern int GetCursorPos(
            ref POINTAPI lpPoint
        );

        [StructLayout(LayoutKind.Sequential)]//定义与API相兼容结构体，实际上是一种内存转换  
        public struct POINTAPI
        {
            public int X;
            public int Y;
        }

        [DllImport("user32.dll", EntryPoint = "WindowFromPoint")]//指定坐标处窗体句柄  
        public static extern int WindowFromPoint(
            int xPoint,
            int yPoint
        );

        [DllImport("User32")]
        public extern static void mouse_event(int dwFlags, int dx, int dy, int dwData, IntPtr dwExtraInfo);
        [DllImport("User32")]
        public extern static bool GetCursorPos(out POINT p);
        [StructLayout(LayoutKind.Sequential)]
        public struct POINT { public int X; public int Y; }
        public enum MouseEventFlags
        {
            Move = 0x0001,
            LeftDown = 0x0002,
            LeftUp = 0x0004,
            RightDown = 0x0008,
            RightUp = 0x0010,
            MiddleDown = 0x0020,
            MiddleUp = 0x0040,
            Wheel = 0x0800, Absolute = 0x8000
        }
        /// <summary>  
        /// 自动双击事件  
        /// </summary>  
        /// <param name="x">x坐标</param>  
        /// <param name="y">y坐标</param>  
        private void AutoDoubleClick(int x, int y)
        {
            POINT point = new POINT();
            GetCursorPos(out point);
            try
            {
                SetCursorPos(x, y);
                mouse_event((int)(MouseEventFlags.LeftDown | MouseEventFlags.Absolute), 0, 0, 0, IntPtr.Zero);
                mouse_event((int)(MouseEventFlags.LeftUp | MouseEventFlags.Absolute), 0, 0, 0, IntPtr.Zero);
                mouse_event((int)(MouseEventFlags.LeftDown | MouseEventFlags.Absolute), 0, 0, 0, IntPtr.Zero);
                mouse_event((int)(MouseEventFlags.LeftUp | MouseEventFlags.Absolute), 0, 0, 0, IntPtr.Zero);
                mouse_event((int)(MouseEventFlags.LeftDown | MouseEventFlags.Absolute), 0, 0, 0, IntPtr.Zero);
                mouse_event((int)(MouseEventFlags.LeftUp | MouseEventFlags.Absolute), 0, 0, 0, IntPtr.Zero);
            }
            finally
            {
                SetCursorPos(point.X, point.Y);
            }
        }
        /// <summary>  
        /// 自动单机事件  
        /// </summary>  
        /// <param name="x">x坐标</param>  
        /// <param name="y">y坐标</param>  
        private void AutoClick(int x, int y)
        {
            POINT point = new POINT();
            GetCursorPos(out point);
            try
            {
                SetCursorPos(x, y);
                mouse_event((int)(MouseEventFlags.LeftDown | MouseEventFlags.Absolute), 0, 0, 0, IntPtr.Zero);
                mouse_event((int)(MouseEventFlags.LeftUp | MouseEventFlags.Absolute), 0, 0, 0, IntPtr.Zero);
            }
            finally
            {
                SetCursorPos(point.X, point.Y);
            }

        }

        public Form1()
        {
            InitializeComponent();
        }
        // <summary>  
        /// 获取指定窗体的标题  
        /// </summary>  
        /// <param name="WinHandle">窗体句柄</param>  
        /// <param name="Title">缓冲区取用于存储标题</param>  
        /// <param name="size">缓冲区大小</param>  
        /// <returns></returns>  
        [DllImport("User32.dll")]
        public static extern int GetWindowText(IntPtr WinHandle, StringBuilder Title, int size);
        [DllImport("user32.dll")]
        private static extern int GetClassName(IntPtr hWnd, StringBuilder lpString, int nMaxCount);

        /// <summary>  
        /// 用于枚举子窗体是的委托  
        /// </summary>  
        /// <param name="WindowHandle">窗体句柄</param>  
        /// <param name="num">自定义</param>  
        /// <returns></returns>  
        public delegate bool EnumChildWindow(IntPtr WindowHandle, string num);
        /// <summary>  
        /// 获取指定窗体的所有子窗体  
        /// </summary>  
        /// <param name="WinHandle">窗体句柄</param>  
        /// <param name="ec">回调委托</param>  
        /// <param name="name">自定义</param>  
        /// <returns></returns>  
        [DllImport("User32.dll")]
        public static extern int EnumChildWindows(IntPtr WinHandle, EnumChildWindow ecw, string name);
        IList<IntPtr> _WindowsList = new List<IntPtr>();
        public bool GetWindows(IntPtr p_Handle, int p_Param)
        {
            StringBuilder _ClassName = new StringBuilder(255);
            StringBuilder title = new StringBuilder(255);
            GetWindowText(p_Handle, title, 255);
            GetClassName(p_Handle, _ClassName, 255);


            if (_ClassName.ToString() == "IEFrame")
                _WindowsList.Add(p_Handle);
            return true;
        }


        /// <summary>  
        /// 枚举窗体  
        /// </summary>  
        /// <param name="handle"></param>  
        /// <param name="num"></param>  
        /// <returns></returns>  
        private bool EnumChild(IntPtr handle, string num)
        {
            StringBuilder title = new StringBuilder();
            //StringBuilder type = new StringBuilder();  
            title.Length = 100;
            //type.Length = 100;  

            GetWindowText(handle, title, 100);//取标题  
            //GetClassName(handle, type, 100);//取类型  
            //listBox2.Items.Add(title);
            return true;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            //EnumWindowsProc _Proc = new EnumWindowsProc(GetWindows);
            //EnumWindows(_Proc, 0);

            //if (_WindowsList.Count > 0)
            //{
            //    hwnd = _WindowsList[0];
            //    //SetWindowPos(hwnd, -1, 0, 0, 0, 0, 1 | 2);  
            //    SetForegroundWindow(hwnd);
            //}

            ////1)、根据窗口的标题得到句柄
            //IntPtr myIntPtr = FindWindow(null, "360安全浏览器"); //null为类名，可以用Spy++得到，也可以为空
            //ShowWindow(myIntPtr, SW_RESTORE); //将窗口还原
            //SetForegroundWindow(myIntPtr); //如果没有ShowWindow，此方法不能设置最小化的窗口
            //CallBack myCallBack = new CallBack(Recall);
            //EnumWindows(myCallBack, 0);
            ////2 打开窗口
            ////Process proc = Process.Start(@"目标程序路径");
            ////SetActiveWindow(proc.MainWindowHandle);
            ////SetForegroundWindow(proc.MainWindowHandle);
            ////1 利用发送消息API（SendMessage）向窗口发送数据
            //string _GameID = "";
            //string _GamePass = "";
            //InputStr(myIntPtr, _GameID); //输入游戏ID
            //SendMessage(myIntPtr, WM_SYSKEYDOWN, 0X09, 0); //输入TAB（0x09）
            //SendMessage(myIntPtr, WM_SYSKEYUP, 0X09, 0);
            //InputStr(myIntPtr, _GamePass); //输入游戏密码
            //SendMessage(myIntPtr, WM_SYSKEYDOWN, 0X0D, 0); //输入ENTER（0x0d）
            //SendMessage(myIntPtr, WM_SYSKEYUP, 0X0D, 0);

            ////2 利用鼠标和键盘模拟向窗口发送数据

            //SetWindowPos(PW, (IntPtr)(-1), 0, 0, 0, 0, 0x0040 | 0x0001); //设置窗口位置
            //SetCursorPos(476, 177); //设置鼠标位置
            //mouse_event(0x0002, 0, 0, 0, 0); //模拟鼠标按下操作
            //mouse_event(0x0004, 0, 0, 0, 0); //模拟鼠标放开操作
            //SendKeys.Send(_GameID);   //模拟键盘输入游戏ID
            //SendKeys.Send("{TAB}"); //模拟键盘输入TAB
            //SendKeys.Send(_GamePass); //模拟键盘输入游戏密码
            //SendKeys.Send("{ENTER}"); //模拟键盘输入ENTER
            ////另：上面还提到了keybd_event方法，用法和mouse_event方法类似，作用和SendKeys.Send一样。
        }
        //回调方法Recall
        public bool Recall(int hwnd, int lParam)
        {
            StringBuilder sb = new StringBuilder(256);
            IntPtr PW = new IntPtr(hwnd);
            GetWindowTextW(PW, sb, sb.Capacity); //得到窗口名并保存在strName中
            string strName = sb.ToString();
            GetClassNameW(PW, sb, sb.Capacity); //得到窗口类名并保存在strClass中
            string strClass = sb.ToString();
            if (strName.IndexOf("窗口名关键字") >= 0 && strClass.IndexOf("类名关键字") >= 0)
            {
                return false; //返回false中止EnumWindows遍历
            }
            else
            {
                return true; //返回true继续EnumWindows遍历
            }
        }

        ///// <summary>
        //　/// 发送一个字符串
        //　/// </summary>
        //　/// <param name="myIntPtr">窗口句柄</param>
        //　/// <param name="Input">字符串</param>
        //public void InputStr(IntPtr myIntPtr, string Input)
        //{
        //    byte[] ch = (Encoding.ASCII.GetBytes(Input));
        //    for (int i = 0; i < ch.Length; i++)
        //    {
        //        SendMessage(PW, WM_CHAR, ch, 0);
        //    }
        //}

        private void button2_Click(object sender, EventArgs e)
        {
            MessageBox.Show(GetURL());
        }
        //[DllImport("User32.dll")] //User32.dll是Windows操作系统的核心动态库之一
        //static extern int FindWindow(string lpClassName, string lpWindowName);
        [DllImport("User32.dll")]
        static extern int FindWindowEx(int hwndParent, int hwndChildAfter, string lpszClass, string lpszWindow);
        [DllImport("User32.dll")]
        static extern int GetWindowText(int hwnd, StringBuilder buf, int nMaxCount);
        [DllImport("User32.dll")]
        static extern int SendMessage(int hWnd, int Msg, int wParam, StringBuilder lParam);

        const int WM_GETTEXT = 0x000D; //获得文本消息的16进制表示
        /// <summary>
        /// Get the URL of the current opened IE
        /// </summary>
        public static string GetURL()
        {
            IntPtr parent = FindWindow("IEFrame", null);
            int child = FindWindowEx(parent.ToInt32(), 0, "WorkerW", null);
            child = FindWindowEx(child, 0, "ReBarWindow32", null);
            child = FindWindowEx(child, 0, "Address Band Root", null);
            //child = FindWindowEx(child, 0, "ComboBox", null);
            //child = FindWindowEx(child, 0, "ComboBoxEx32", null);
            //child = FindWindowEx(child, 0, "ComboBox", null);
            child = FindWindowEx(child, 0, "Edit", null);   //通过SPY++获得地址栏的层次结构，然后一层一层获得
            StringBuilder buffer = new StringBuilder(1024);

            //child表示要操作窗体的句柄号
            //WM_GETTEXT表示一个消息，怎么样来驱动窗体
            //1024表示要获得text的大小
            //buffer表示获得text的值存放在内存缓存中
            int num = SendMessage(child, WM_GETTEXT, 1024, buffer);
            string URL = buffer.ToString();
            return URL;
        }
        [DllImport("oleacc.DLL", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        static extern bool ObjectFromLresult(
            [In()]int lResult,
            [In()]byte[] riid,
            [In()]int wParam,
            [Out(), MarshalAs(UnmanagedType.IUnknown)]out object ppvObject
            );
        [DllImport("user32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.I4)]
        static extern int RegisterWindowMessage(
            [In()]string lpString
            );

        [DllImport("user32.dll", EntryPoint = "SendMessageTimeoutA", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        static extern bool SendMessageTimeout(
            [In()]int MSG,
            [In()]int hWnd,
            [In()]int wParam,
            [In()]int lParam,
            [In()]int fuFlags,
            [In()]int uTimeout,
            [In(), Out()]ref int lpdwResult
            );
        //[DllImport("user32.dll", SetLastError = true)]
        //static extern int GetWindow(
        //    [In()]int hWnd,
        //    [In()]int uCmd
        //    );

        //internal const int GW_CHILD = 0x5;

        internal byte[] RIID = new byte[] { 32, 197, 111, 98, 30, 164, 207, 17, 167, 49, 0, 160, 201, 8, 38, 55 };
        //声明Document对象（如果用内嵌浏览器，我们得到的是一个HTMLDocument） 
        HTMLDocumentClass document = null;
        [DllImport("user32", EntryPoint = "FindWindow")]
        public static extern int FindWindowA(string lpClassName, string lpWindowName);
        private void button3_Click(object sender, EventArgs e)
        {
            //查找打开的窗口句柄 
            int iehwnd = FindWindowA("IEFrame", "百度一下，你就知道 - Internet Explorer");
            //初始化所有IE窗口 
            IShellWindows sw = new ShellWindowsClass();
            //轮询所有IE窗口 
            for (int i = sw.Count - 1; i >= 0; i--)
            {
                //得到每一个IE的 IWebBrowser2 对象 
                IWebBrowser2 iwb2 = sw.Item(i) as IWebBrowser2;
                //比对 得到的 句柄是否符合查找的窗口句柄 
                if (iwb2.HWND == iehwnd)
                {
                    //查找成功 进行赋值 
                    document = (HTMLDocumentClass)iwb2.Document;
                    //对网页进行操作 
                    document.getElementById("kw").setAttribute("value","1111");
                    //document.getElementById("kw").onclick();
                }
            }
            //IntPtr parent = FindWindow("IEFrame", null);
            //int child = FindWindowEx(parent.ToInt32(), 0, "Frame Tab", null);
            //child = FindWindowEx(child, 0, "TabWindowClass", null);
            //child = FindWindowEx(child, 0, "Shell DocObject View", null);
            //child = FindWindowEx(child, 0, "Internet Explorer_Server", null);

            ////var oBrowser = DiagnosticsGlobalScope.browser;
            //IHTMLDocument2 doc =(IHTMLDocument2)GetDocument(child);
            //MessageBox.Show(doc.title);
            //IHTMLElement2 elm=(IHTMLElement2)doc.all.item("wd",0);
            //MessageBox.Show(elm.ToString());
            //FramesCollection frames= doc.frames;
            ////if (frames.length > 0)
            ////{
            ////    for(int i = 0; i < frames.length; i++)
            ////    {
            ////        IHTMLWindow2 window=frames.item(i);
            ////    }
            ////}
            //IHTMLElementCollection coll = doc.forms;
            //if (coll.length > 0)
            //{
            //    for (int i = 0; i < coll.length; i++)
            //    {
            //        IHTMLElement window = coll.item(i);
            //        string name=window.getAttribute("name");
            //        if (name == "wd")
            //        {
            //            window.setAttribute("value", "1");
            //        }
            //    }
            //}
            //IHTMLDOMChildrenCollection collect = (IHTMLDOMChildrenCollection)doc.childNodes;

            //foreach (IHTMLDOMNode node in collect)
            //{
            //    //因为关闭节点也会有（比如</a>，但是这样的节点会被定义为HTMLUnknownElementClass）  
            //    //所以要判断这个节点是不是未知节点不是才处理  
            //    if (!(node is IHTMLUnknownElement))
            //    {

            //        //获取属性集合  
            //        IHTMLAttributeCollection attrs = (IHTMLAttributeCollection)node.attributes;
            //        foreach (IHTMLDOMAttribute attr in attrs)
            //        {

            //            //只有specified=true的属性才是你要的  
            //            if (attr.specified)
            //            {
            //                Console.Write(attr.nodeValue);
            //            }
            //        }
            //    }
            //}
        }

        private object GetDocument(int hWnd)
        {
            object _ComObject = null;
            int lpdwResult = 0, Msg = RegisterWindowMessage("WM_HTML_GETOBJECT");
            if (!SendMessageTimeout(hWnd, Msg, 0, 0, 2, 1000, ref lpdwResult))
            {
                return null;
            }
            if (ObjectFromLresult(lpdwResult, RIID, 0, out _ComObject))
            {
                return null;
            }
            return _ComObject;
        }
    }
}
