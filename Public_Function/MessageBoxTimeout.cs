using System;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace PresPio.Public_Function
    {
    public class MessageBoxTimeout
        {
        /// <summary>
        /// 自动超时消息提示框
        /// </summary>
        public class MessageBoxTimeOut
            {
            /// <summary>
            /// 标题
            /// </summary>
            private static string _caption;

            /// <summary>
            /// 显示消息框
            /// </summary>
            /// <param name="text">消息内容</param>
            /// <param name="caption">标题</param>
            /// <param name="timeout">超时时间，单位：毫秒</param>
            public static void Show(string text, string caption, int timeout)
                {
                _caption = caption;
                StartTimer(timeout);
                MessageBox.Show(text, caption);
                }

            /// <summary>
            /// 显示消息框
            /// </summary>
            /// <param name="text">消息内容</param>
            /// <param name="caption">标题</param>
            /// <param name="timeout">超时时间，单位：毫秒</param>
            /// <param name="buttons">消息框上的按钮</param>
            public static void Show(string text, string caption, int timeout, MessageBoxButtons buttons)
                {
                _caption = caption;
                StartTimer(timeout);
                MessageBox.Show(text, caption, buttons);
                }

            /// <summary>
            /// 显示消息框
            /// </summary>
            /// <param name="text">消息内容</param>
            /// <param name="caption">标题</param>
            /// <param name="timeout">超时时间，单位：毫秒</param>
            /// <param name="buttons">消息框上的按钮</param>
            /// <param name="icon">消息框上的图标</param>
            public static void Show(string text, string caption, int timeout, MessageBoxButtons buttons, MessageBoxIcon icon)
                {
                _caption = caption;
                StartTimer(timeout);
                MessageBox.Show(text, caption, buttons, icon);
                }

            /// <summary>
            /// 显示消息框
            /// </summary>
            /// <param name="owner">消息框所有者</param>
            /// <param name="text">消息内容</param>
            /// <param name="caption">标题</param>
            /// <param name="timeout">超时时间，单位：毫秒</param>
            public static void Show(IWin32Window owner, string text, string caption, int timeout)
                {
                _caption = caption;
                StartTimer(timeout);
                MessageBox.Show(owner, text, caption);
                }

            /// <summary>
            /// 显示消息框
            /// </summary>
            /// <param name="owner">消息框所有者</param>
            /// <param name="text">消息内容</param>
            /// <param name="caption">标题</param>
            /// <param name="timeout">超时时间，单位：毫秒</param>
            /// <param name="buttons">消息框上的按钮</param>
            public static void Show(IWin32Window owner, string text, string caption, int timeout, MessageBoxButtons buttons)
                {
                _caption = caption;
                StartTimer(timeout);
                MessageBox.Show(owner, text, caption, buttons);
                }

            /// <summary>
            /// 显示消息框
            /// </summary>
            /// <param name="owner">消息框所有者</param>
            /// <param name="text">消息内容</param>
            /// <param name="caption">标题</param>
            /// <param name="timeout">超时时间，单位：毫秒</param>
            /// <param name="buttons">消息框上的按钮</param>
            /// <param name="icon">消息框上的图标</param>
            public static void Show(IWin32Window owner, string text, string caption, int timeout, MessageBoxButtons buttons, MessageBoxIcon icon)
                {
                _caption = caption;
                StartTimer(timeout);
                MessageBox.Show(owner, text, caption, buttons, icon);
                }

            private static void StartTimer(int interval)
                {
                Timer timer = new Timer();
                timer.Interval = interval;
                timer.Tick += new EventHandler(Timer_Tick);
                timer.Enabled = true;
                }

            private static void Timer_Tick(object sender, EventArgs e)
                {
                KillMessageBox();
                //停止计时器
                ((Timer)sender).Enabled = false;
                }

            [DllImport("user32.dll", EntryPoint = "FindWindow", CharSet = CharSet.Auto)]
            private static extern IntPtr FindWindow(string lpClassName, string lpWindowName);

            [DllImport("user32.dll", CharSet = CharSet.Auto)]
            private static extern int PostMessage(IntPtr hWnd, int msg, IntPtr wParam, IntPtr lParam);

            private const int WM_CLOSE = 0x10;

            private static void KillMessageBox()
                {
                //查找MessageBox的弹出窗口,注意对应标题
                IntPtr ptr = FindWindow(null, _caption);
                if (ptr != IntPtr.Zero)
                    {
                    //查找到窗口则关闭
                    PostMessage(ptr, WM_CLOSE, IntPtr.Zero, IntPtr.Zero);
                    }
                }
            }
        }
    }