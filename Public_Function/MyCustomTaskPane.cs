using System;
using Microsoft.Office.Core;

namespace PresPio.Function
    {
    internal class MyCustomTaskPane
        {
        public MyCustomTaskPane()
            {
            // 创建自定义窗格
            }

        private void taskPane_DialogLauncherClick(object sender, EventArgs e)
            {
            CustomTaskPane taskPane = (CustomTaskPane)sender;

            if (taskPane.Visible)
                {
                // 自定义窗格已经打开，将其关闭
                taskPane.Visible = false;
                }
            else
                {
                // 自定义窗格尚未打开，将其打开
                taskPane.Visible = true;
                }
            }
        }
    }