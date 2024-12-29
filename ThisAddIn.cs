using Microsoft.Office.Interop.PowerPoint;
using Microsoft.Office.Tools;
using System;
using System.Windows.Forms;

namespace PresPio
    {
    public partial class ThisAddIn
        {
        // 静态类，用于共享任务窗格
        public static class TaskPaneShared
            {
            public static CustomTaskPane taskPane; // 定义一个任务窗格
            }

        private Theme _currentTheme; // 当前主题

        private void ThisAddIn_Startup(object sender, EventArgs e)
            {
            var classGetIcons = new RibbonIcoHelper();
            classGetIcons.GetIcons("PPT"); // 加载图片

            // 订阅选中幻灯片事件
            Application.WindowSelectionChange += Application_WindowSelectionChange;
            // 订阅形状大小改变事件
            //  Application.AfterShapeSizeChange += Application_AfterShapeSizeChange;
            }

        private void ThisAddIn_Shutdown(object sender, EventArgs e)
            {
            // 取消订阅选中幻灯片事件
            Application.WindowSelectionChange -= Application_WindowSelectionChange;
            // 取消订阅形状大小改变事件
            // Application.AfterShapeSizeChange -= Application_AfterShapeSizeChange;
            }

        // 当选中幻灯片或形状时触发的事件处理程序
        private void Application_WindowSelectionChange(Selection sel)
            {
            var ribbon = Globals.Ribbons.Ribbon1; // 获取功能区实例

            if (sel.Type == PpSelectionType.ppSelectionShapes)
                {
                // 如果选中了形状，启用相应的按钮
                EnableShapeButtons(ribbon, sel.ShapeRange.Count);
                }
            else
                {
                // 如果未选中形状，禁用相应的按钮
                DisableShapeButtons(ribbon);
                }
            }

        // 当形状大小改变时触发的事件处理程序
        private void Application_AfterShapeSizeChange(Shape shape)
            {
            // 显示消息框提示形状大小已更改
            MessageBox.Show("形状大小已更改");
            }

        // 启用与形状相关的功能区按钮
        private void EnableShapeButtons(MyRibbon ribbon, int shapeCount)
            {
            ribbon.button96.Enabled = true;    // 启用编辑顶点按钮
            ribbon.splitButton14.Enabled = true; // 启用原位复制按钮
            ribbon.button13.Enabled = true;    // 启用边缘剪裁按钮
            ribbon.button143.Enabled = true;   // 启用内容翻转按钮
            ribbon.button146.Enabled = true;   // 启用内容居中按钮
            ribbon.button54.Enabled = true;    // 启用配色按钮
            ribbon.button52.Enabled = true;    // 启用配色按钮
            ribbon.button127.Enabled = true;   // 启用配色按钮
            ribbon.button132.Enabled = true;   // 启用配色按钮
            ribbon.button144.Enabled = true;   // 启用矩阵复制按钮
            ribbon.splitButton9.Enabled = true; // 启用间距等按钮
            ribbon.splitButton10.Enabled = true; // 启用间距等按钮
            ribbon.splitButton11.Enabled = true; // 启用间距等按钮
            ribbon.button152.Enabled = true;   // 启用超级组合按钮
            ribbon.splitButton19.Enabled = true;
            ribbon.splitButton20.Enabled = true;
            if (shapeCount == 2)
                {
                ribbon.button109.Enabled = true; // 启用选择两个内容按钮
                }
            else
                {
                ribbon.button109.Enabled = false; // 禁用选择两个内容按钮
                }

            if (shapeCount >= 2)
                {
                ribbon.button145.Enabled = true; // 启用格式统一按钮
                ribbon.button78.Enabled = true;  // 启用组合按钮
                ribbon.splitButton17.Enabled = true;  // 互换按钮
                }
            else
                {
                ribbon.button145.Enabled = false; // 禁用格式统一按钮
                ribbon.button78.Enabled = false;  // 禁用组合按钮
                }

            ribbon.button7.Enabled = true; // 启用矩阵复制按钮
            }

        // 禁用与形状相关的功能区按钮
        private void DisableShapeButtons(MyRibbon ribbon)
            {
            ribbon.button96.Enabled = false;   // 禁用编辑顶点按钮
            ribbon.splitButton14.Enabled = false; // 禁用原位复制按钮
            ribbon.button13.Enabled = false;   // 禁用边缘剪裁按钮
            ribbon.button143.Enabled = false;  // 禁用内容翻转按钮
            ribbon.button146.Enabled = false;  // 禁用内容居中按钮
            ribbon.button54.Enabled = false;   // 禁用配色按钮
            ribbon.button52.Enabled = false;   // 禁用配色按钮
            ribbon.button127.Enabled = false;  // 禁用配色按钮
            ribbon.button132.Enabled = false;  // 禁用配色按钮
            ribbon.button144.Enabled = false;  // 禁用矩阵复制按钮
            ribbon.button145.Enabled = false;  // 禁用格式统一按钮
            ribbon.splitButton9.Enabled = false; // 禁用间距等按钮
            ribbon.splitButton10.Enabled = false; // 禁用间距等按钮
            ribbon.splitButton11.Enabled = false; // 禁用间距等按钮
            ribbon.button78.Enabled = false;   // 禁用组合按钮
            ribbon.button152.Enabled = false;  // 禁用超级组合按钮
            ribbon.button7.Enabled = false;    // 禁用矩阵复制按钮
            ribbon.button109.Enabled = false;  // 禁用选择两个内容按钮
            ribbon.splitButton19.Enabled = false;
            ribbon.splitButton17.Enabled = false;  // 互换按钮
            ribbon.splitButton20.Enabled = false;
            }

        private void InternalStartup()
            {
            this.Startup += ThisAddIn_Startup;  // 注册启动事件处理程序
            this.Shutdown += ThisAddIn_Shutdown; // 注册关闭事件处理程序
            }
        }
    }