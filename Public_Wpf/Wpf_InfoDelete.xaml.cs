using HandyControl.Controls;
using Microsoft.Office.Interop.PowerPoint;
using System.Collections.Generic;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;

using Powerpoint = Microsoft.Office.Interop.PowerPoint;

namespace PresPio
    {
    /// <summary>
    /// Wpf_InfoDelete.xaml 的交互逻辑
    /// </summary>
    public partial class Wpf_InfoDelete
        {
        public Powerpoint.Application app;

        public Wpf_InfoDelete()
            {
            app = Globals.ThisAddIn.Application;
            InitializeComponent();
            }

        private void Button_Click(object sender, RoutedEventArgs e)
            {
            // 获取当前 PowerPoint 应用程序和演示文稿
            var app = Globals.ThisAddIn.Application;
            var presentation = app.ActivePresentation;

            // 定义用于存储需要移除的文档信息类型的字典
            Dictionary<CheckBox, PpRemoveDocInfoType> docInfoTypes = new Dictionary<CheckBox, PpRemoveDocInfoType>
            {
                { uiCheckBox2, PpRemoveDocInfoType.ppRDIComments },
                { uiCheckBox3, PpRemoveDocInfoType.ppRDIContentType },
                { uiCheckBox4, PpRemoveDocInfoType.ppRDIDocumentManagementPolicy },
                { uiCheckBox5, PpRemoveDocInfoType.ppRDIDocumentProperties },
                { uiCheckBox6, PpRemoveDocInfoType.ppRDIDocumentServerProperties },
                { uiCheckBox7, PpRemoveDocInfoType.ppRDIDocumentWorkspace },
                { uiCheckBox8, PpRemoveDocInfoType.ppRDIInkAnnotations },
                { uiCheckBox9, PpRemoveDocInfoType.ppRDIPublishPath },
                { uiCheckBox10, PpRemoveDocInfoType.ppRDIRemovePersonalInformation },
                { uiCheckBox11, PpRemoveDocInfoType.ppRDISlideUpdateInformation }
            };

            // 遍历复选框，并移除选中状态的文档信息
            foreach (var kvp in docInfoTypes)
                {
                if (kvp.Key.IsChecked == true)
                    {
                    presentation.RemoveDocumentInformation(kvp.Value);
                    }
                }

            // 显示成功消息
            Growl.SuccessGlobal("清理成功！");
            }

        private void Button_Click_1(object sender, RoutedEventArgs e)
            {
            foreach (CheckBox control in GetAllChildControls(Grid1))
                {
                control.IsChecked = false;
                }
            }

        public IEnumerable<Control> GetAllChildControls(DependencyObject parent)
            {
            var count = VisualTreeHelper.GetChildrenCount(parent);
            for (int i = 0 ; i < count ; i++)
                {
                var child = VisualTreeHelper.GetChild(parent, i);
                if (child is Control)
                    {
                    yield return child as Control;
                    }

                foreach (var grandChild in GetAllChildControls(child))
                    {
                    yield return grandChild;
                    }
                }
            }
        }
    }