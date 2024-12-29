using Microsoft.Office.Interop.PowerPoint;
using PowerPoint= Microsoft.Office.Interop.PowerPoint;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.Collections.ObjectModel;

namespace PresPio
    {
    /// <summary>
    /// Wpf_NameUnification.xaml 的交互逻辑
    /// </summary>
    public partial class Wpf_NameUnification
        {
        public PowerPoint.Application app { get; set; }
        public Wpf_NameUnification()
            {
            InitializeComponent();
            }
      
        private void NameWindow_Loaded(object sender, RoutedEventArgs e)
            {

      
            // 获取 PowerPoint 应用程序和当前幻灯片
            var app = Globals.ThisAddIn.Application;
            PowerPoint.Slide slide = app.ActiveWindow.View.Slide;

            // 遍历幻灯片中的形状，并将其类型和中文名称添加到 ShapeList 中，然后将 ShapeList 添加到 listViewItems 列表中
            foreach (PowerPoint.Shape shape in slide.Shapes)
                {
                // 将 listViewItems
                string ItemName=shape.Name ;
                NameListView.Items.Add(ItemName);
                }

            }
        }
  }
