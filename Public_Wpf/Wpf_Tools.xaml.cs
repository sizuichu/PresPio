using System.Windows;
using Itp.WpfAppBar;
using Powerpoint = Microsoft.Office.Interop.PowerPoint;

// ReSharper disable IdentifierTypo
// ReSharper disable InconsistentNaming
// ReSharper disable EnumUnderlyingTypeIsInt
// ReSharper disable MemberCanBePrivate.Local
// ReSharper disable UnusedMember.Local
// ReSharper disable UnusedMember.Global
namespace PresPio
    {
    /// <summary>
    /// Wpf_Tools.xaml 的交互逻辑
    /// </summary>
    public partial class Wpf_Tools
        {
        private Powerpoint.Application app;

        public Wpf_Tools()
            {
            InitializeComponent();
            app = Globals.ThisAddIn.Application;
            GenerateButtons();
            }

        private void ToolsWindow_Loaded(object sender, RoutedEventArgs e)
            {
            this.DockMode = AppBarDockMode.Right;
            }

        //窗体功能
        #region

        private void ToolsWindow_Closed(object sender, System.EventArgs e)
            {
            }

        private void Button_Click_1(object sender, RoutedEventArgs e)
            {
            this.ShowInTaskbar = true;
            this.WindowState = WindowState.Minimized;
            }

        private void GenerateButtons()
            {
            }

        private void LeftBtn_Click(object sender, RoutedEventArgs e)
            {
            this.DockMode = AppBarDockMode.Left;
            }

        private void rightBtn_Click(object sender, RoutedEventArgs e)
            {
            this.DockMode = AppBarDockMode.Right;
            }

        private void closeBtn_Click(object sender, RoutedEventArgs e)
            {
            MyRibbon RB = Globals.Ribbons.Ribbon1;
            RB.button138.Enabled = true;
            this.Close();
            }

        private void Button_Click(object sender, RoutedEventArgs e)
            {
            }

        private void ToolsWindow_MouseEnter(object sender, System.Windows.Input.MouseEventArgs e)
            {
            // 激活控件
            if (sender is UIElement element)
                {
                element.Focus();
                }
            }

        private void ToolsWindow_MouseLeave(object sender, System.Windows.Input.MouseEventArgs e)
            {
            //// 取消激活控件
            //if (sender is UIElement element && !element.IsMouseOver)
            //{
            //   app.Activate();
            //}
            }

        private void CloseBtn_Click_1(object sender, RoutedEventArgs e)
            {
            var ribbon = Globals.Ribbons.Ribbon1; // 获取功能区实例
            ribbon.group5.Visible = true;
            this.Close();
            }

        private void LeftBtn_Click_1(object sender, RoutedEventArgs e)
            {
            this.DockMode = AppBarDockMode.Left;
            }

        private void RightBtn_Click_1(object sender, RoutedEventArgs e)
            {
            this.DockMode = AppBarDockMode.Right;
            }

        #endregion

        //以下是功能
        private void Horizontally_Click(object sender, RoutedEventArgs e)
            {
            app = Globals.ThisAddIn.Application;
            app.CommandBars.ExecuteMso("FlipHorizontal");
            }
        }
    }