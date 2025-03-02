using System.Windows;
using Itp.WpfAppBar;

namespace PresPio
    {
    /// <summary>
    /// Wpf_colorTheme.xaml 的交互逻辑
    /// </summary>
    public partial class Wpf_colorTheme
        {
        public Wpf_colorTheme()
            {
            InitializeComponent();
            this.DockMode = AppBarDockMode.Right;
            }

        private void CloseBtn_Click(object sender, RoutedEventArgs e)
            {
            this.Close();
            }
        }
    }