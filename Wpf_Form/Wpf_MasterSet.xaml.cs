using System.Windows;

namespace PresPio
    {
    /// <summary>
    /// Wpf_MasterSet.xaml 的交互逻辑
    /// </summary>
    public partial class Wpf_MasterSet
    {
        public Wpf_MasterSet()
        {
            InitializeComponent();
        }

        private void MastWindown_Loaded(object sender, RoutedEventArgs e)
        {
            float Fsize = Properties.Settings.Default.Pla_size;
            bool N1 = Properties.Settings.Default.Pla_N1;
            bool N2 = Properties.Settings.Default.Pla_N2;
            bool N3 = Properties.Settings.Default.Pla_N3;
            bool N4 = Properties.Settings.Default.Pla_N4;
            numericUpDown.Value = (int)Fsize;
            uiCheckBox1.IsChecked = N1;
            uiCheckBox2.IsChecked = N2;
            uiCheckBox3.IsChecked = N3;
            uiCheckBox4.IsChecked = N4;
        }

        private void uiCheckBox1_Checked(object sender, RoutedEventArgs e)
        {
            Properties.Settings.Default.Pla_N1 = (bool)uiCheckBox1.IsChecked;
            Properties.Settings.Default.Save();
        }

        private void uiCheckBox2_Checked(object sender, RoutedEventArgs e)
        {
            Properties.Settings.Default.Pla_N2 = (bool)uiCheckBox2.IsChecked;
            Properties.Settings.Default.Save();
        }

        private void uiCheckBox3_Checked(object sender, RoutedEventArgs e)
        {
            Properties.Settings.Default.Pla_N3 = (bool)uiCheckBox3.IsChecked;
            Properties.Settings.Default.Save();
        }

        private void uiCheckBox4_Checked(object sender, RoutedEventArgs e)
        {
            Properties.Settings.Default.Pla_N4 = (bool)uiCheckBox4.IsChecked;
            Properties.Settings.Default.Save();
        }

        private void numericUpDown_ValueChanged(object sender, HandyControl.Data.FunctionEventArgs<double> e)
        {
            Properties.Settings.Default.Pla_size = (float)numericUpDown.Value;
            Properties.Settings.Default.Save();
        }
    }
}
