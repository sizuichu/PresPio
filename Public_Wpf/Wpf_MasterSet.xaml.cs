using System.Windows;
using HandyControl.Controls;

namespace PresPio
    {
    public partial class Wpf_MasterSet
        {
        public Wpf_MasterSet()
            {
            InitializeComponent();
            LoadSettings();
            }

        private void LoadSettings()
            {
            // 从设置中加载选项状态
            DeleteSourceCheck.IsChecked = Properties.Settings.Default.Pla_N1;
            SkipShapeCheck.IsChecked = Properties.Settings.Default.Pla_N2;
            KeepTextCheck.IsChecked = Properties.Settings.Default.Pla_N3;
            SkipPictureCheck.IsChecked = Properties.Settings.Default.Pla_N4;
            }

        private void SaveSettings()
            {
            // 保存设置到配置文件
            Properties.Settings.Default.Pla_N1 = DeleteSourceCheck.IsChecked ?? false;
            Properties.Settings.Default.Pla_N2 = SkipShapeCheck.IsChecked ?? false;
            Properties.Settings.Default.Pla_N3 = KeepTextCheck.IsChecked ?? false;
            Properties.Settings.Default.Pla_N4 = SkipPictureCheck.IsChecked ?? false;
            Properties.Settings.Default.Save();
            }

        private void SaveButton_Click(object sender, RoutedEventArgs e)
            {
            SaveSettings();
            Growl.Success("设置已保存");
            }

        private void CloseButton_Click(object sender, RoutedEventArgs e)
            {
            SaveSettings();
            Close();
            }

        protected override void OnClosing(System.ComponentModel.CancelEventArgs e)
            {
            SaveSettings();
            base.OnClosing(e);
            }
        }
    }