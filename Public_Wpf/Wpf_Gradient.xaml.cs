using HandyControl.Controls;
using System.Windows;

namespace PresPio
    {
    /// <summary>
    /// Wpf_Gradient.xaml 的交互逻辑
    /// </summary>
    public partial class Wpf_Gradient
        {
        public Wpf_Gradient()
            {
            InitializeComponent();
            LoadSet();
            }

        public void LoadSet()
            {
            int Num = Properties.Settings.Default.GradeColorN;
            ColorSet.Value = Num;
            }

        private void ColorSet_ValueChanged(object sender, HandyControl.Data.FunctionEventArgs<double> e)
            {
            int GradeColorN = (int)ColorSet.Value;
            Properties.Settings.Default.GradeColorN = GradeColorN;
            Properties.Settings.Default.Save();
            }

        private void Button_Click(object sender, RoutedEventArgs e)
            {
            this.Close();
            Growl.SuccessGlobal("设置成功！");
            }
        }
    }