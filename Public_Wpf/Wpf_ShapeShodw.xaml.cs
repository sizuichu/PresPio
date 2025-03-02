using System.Windows;
using System.Windows.Media;
using HandyControl.Controls;
using HandyControl.Tools;
using Microsoft.Office.Core;

namespace PresPio
    {
    /// <summary>
    /// Wpf_ShapeShodw.xaml 的交互逻辑
    /// </summary>
    public partial class Wpf_ShapeShodw
        {
        public Wpf_ShapeShodw()
            {
            InitializeComponent();
            }

        private void GrowWinDow_Loaded(object sender, RoutedEventArgs e)
            {
            //获取设置
            ShodwTra.Value = (int)Properties.Settings.Default.ShodwTra;
            ShodwSize.Value = (int)Properties.Settings.Default.ShodwSize;
            ShodwBlur.Value = (int)Properties.Settings.Default.ShodwBlur;
            ShodwX.Value = (int)Properties.Settings.Default.ShodwX;

            // 获取存储在应用程序设置中的颜色值
            System.Drawing.Color color = Properties.Settings.Default.ShodwColor;

            // 将System.Drawing.Color转换为System.Windows.Media.Color
            System.Windows.Media.Color wpfColor = System.Windows.Media.Color.FromArgb(color.A, color.R, color.G, color.B);

            // 使用SolidColorBrush创建一个画刷，这个画刷使用从设置中获取的颜色值
            SolidColorBrush brush = new SolidColorBrush(wpfColor);

            ColorBtn.Background = brush;

            //加载边框按钮
            if (Properties.Settings.Default.ShodwCheck == MsoTriState.msoTrue)
                {
                CheckBtn.IsChecked = true;
                }
            else
                {
                CheckBtn.IsChecked = false;
                }
            }

        private void CheckBtn_Checked(object sender, RoutedEventArgs e)
            {
            if (CheckBtn.IsChecked == true)
                {
                Properties.Settings.Default.ShodwCheck = Microsoft.Office.Core.MsoTriState.msoTrue;
                }
            else
                {
                Properties.Settings.Default.ShodwCheck = Microsoft.Office.Core.MsoTriState.msoFalse;
                }
            Properties.Settings.Default.Save();
            }

        private void ShodwTra_ValueChanged(object sender, HandyControl.Data.FunctionEventArgs<double> e)
            {
            Properties.Settings.Default.ShodwTra = ShodwTra.Value;
            Properties.Settings.Default.Save();
            }

        private void ShodwSize_ValueChanged(object sender, HandyControl.Data.FunctionEventArgs<double> e)
            {
            Properties.Settings.Default.ShodwSize = (float)ShodwSize.Value;
            Properties.Settings.Default.Save();
            }

        private void ShodwBlur_ValueChanged(object sender, HandyControl.Data.FunctionEventArgs<double> e)
            {
            Properties.Settings.Default.ShodwBlur = (float)ShodwBlur.Value;
            Properties.Settings.Default.Save();
            }

        private void ShodwX_ValueChanged(object sender, HandyControl.Data.FunctionEventArgs<double> e)
            {
            Properties.Settings.Default.ShodwX = (float)ShodwX.Value;
            Properties.Settings.Default.Save();
            }

        private void Button_Click(object sender, RoutedEventArgs e)
            {
            //恢复默认值
            CheckBtn.IsChecked = false;
            Properties.Settings.Default.ShodwTra = 85;
            ShodwTra.Value = 85;
            Properties.Settings.Default.ShodwSize = 102;
            ShodwSize.Value = 102;
            Properties.Settings.Default.ShodwBlur = 5;
            ShodwBlur.Value = 5;
            Properties.Settings.Default.ShodwX = 0;
            ShodwX.Value = 0;
            Properties.Settings.Default.Save();
            }

        private void ColorBtn_Click(object sender, RoutedEventArgs e)
            {
            var picker = SingleOpenHelper.CreateControl<ColorPicker>();
            var window = new PopupWindow
                {
                PopupElement = picker,
                AllowsTransparency = true,
                WindowStyle = WindowStyle.None,
                MinWidth = 0,
                MinHeight = 0,
                };

            // 获取当前窗口的位置
            var ownerWindow = System.Windows.Window.GetWindow(this);
            var ownerWindowPoint = ownerWindow.PointToScreen(new System.Windows.Point(0, 0));

            // 设置 PopupWindow 的位置
            window.Left = ownerWindowPoint.X + ownerWindow.Width / 2; // 将窗口放置在当前窗口的右侧，这里假设间距为10像素
            window.Top = ownerWindowPoint.Y / 2; // 与当前窗口顶部对齐

            // 添加确定按钮事件处理
            picker.Confirmed += (colorPicker, args) =>
            {
                var selectedColor = picker.SelectedBrush;
                ColorBtn.Background = selectedColor;

                // 先将Brush转换为SolidColorBrush
                SolidColorBrush solidColorBrush = selectedColor as SolidColorBrush;

                // 获取颜色属性
                var mediaColor = solidColorBrush.Color;

                // 将System.Windows.Media.Color转换为System.Drawing.Color
                System.Drawing.Color drawingColor = System.Drawing.Color.FromArgb(mediaColor.A, mediaColor.R, mediaColor.G, mediaColor.B);

                Properties.Settings.Default.ShodwColor = drawingColor;
                Properties.Settings.Default.Save();
                window.Close();
            };

            // 添加取消选择事件处理
            picker.Canceled += (colorPicker, args) => window.Close();

            window.Show();
            }

        private void Button_Click_1(object sender, RoutedEventArgs e)
            {
            this.Close();
            }

        private void Button_Click_2(object sender, RoutedEventArgs e)
            {
            Growl.SuccessGlobal("设置成功！");
            this.Close();
            }
        }
    }