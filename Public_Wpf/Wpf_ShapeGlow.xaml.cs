using System.Windows;
using System.Windows.Media;
using HandyControl.Controls;
using HandyControl.Tools;

namespace PresPio
    {
    /// <summary>
    /// Wpf_ShapeGlow.xaml 的交互逻辑
    /// </summary>
    public partial class Wpf_ShapeGlow
        {
        public Wpf_ShapeGlow()
            {
            InitializeComponent();
            LoadSet();
            }

        public void LoadSet()
            {
            //发光值设置
            sizeBtn.Value = Properties.Settings.Default.GlowNum;

            //透明度设置
            growBtn.Value = (int)Properties.Settings.Default.GlowTra;

            //颜色设置
            // 获取存储在应用程序设置中的颜色值
            System.Drawing.Color color = Properties.Settings.Default.GlowColor;

            // 将System.Drawing.Color转换为System.Windows.Media.Color
            System.Windows.Media.Color wpfColor = System.Windows.Media.Color.FromArgb(color.A, color.R, color.G, color.B);

            // 使用SolidColorBrush创建一个画刷，这个画刷使用从设置中获取的颜色值
            SolidColorBrush brush = new SolidColorBrush(wpfColor);

            // 将按钮的背景颜色设置为这个画刷
            ColorButton.Background = brush;
            }

        private void OkBtn_Click(object sender, RoutedEventArgs e)
            {
            this.Close();
            Growl.Info("设置成功！");
            }

        private void ColorButton_Click(object sender, RoutedEventArgs e)
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
                ColorButton.Background = selectedColor;

                // 先将Brush转换为SolidColorBrush
                SolidColorBrush solidColorBrush = selectedColor as SolidColorBrush;

                // 获取颜色属性
                var mediaColor = solidColorBrush.Color;

                // 将System.Windows.Media.Color转换为System.Drawing.Color
                System.Drawing.Color drawingColor = System.Drawing.Color.FromArgb(mediaColor.A, mediaColor.R, mediaColor.G, mediaColor.B);

                Properties.Settings.Default.GlowColor = drawingColor;
                Properties.Settings.Default.Save();
                window.Close();
            };

            // 添加取消选择事件处理
            picker.Canceled += (colorPicker, args) => window.Close();

            window.Show();
            }

        private void sizeBtn_ValueChanged(object sender, HandyControl.Data.FunctionEventArgs<double> e)
            {
            Properties.Settings.Default.GlowNum = (int)sizeBtn.Value;
            Properties.Settings.Default.Save();
            }

        private void growBtn_ValueChanged(object sender, HandyControl.Data.FunctionEventArgs<double> e)
            {
            Properties.Settings.Default.GlowTra = growBtn.Value;
            Properties.Settings.Default.Save();
            }
        }
    }