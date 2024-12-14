using HandyControl.Controls;
using HandyControl.Tools;
using HandyControl.Tools.Extension;
using System.Drawing;
using System.Windows;
using System.Windows.Documents;

namespace PresPio
    {
    /// <summary>
    /// Wpf_About.xaml 的交互逻辑
    /// </summary>
    public partial class Wpf_About
        {
        public Wpf_About()
            {
            HandyControl.Tools.ConfigHelper.Instance.SetLang("zh-CN");
            InitializeComponent();
           
            //设置内容
            checkBox1.IsChecked = Properties.Settings.Default.Group2;
            checkBox2.IsChecked = Properties.Settings.Default.Group3;
            checkBox3.IsChecked = Properties.Settings.Default.Group4;
            checkBox4.IsChecked = Properties.Settings.Default.Group5;
            checkBox5.IsChecked = Properties.Settings.Default.Group6;
            checkBox6.IsChecked = Properties.Settings.Default.Group7;
          
            }


     
        private void About_Windown_Loaded(object sender, System.Windows.RoutedEventArgs e)
            {

            AboutText();
            }
        private void AboutText()
            {
            // 创建新的 FlowDocument
            FlowDocument newDocument = new FlowDocument();

            // 创建一个新的段落并添加到 FlowDocument 中
            Paragraph newParagraph = new Paragraph(new Run("This is the new content of the RichTextBox."));
            newDocument.Blocks.Add(newParagraph);

            // 设置 RichTextBox 的 FlowDocument 属性为新的 FlowDocument
            myRichTextBox.Document = newDocument;
            }
        private void checkBox1_Click(object sender, System.Windows.RoutedEventArgs e)
            {
            // 获取CheckBox的布尔值
            bool isCheckedValue = checkBox1.IsChecked.HasValue && checkBox1.IsChecked.Value;
            Globals.Ribbons.Ribbon1.group2.Visible = isCheckedValue;
            Properties.Settings.Default.Group2 = isCheckedValue;
            Properties.Settings.Default.Save();
            }

        private void checkBox2_Click(object sender, System.Windows.RoutedEventArgs e)
            {
            // 获取CheckBox的布尔值
            bool isCheckedValue = checkBox2.IsChecked.HasValue && checkBox2.IsChecked.Value;
            Globals.Ribbons.Ribbon1.group3.Visible = isCheckedValue;
            Properties.Settings.Default.Group3 = isCheckedValue;
            Properties.Settings.Default.Save();
            }

        private void checkBox3_Click(object sender, System.Windows.RoutedEventArgs e)
            {
            // 获取CheckBox的布尔值
            bool isCheckedValue = checkBox3.IsChecked.HasValue && checkBox3.IsChecked.Value;
            Globals.Ribbons.Ribbon1.group4.Visible = isCheckedValue;
            Properties.Settings.Default.Group4 = isCheckedValue;
            Properties.Settings.Default.Save();
            }

        private void checkBox4_Click(object sender, System.Windows.RoutedEventArgs e)
            {
            // 获取CheckBox的布尔值
            bool isCheckedValue = checkBox4.IsChecked.HasValue && checkBox4.IsChecked.Value;
            Globals.Ribbons.Ribbon1.group5.Visible = isCheckedValue;
            Properties.Settings.Default.Group5 = isCheckedValue;
            Properties.Settings.Default.Save();
            }

        private void checkBox6_Click(object sender, System.Windows.RoutedEventArgs e)
            {
            // 获取CheckBox的布尔值
            bool isCheckedValue = checkBox6.IsChecked.HasValue && checkBox6.IsChecked.Value;
            Globals.Ribbons.Ribbon1.tab2.Visible = isCheckedValue;
            Properties.Settings.Default.Group6 = isCheckedValue;
            Properties.Settings.Default.Save();
            }



        private void ColorPicker_Confirmed(object sender, HandyControl.Data.FunctionEventArgs<System.Windows.Media.Color> e)
            {
            var selectedColor = ColorPicker.SelectedBrush.Color; //选择颜色转为rgb
            int alpha = selectedColor.A;
            int red = selectedColor.R;
            int green = selectedColor.G;
            int blue = selectedColor.B;
            Color iconColor = Color.FromArgb(alpha, red, green, blue);
            // 保存选择的颜色到应用程序设置
            Properties.Settings.Default.iconColor = iconColor;
            Properties.Settings.Default.Save();
            RibbonIcoHelper class_GetIcons = new RibbonIcoHelper();
            class_GetIcons.GetIcons("PPT");
            // Growl.Success("更新成功");
            }

        private void ColorPicker_Canceled(object sender, System.EventArgs e)
            {
            this.Close();
            }



        private void GraImages_MouseDown(object sender, System.Windows.Input.MouseButtonEventArgs e)
            {
            DrawerRight.IsOpen = true;
            }

        private void CopyShieId_Click(object sender, RoutedEventArgs e)
            {
            DrawerRight.IsOpen = true;
            }
        }
    }
