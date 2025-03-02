using System;
using System.Windows;
using System.Windows.Documents;
using System.Windows.Media;
using HandyControl.Controls;
using Color = System.Drawing.Color;
using MediaColor = System.Windows.Media.Color;

namespace PresPio
    {
    public partial class Wpf_About
        {
        private Color currentIconColor;

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

            // 初始化高级设置
            AutoSizeCheckBox.IsChecked = Properties.Settings.Default.AutoSize;
            AllowEditCheckBox.IsChecked = Properties.Settings.Default.AllowEdit;
            ApplyToMasterCheckBox.IsChecked = Properties.Settings.Default.ApplyToMaster;

            // 初始化图标颜色
            currentIconColor = Properties.Settings.Default.iconColor;
            UpdateColorPreview(currentIconColor);
            }

        private void UpdateColorPreview(Color color)
            {
            // 更新色块预览
            if (MenuColorPreview != null)
                {
                var brush = new SolidColorBrush(MediaColor.FromArgb(
                    color.A,
                    color.R,
                    color.G,
                    color.B));
                MenuColorPreview.Background = brush;
                }
            }

        private void About_Windown_Loaded(object sender, RoutedEventArgs e)
            {
            AboutText();
            }

        private void AboutText()
            {
            FlowDocument doc = new FlowDocument();

            // 添加标题
            Paragraph titlePara = new Paragraph(new Run("关于 PresPio"))
                {
                FontSize = 20,
                FontWeight = FontWeights.Bold,
                Margin = new Thickness(0, 0, 0, 10)
                };
            doc.Blocks.Add(titlePara);

            // 添加简
            Paragraph introPara = new Paragraph()
                {
                Margin = new Thickness(0, 0, 0, 15)
                };
            introPara.Inlines.Add(new Run("PresPio 是一款专业的 PowerPoint 插件工具，旨在提升演示文稿的制作效率和质量。"));
            doc.Blocks.Add(introPara);

            // 添加功能特点
            Paragraph featureTitle = new Paragraph(new Run("主要功能："))
                {
                FontSize = 16,
                FontWeight = FontWeights.Bold,
                Margin = new Thickness(0, 0, 0, 10)
                };
            doc.Blocks.Add(featureTitle);

            List featureList = new List();
            featureList.MarkerStyle = TextMarkerStyle.Disc;
            featureList.ListItems.Add(new ListItem(new Paragraph(new Run("颜色组：快速调整和统一演示文稿配色"))));
            featureList.ListItems.Add(new ListItem(new Paragraph(new Run("文字组：提供丰富的文字排版和样式工具"))));
            featureList.ListItems.Add(new ListItem(new Paragraph(new Run("图形组：内置多种精美图形和图表模板"))));
            featureList.ListItems.Add(new ListItem(new Paragraph(new Run("工具组：实用的辅助工具集��"))));
            featureList.ListItems.Add(new ListItem(new Paragraph(new Run("实验组：创新功能的测试区域"))));
            doc.Blocks.Add(featureList);

            // 添加版权信息
            Paragraph copyrightPara = new Paragraph()
                {
                Margin = new Thickness(0, 15, 0, 0)
                };
            copyrightPara.Inlines.Add(new Run("© 2024 PresPio. 保留所有权利。"));
            doc.Blocks.Add(copyrightPara);

            myRichTextBox.Document = doc;
            }

        private void checkBox1_Click(object sender, RoutedEventArgs e)
            {
            Properties.Settings.Default.Group2 = checkBox1.IsChecked ?? false;
            Properties.Settings.Default.Save();
            }

        private void checkBox2_Click(object sender, RoutedEventArgs e)
            {
            Properties.Settings.Default.Group3 = checkBox2.IsChecked ?? false;
            Properties.Settings.Default.Save();
            }

        private void checkBox3_Click(object sender, RoutedEventArgs e)
            {
            Properties.Settings.Default.Group4 = checkBox3.IsChecked ?? false;
            Properties.Settings.Default.Save();
            }

        private void checkBox4_Click(object sender, RoutedEventArgs e)
            {
            Properties.Settings.Default.Group5 = checkBox4.IsChecked ?? false;
            Properties.Settings.Default.Save();
            }

        private void checkBox6_Click(object sender, RoutedEventArgs e)
            {
            Properties.Settings.Default.Group7 = checkBox6.IsChecked ?? false;
            Properties.Settings.Default.Save();
            }

        private void checkBox7_Click(object sender, RoutedEventArgs e)
            {
            Properties.Settings.Default.AutoSize = AutoSizeCheckBox.IsChecked ?? false;
            Properties.Settings.Default.Save();
            }

        private void checkBox8_Click(object sender, RoutedEventArgs e)
            {
            Properties.Settings.Default.AllowEdit = AllowEditCheckBox.IsChecked ?? false;
            Properties.Settings.Default.Save();
            }

        private void checkBox9_Click(object sender, RoutedEventArgs e)
            {
            Properties.Settings.Default.ApplyToMaster = ApplyToMasterCheckBox.IsChecked ?? false;
            Properties.Settings.Default.Save();
            }

        private void ColorPicker_Confirmed(object sender, EventArgs e)
            {
            var selectedColor = MenuColorPicker.SelectedBrush.Color;
            Color iconColor = Color.FromArgb(
                selectedColor.A,
                selectedColor.R,
                selectedColor.G,
                selectedColor.B);

            Properties.Settings.Default.iconColor = iconColor;
            Properties.Settings.Default.Save();

            currentIconColor = iconColor;
            UpdateColorPreview(currentIconColor);

            RibbonIcoHelper class_GetIcons = new RibbonIcoHelper();
            class_GetIcons.GetIcons("PPT");

            MenuColorPickerPopup.IsOpen = false;
            Growl.Success("菜单颜色已更新");
            }

        private void ColorPicker_Canceled(object sender, EventArgs e)
            {
            MenuColorPickerPopup.IsOpen = false;
            }

        private void MenuColorPreview_MouseDown(object sender, System.Windows.Input.MouseButtonEventArgs e)
            {
            MenuColorPickerPopup.IsOpen = true;
            e.Handled = true;
            }
        }
    }