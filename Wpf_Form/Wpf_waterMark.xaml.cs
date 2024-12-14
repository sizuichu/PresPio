using HandyControl.Controls;
using System;
using System.Collections.Generic;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Forms;
using System.Windows.Input;
using System.Windows.Media;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using System.Drawing;
using System.Threading.Tasks;
using System.Diagnostics;
using System.Linq;
using ShapeRange = Microsoft.Office.Interop.PowerPoint.ShapeRange;
using Shape = Microsoft.Office.Interop.PowerPoint.Shape;
using System.Runtime.InteropServices;

namespace PresPio
    {
    /// <summary>
    /// Wpf_waterMark.xaml 的交互逻辑
    /// </summary>
    public partial class Wpf_waterMark
        {
        public PowerPoint.Application app;
        public Wpf_waterMark()
            {
            InitializeComponent();
            }


        private void markWindow_Loaded(object sender, RoutedEventArgs e)
            {

            // 获取当前应用程序实例
            var app = Globals.ThisAddIn.Application;

            // 从应用程序设置中获取水印文本
            string waterText = Properties.Settings.Default.WaterText;

            // 将水印文本分配给文本框和标签内容
            FontTextBox.Text = waterText;
            markLabel.Content = waterText;
            Watermark.Mark = waterText;
            // 从应用程序设置中获取字体、颜色和大小
            string waterFont = Properties.Settings.Default.WaterFont;
            System.Drawing.Color drawingColor = Properties.Settings.Default.WaterColor;
            double waterSize = Properties.Settings.Default.WaterSize;
            Watermark.FontSize = waterSize / 2;

            Num1.Value = Properties.Settings.Default.WaterRow;
            Num2.Value = Properties.Settings.Default.WaterColumn;
            // 转换颜色为WPF的MediaColor
            System.Windows.Media.Color mediaColor = System.Windows.Media.Color.FromArgb(drawingColor.A, drawingColor.R, drawingColor.G, drawingColor.B);

            // 应用字体、颜色和大小到标签
            markLabel.FontFamily = new System.Windows.Media.FontFamily(waterFont);
            markLabel.FontSize = waterSize;
            markLabel.Foreground = new SolidColorBrush(mediaColor);

            // 将水印大小显示为整数
            Fontsize.Content = (int)waterSize;

            }

        private void FontBtn_Click(object sender, RoutedEventArgs e)
            {
            FontDialog fontDialog = new FontDialog();
            fontDialog.Font = new System.Drawing.Font(Properties.Settings.Default.WaterFont, Properties.Settings.Default.WaterSize);

            if (fontDialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                // 更新界面字体
                markLabel.FontSize = fontDialog.Font.Size;
                markLabel.FontFamily = new System.Windows.Media.FontFamily(fontDialog.Font.Name);
                Fontsize.Content = (int)fontDialog.Font.Size;
                Watermark.FontSize = (int)fontDialog.Font.Size / 2;
                // 保存用户选择到应用程序设置
                Properties.Settings.Default.WaterFont = fontDialog.Font.Name;
                Properties.Settings.Default.Save();
                Properties.Settings.Default.WaterSize = fontDialog.Font.Size;
                Properties.Settings.Default.WaterFonts = fontDialog.Font.Style;
                Properties.Settings.Default.Save();
                }
            else
                {
                // 用户取消选择
                return;
                }

            }


        private void FontTextBox_TextChanged(object sender, TextChangedEventArgs e)
            {
            markLabel.Content = FontTextBox.Text;
            Watermark.Mark = FontTextBox.Text;
            Properties.Settings.Default.WaterText = FontTextBox.Text;
            Properties.Settings.Default.Save();

            }

        private void markLabel_MouseDoubleClick(object sender, MouseButtonEventArgs e)
            {
            colorPicker.Visibility = System.Windows.Visibility.Visible;
            }

        private void colorPicker_Canceled(object sender, EventArgs e)
            {
            colorPicker.Visibility = System.Windows.Visibility.Collapsed;
            }

        private void colorPicker_Confirmed(object sender, HandyControl.Data.FunctionEventArgs<System.Windows.Media.Color> e)
            {

            var solidColorBrush = colorPicker.SelectedBrush;

            // 获取SolidColorBrush的颜色
            System.Windows.Media.Color color = solidColorBrush.Color;

            // 设置标签的前景色为获取的颜色
            markLabel.Foreground = new SolidColorBrush(color);

            // 隐藏颜色选择器
            colorPicker.Visibility = System.Windows.Visibility.Collapsed;

            // 将 System.Windows.Media.Color 转换为 System.Drawing.Color
            System.Drawing.Color drawingColor = System.Drawing.Color.FromArgb(color.A, color.R, color.G, color.B);

            // 将 System.Drawing.Color 保存到应用程序设置中
            Properties.Settings.Default.WaterColor = drawingColor;
            Properties.Settings.Default.Save();
            }

        private void Num1_ValueChanged(object sender, HandyControl.Data.FunctionEventArgs<double> e)
            {
            Properties.Settings.Default.WaterRow = (int)Num1.Value;
            Properties.Settings.Default.Save();

            }

        private void Num2_ValueChanged(object sender, HandyControl.Data.FunctionEventArgs<double> e)
            {
            Properties.Settings.Default.WaterColumn = (int)Num2.Value;
            Properties.Settings.Default.Save();
            }

        /// <summary>
        /// 删除水印
        /// </summary>
        /// <returns></returns>
        public void DelwaterrMark()
            {
            app = Globals.ThisAddIn.Application;
            Slide slide = app.ActiveWindow.View.Slide;
            Presentation pre = app.ActivePresentation;

            if ((System.Windows.Forms.Control.ModifierKeys & Keys.Control) == Keys.Control)  //判断Ctrl键
                {
                int count = slide.Master.Shapes.Count;
                for (int i = count ; i >= 1 ; i--)
                    {
                    if (slide.Master.Shapes[i].Name == "超级水印")
                        {
                        slide.Master.Shapes[i].Delete();
                        Marshal.ReleaseComObject(slide.Master.Shapes[i]); // 释放资源
                        }
                    }
                }
            else
                {
                int slideCount = pre.Slides.Count;
                for (int i = slideCount ; i >= 1 ; i--)
                    {
                    int shapeCount = pre.Slides[i].Shapes.Count;
                    for (int j = shapeCount ; j >= 1 ; j--)
                        {
                        PowerPoint.Shape shp = pre.Slides[i].Shapes[j];
                        if (shp.Name == "超级水印")
                            {
                            shp.Delete();
                            Marshal.ReleaseComObject(shp); // 释放资源
                            }
                        }
                    }
                }

            // 显式调用垃圾回收器
            GC.Collect();
            GC.WaitForPendingFinalizers();
            }
        private void Button_Click(object sender, RoutedEventArgs e)
            {
            DelwaterrMark();

            Growl.Success("水印删除成功！");
            }
        private List<string> shps = new List<string>();
        private void Button_Click_1(object sender, RoutedEventArgs e)
            {
            app = Globals.ThisAddIn.Application;
            DelwaterrMark();
            string FontName = Properties.Settings.Default.WaterFont;
            float FontSize = Properties.Settings.Default.WaterSize;
            string FontText = Properties.Settings.Default.WaterText;
            Slide slide = app.ActiveWindow.View.Slide;
            Presentation pre = app.ActivePresentation;
            List<Shape> shps = new List<Shape>();
            int n = Properties.Settings.Default.WaterRow;
            int m = Properties.Settings.Default.WaterColumn;
            float Width = pre.PageSetup.SlideWidth + 50;
            float Height = pre.PageSetup.SlideHeight + 30;
            float Wcell = Width / n;
            float Hcell = Height / m;
            System.Drawing.Color WaterColor = Properties.Settings.Default.WaterColor;
            Microsoft.Office.Interop.PowerPoint.Shape shp;

            try
                {
                for (int i = 0 ; i < n ; i++)
                    {
                    for (int j = 0 ; j < m ; j++)
                        {
                        shp = slide.Shapes.AddTextbox(MsoTextOrientation.msoTextOrientationHorizontalRotatedFarEast, Wcell * i - 10, Hcell * j + 10, 100, 30);
                        shp.TextFrame.TextRange.Text = FontText;//水印内容
                        shp.TextFrame.TextRange.Font.NameFarEast = FontName;
                        shp.TextFrame.TextRange.Font.NameOther = FontName;//字体名称
                        shp.TextFrame.TextRange.Font.Size = FontSize;//字体大小
                        shp.TextFrame.TextRange.Font.Color.RGB = RGB2Int(WaterColor.R, WaterColor.G, WaterColor.B);
                        shp.TextFrame.TextRange.Font.Color.Brightness = 1f - (WaterColor.A / 255f);
                        shp.Rotation = -45;//旋转角度
                        shps.Add(shp);
                        }
                    }

                // Group shapes
                ShapeRange shpRange = slide.Shapes.Range(shps.Select(s => s.Name).ToArray());
                Shape shp2 = shpRange.Group();
                float slideCenterX = (Globals.ThisAddIn.Application.ActivePresentation.PageSetup.SlideWidth - shp2.Width) / 2;
                float slideCenterY = (Globals.ThisAddIn.Application.ActivePresentation.PageSetup.SlideHeight - shp2.Height) / 2;
                shp2.Left = slideCenterX;
                shp2.Top = slideCenterY;

                // Paste shapes
                if ((System.Windows.Forms.Control.ModifierKeys & Keys.Control) == Keys.Control)  //判断Ctrl键
                    {
                    shp2.Copy();
                    ShapeRange shr = slide.Master.Shapes.Paste();
                    shr.Name = "超级水印";
                    }
                else
                    {
                    shp2.Copy();
                    foreach (PowerPoint.Slide item in app.ActiveWindow.Selection.SlideRange)
                        {
                        ShapeRange shr = item.Shapes.Paste();
                        shr.Name = "超级水印";
                        }
                    }

                shp2.Delete();
                }
            catch (Exception ex)
                {
                // Handle exception
                Debug.WriteLine(ex.Message);
                }



            }


        /// <summary>
        /// RGB转换为INT色值
        /// </summary>
        /// <param name="R"></param>
        /// <param name="G"></param>
        /// <param name="B"></param>
        /// <returns></returns>
        public int RGB2Int(int R, int G, int B)
            {
            int PPTRGB = R + G * 256 + B * 256 * 256;
            return PPTRGB;
            }
        public System.Drawing.Color Int2RGB(int color)
            {
            int B = color / (256 * 256);
            int G = (color - B * 256 * 256) / 256;
            int R = color - B * 256 * 256 - G * 256;
            return System.Drawing.Color.FromArgb(R, G, B);
            }
        }
    }