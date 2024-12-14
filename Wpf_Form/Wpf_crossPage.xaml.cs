using Microsoft.Office.Interop.PowerPoint;
using System;
using System.IO;
using System.Windows;
using System.Windows.Media.Imaging;
using Path = System.IO.Path;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace PresPio
    {
    /// <summary>
    /// Wpf_crossPage.xaml 的交互逻辑
    /// </summary>
    public partial class Wpf_crossPage
        {

        public PowerPoint.Application app { get; set; }
        public Wpf_crossPage()
            {
            app = Globals.ThisAddIn.Application;
            InitializeComponent();
            // 将事件处理程序添加到 WindowSelectionChange 事件
            app.WindowSelectionChange += new EApplication_WindowSelectionChangeEventHandler(Application_ShapeSelectionChange);
            //形状改变大小事件事件
            app.AfterShapeSizeChange += new Microsoft.Office.Interop.PowerPoint.EApplication_AfterShapeSizeChangeEventHandler(Application_AfterShapeSizeChange);


            Presentation pre = app.ActivePresentation;
            int num = pre.Slides.Count;
            int num1 = app.ActiveWindow.Selection.SlideRange.SlideIndex;
            NumericUpDown1.Minimum = 1;
            NumericUpDown1.Maximum = num;
            NumericUpDown1.Value = num1;
            NumericUpDown2.Maximum = num;//定义最大数字为幻灯片数
            NumericUpDown2.Minimum = 1;
            NumericUpDown2.Value = num;
            }
      public void Application_ShapeSelectionChange(Microsoft.Office.Interop.PowerPoint.Selection Sel)
            {
            LoadImg();
            }

        private void Application_AfterShapeSizeChange(Microsoft.Office.Interop.PowerPoint.Shape Shape)
            {
            LoadImg();

            }
        private void CrossWindow_Loaded(object sender, RoutedEventArgs e)
            {
            LoadImg();
            }

        public void LoadImg()
            {
            try
                {
                Selection sel = app.ActiveWindow.Selection;
                Slide slide = app.ActiveWindow.View.Slide;
                if (sel.Type == PpSelectionType.ppSelectionShapes)
                    {
                    PowerPoint.ShapeRange shr = sel.ShapeRange;
                    string tempFolderName = Guid.NewGuid().ToString("N");
                    string oPath_Dir = System.IO.Path.Combine(Path.GetTempPath(), "MyAppTemp", tempFolderName);
                    Directory.CreateDirectory(oPath_Dir);
                    int count = sel.ShapeRange.Count;
                    if (count == 1)
                        {
                        string oPath_Full = Path.Combine(oPath_Dir, "temp_" + Guid.NewGuid().ToString("N") + ".png");
                        Properties.Settings.Default.PicPath = oPath_Full;
                        Properties.Settings.Default.Save();
                        shr[count].Export(oPath_Full, PpShapeFormat.ppShapeFormatPNG, 0, 0);
                        if (File.Exists(oPath_Full))
                            {
                            // 加载图片
                            BitmapImage bitmapImage = new BitmapImage();
                            bitmapImage.BeginInit();
                            bitmapImage.UriSource = new Uri(oPath_Full, UriKind.RelativeOrAbsolute);
                            bitmapImage.CacheOption = BitmapCacheOption.OnLoad;
                            bitmapImage.EndInit();

                            // 将图片路径设置为 Image 控件的资源
                            imageBox.Source = bitmapImage;

                            }
                        if (File.Exists(oPath_Full))
                            {
                            File.Delete(oPath_Full);//删除临时文件
                            
                            }
                        }
                    if (Directory.Exists(oPath_Dir))
                        {
                        Directory.Delete(oPath_Dir, true);//删除临时文件夹及其中的文件
                        }
                    }
                else
                    {
                    // System.Windows.Forms.MessageBox.Show("请在PPT中选择图片内容或加载本地图片", "温馨提示");
                    }
                }
            catch
                {
                return;
                }
            }

        private void CopBtn_Click(object sender, RoutedEventArgs e)
            {
            try
                {
                Selection sel = app.ActiveWindow.Selection;
                Presentation pre = app.ActivePresentation;
                int num1 = (int)NumericUpDown1.Value;
                int num2 = (int)NumericUpDown2.Value;
                sel.ShapeRange.Copy();
                if (sel.Type != PpSelectionType.ppSelectionShapes || sel.Type != PpSelectionType.ppSelectionText)
                    {
                    for (int i = num1 ; i <= num2 ; i++)
                        {
                        int num = i + 1;
                        pre.Slides.Range(num).Shapes.Paste().Name = "批量复制";

                        }
                  
                    }
                else
                    {
                    MessageBox.Show("请选择单一内容后再试！");
                    }
                }
            catch
                {
                MessageBox.Show("请选择单一内容后再试！");
                }
            }

        private void DelBtn_Click(object sender, RoutedEventArgs e)
            {
            try
                {
                Slide slide = app.ActiveWindow.View.Slide;
                Presentation pre = app.ActiveWindow.Presentation;
                Selection sel = app.ActiveWindow.Selection;
                if (sel.Type != PpSelectionType.ppSelectionShapes || sel.Type != PpSelectionType.ppSelectionText)
                    {
                    float myTop = sel.ShapeRange.Top;
                    float myLeft = sel.ShapeRange.Left;
                    float myWidh = sel.ShapeRange.Width;
                    float myHeight = sel.ShapeRange.Height;//获取所选的位置

                    foreach (Slide oSlide in pre.Slides)
                        {
                        foreach (PowerPoint.Shape oshape in oSlide.Shapes)
                            {

                            float B1 = Math.Abs(myTop - oshape.Top);
                            float B2 = Math.Abs(myLeft - oshape.Left);
                            float B3 = Math.Abs(myWidh - oshape.Width);
                            float B4 = Math.Abs(myHeight - oshape.Height);
                            bool C1 = B1 < 1;
                            bool C2 = B2 < 1;
                            bool C3 = B3 < 1;
                            bool C4 = B4 < 1;
                            if (C1 && C2 && C3 && C4)
                                {
                                oshape.Delete();
                                }
                            }
                        }
                    }
                else
                    {
                    System.Windows.MessageBox.Show("请选择单一内容后操作", "温馨提示");
                    }

                }
            catch
                {
                return;
                }
            System.Windows.MessageBox.Show("删除成功！");
            }
        }
    }
