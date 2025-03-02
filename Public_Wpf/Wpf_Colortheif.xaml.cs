using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using ColorThiefDotNet;
using HandyControl.Controls;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using Newtonsoft.Json;
using Brush = System.Windows.Media.Brush;
using Brushes = System.Windows.Media.Brushes;
using Color = System.Windows.Media.Color;
using OpenFileDialog = System.Windows.Forms.OpenFileDialog;
using Path = System.IO.Path;
using Pen = System.Windows.Media.Pen;
using Point = System.Windows.Point;

using PowerPoint = Microsoft.Office.Interop.PowerPoint;

using SaveFileDialog = System.Windows.Forms.SaveFileDialog;

namespace PresPio
    {
    /// <summary>
    /// Wpf_Colortheif.xaml 的交互逻辑
    /// </summary>
    public partial class Wpf_Colortheif
        {
        public PowerPoint.Application app = Globals.ThisAddIn.Application; //加载PPT项目

        public Wpf_Colortheif()
            {
            InitializeComponent();
            SetToggleButtonEvents(); //设置按钮事件
            GetThemeBtn();//获取主题色

            //主题改变事件
            app = Globals.ThisAddIn.Application;
            (app as EApplication_Event).ColorSchemeChanged += ThisAddIn_ColorSchemeChanged;
            }

        private void ThisAddIn_ColorSchemeChanged(SlideRange SldRange)
            {
            GetThemeBtn();//获取主题色
            }

        private void WinDown_Loaded(object sender, RoutedEventArgs e)
            {
            Selection sel = app.ActiveWindow.Selection;
            if (sel.Type == PpSelectionType.ppSelectionShapes)
                {
                loadImgBtn_Click(sender, e);
                }
            }

        /// <summary>
        /// 颜色转换
        /// </summary>
        /// <param name="color"></param>
        /// <returns></returns>
        public Brush ConvertColorToBrush(System.Windows.Media.Color color)
            {
            return new SolidColorBrush(color);
            }

        /// <summary>
        /// 装换本地文件为 BitmapFrame
        /// </summary>
        /// <param name="imagePath"></param>
        /// <returns></returns>
        public BitmapFrame ConvertImageToBitmapFrame(string imagePath)
            {
            using (FileStream stream = new FileStream(imagePath, FileMode.Open, FileAccess.Read))
                {
                BitmapDecoder decoder = BitmapDecoder.Create(stream, BitmapCreateOptions.PreservePixelFormat, BitmapCacheOption.OnLoad);
                return decoder.Frames[0];
                }
            }

        private void Button_Click(object sender, RoutedEventArgs e)
            {
            var openFileDialog = new System.Windows.Forms.OpenFileDialog();
            openFileDialog.Filter = "图片文件 (*.jpg;*.jpeg;*.png;*.bmp;*.gif;*.tiff)|*.jpg;*.jpeg;*.png;*.bmp;*.gif;*.tiff";
            if (openFileDialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                string fileName = openFileDialog.FileName;
                var img = ConvertImageToBitmapFrame(fileName);
                ImageViewer.ImageSource = img;
                colorButton(fileName);
                }
            else
                {
                //Growl.Warning("未选择图片");
                }
            }

        private void loadImgBtn_Click(object sender, RoutedEventArgs e)
            {
            try
                {
                Selection sel = app.ActiveWindow.Selection;
                Slide slide = app.ActiveWindow.View.Slide;
                if (sel.Type == PpSelectionType.ppSelectionShapes)
                    {
                    PowerPoint.ShapeRange shr = sel.ShapeRange;
                    string tempFolderName = Guid.NewGuid().ToString("N");
                    string oPath_Dir = Path.Combine(Path.GetTempPath(), "MyAppTemp", tempFolderName);
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
                            var img = ConvertImageToBitmapFrame(oPath_Full);
                            ImageViewer.ImageSource = img;
                            //pictureBox1.Image = Image.FromFile(oPath_Full);
                            colorButton(oPath_Full);
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
                    System.Windows.Forms.MessageBox.Show("请在PPT中选择图片内容或加载本地图片", "温馨提示");
                    }
                }
            catch
                {
                return;
                }
            }

        private void ColorPicker_SelectedColorChanged(object sender, HandyControl.Data.FunctionEventArgs<System.Windows.Media.Color> e)
            {
            if (_currentColorButton != null)
                {
                // 实时更新按钮颜色
                _currentColorButton.Background = new SolidColorBrush(e.Info);
                }
            }

        // 颜色选择器确认按钮点击事件
        private void ColorPicker_Confirmed(object sender, EventArgs e)
            {
            if (_currentColorButton != null)
                {
                // 更新按钮颜色
                _currentColorButton.Background = ColorPicker.SelectedBrush;
                // 关闭颜色选择器
                ColorPickerPopup.IsOpen = false;
                // ���消按钮选中状态
                _currentColorButton.IsChecked = false;
                _currentColorButton = null;
                }
            }

        // 颜色选择器Popup关闭事件
        private void ColorPickerPopup_Closed(object sender, EventArgs e)
            {
            if (_currentColorButton != null)
                {
                // 如果是点击外部关闭，恢复原来的颜色
                _currentColorButton.Background = ColorPicker.SelectedBrush;
                _currentColorButton.IsChecked = false;
                _currentColorButton = null;
                }
            }

        // ToggleButton点击事件
        private void ColorButton_Click(object sender, RoutedEventArgs e)
            {
            var colorButton = sender as ToggleButton;
            if (colorButton != null)
                {
                // 取消其他按钮的选中状态
                var parent = VisualTreeHelper.GetParent(colorButton);
                while (!(parent is UniformGrid) && parent != null)
                    {
                    parent = VisualTreeHelper.GetParent(parent);
                    }

                if (parent is UniformGrid grid)
                    {
                    foreach (var child in grid.Children.OfType<ToggleButton>())
                        {
                        if (child != colorButton)
                            {
                            child.IsChecked = false;
                            }
                        }
                    }
                }

            // 显示颜色选择器
            if (colorButton.IsChecked == true)
                {
                // 保存当前选中的按钮
                _currentColorButton = colorButton;
                // 设置颜色选择器的初始颜色
                ColorPicker.SelectedBrush = (SolidColorBrush)colorButton.Background;
                ColorPickerPopup.IsOpen = true;
                }
            else
                {
                // 关闭颜色选择器
                ColorPickerPopup.IsOpen = false;
                _currentColorButton = null;
                }
            }

        // 添加字段保存当前选中的按钮
        private ToggleButton _currentColorButton;

        public System.Windows.Media.Color ConvertBrushToColor(Brush brush)
            {
            if (brush is SolidColorBrush solidColorBrush)
                {
                return solidColorBrush.Color;
                }
            else if (brush is GradientBrush gradientBrush)
                {
                // 这里可以根据需要选择渐变中的颜色，例如取渐变的起始颜色
                return gradientBrush.GradientStops[0].Color;
                }
            else if (brush is LinearGradientBrush linearGradientBrush)
                {
                // 这里可以根据需要选择渐变中的颜色，例如取渐变的起始颜色
                return linearGradientBrush.GradientStops[0].Color;
                }
            else if (brush is RadialGradientBrush radialGradientBrush)
                {
                // 这里可以根据需要选择渐变中的颜色，例如取渐变的起始颜色
                return radialGradientBrush.GradientStops[0].Color;
                }
            else
                {
                throw new ArgumentException("The provided brush is not a supported type.");
                }
            }

        //按钮共同事件
        private void ToggleButton_Click(object sender, RoutedEventArgs e)
            {
            ToggleButton[] toggleButtons = {
                      ToggleButton1,
                ToggleButton2,
                ToggleButton3,
                ToggleButton4,
                ToggleButton5,
                ToggleButton6,
                ToggleButton7,
                ToggleButton8,
                ToggleButton9,
                ToggleButton10,
                ToggleButton11,
                ToggleButton12,
                ToggleButton13,
                ToggleButton14,
                ToggleButton15,
                ToggleButton16,
                ToggleButton17,
                ToggleButton18,
                ToggleButton19,
                ToggleButton20,
                ToggleButton21,
                ToggleButton22,
                ToggleButton23,
                ToggleButton24,
                ToggleButton25,
                ToggleButton26,
                ToggleButton27,
                ToggleButton28,
                };
            ToggleButton toggleButton = sender as ToggleButton;
            foreach (var item in toggleButtons)
                {
                if (item.Name != toggleButton.Name)
                    {
                    item.IsChecked = false;
                    }
                }
            }

        //通用按钮事件
        private void SetToggleButtonEvents()
            {
            ToggleButton[] toggleButtons = {
                ToggleButton1,
                ToggleButton2,
                ToggleButton3,
                ToggleButton4,
                ToggleButton5,
                ToggleButton6,
                ToggleButton7,
                ToggleButton8,
                ToggleButton9,
                ToggleButton10,
                ToggleButton11,
                ToggleButton12,
                ToggleButton13,
                ToggleButton14,
                ToggleButton15,
                ToggleButton16,
                ToggleButton17,
                ToggleButton18,
                ToggleButton19,
                ToggleButton20,
                ToggleButton21,
                ToggleButton22,
                ToggleButton23,
                ToggleButton24,
                ToggleButton25,
                ToggleButton26,
                ToggleButton27,
                ToggleButton28,
              };
            foreach (ToggleButton toggleButton in toggleButtons)
                {
                toggleButton.Click += ToggleButton_Click;
                }
            }

        //获取颜色按钮
        public static System.Windows.Media.Color ConvertDrawingColorToMediaColor(System.Drawing.Color drawingColor)
            {
            // 提取 ARGB 分量
            byte a = drawingColor.A;
            byte r = drawingColor.R;
            byte g = drawingColor.G;
            byte b = drawingColor.B;

            // 创建新的 System.Windows.Media.Color 对象
            System.Windows.Media.Color mediaColor = System.Windows.Media.Color.FromArgb(a, r, g, b);

            return mediaColor;
            }

        public SolidColorBrush ConvertToWpfColor(System.Drawing.Color drawingColor)
            {
            return new SolidColorBrush(System.Windows.Media.Color.FromArgb(drawingColor.A, drawingColor.R, drawingColor.G, drawingColor.B));
            }

        public void colorButton(string PicPath)
            {
            ToggleButton2.Background = ConvertToWpfColor(GetColor1(PicPath));
            ToggleButton1.Background = ConvertToWpfColor(GetColor2(PicPath)[0]);
            ToggleButton3.Background = ConvertToWpfColor(GetColor2(PicPath)[1]);
            ToggleButton4.Background = ConvertToWpfColor(GetColor2(PicPath)[2]);
            ToggleButton5.Background = ConvertToWpfColor(GetColor2(PicPath)[3]);
            ToggleButton6.Background = ConvertToWpfColor(GetColor2(PicPath)[4]);
            ToggleButton7.Background = ConvertToWpfColor(GetColor2(PicPath)[5]);
            ToggleButton8.Background = ConvertToWpfColor(GetColor2(PicPath)[6]);
            ToggleButton9.Background = ConvertToWpfColor(GetColor2(PicPath)[7]);
            ToggleButton10.Background = ConvertToWpfColor(GetColor2(PicPath)[8]);
            ToggleButton11.Background = ConvertToWpfColor(GetColor2(PicPath)[9]);
            ToggleButton12.Background = ConvertToWpfColor(GetColor2(PicPath)[10]);
            ToggleButton13.Background = ConvertToWpfColor(GetColor2(PicPath)[11]);
            ToggleButton14.Background = ConvertToWpfColor(GetColor2(PicPath)[12]);
            ToggleButton15.Background = ConvertToWpfColor(GetColor2(PicPath)[13]);
            ToggleButton16.Background = ConvertToWpfColor(GetColor2(PicPath)[14]);
            }

        //获得主色的函数
        public System.Drawing.Color GetColor1(string PicPath)
            {
            using (var bitmap = new Bitmap(PicPath))
                {
                var colorThief = new ColorThief();
                var imgColor = colorThief.GetColor(bitmap, 9, true);
                return System.Drawing.Color.FromArgb(imgColor.Color.A, imgColor.Color.R, imgColor.Color.G, imgColor.Color.B);
                }
            }

        //获得主色的函数
        public List<System.Drawing.Color> GetColor2(string picPath)
            {
            using (var bitmap = new Bitmap(picPath))
                {
                var colorThief = new ColorThief();
                var newColors = colorThief.GetPalette(bitmap, 20);
                var result = new List<System.Drawing.Color>(newColors.Count);
                foreach (var color in newColors)
                    {
                    result.Add(System.Drawing.Color.FromArgb(color.Color.A, color.Color.R, color.Color.G, color.Color.B));
                    }
                return result;
                }
            }

        public static BitmapFrame ConvertDrawingBitmapToBitmapFrame(System.Drawing.Bitmap bitmap)
            {
            // 将 System.Drawing.Bitmap 保存为临时文件
            string tempFilePath = Path.GetTempFileName() + ".png";
            bitmap.Save(tempFilePath, System.Drawing.Imaging.ImageFormat.Png);

            // 从临时文件中加载图像
            BitmapImage bitmapImage = new BitmapImage();
            bitmapImage.BeginInit();
            bitmapImage.UriSource = new Uri(tempFilePath);
            bitmapImage.EndInit();

            // 从 BitmapImage 中获取 BitmapFrame
            return BitmapFrame.Create(bitmapImage);
            }

        //获得色块的函数
        public List<int> GetColor3(string PicPath)
            {
            PicPath = Properties.Settings.Default.PicPath;
            Bitmap bitmap = new Bitmap(PicPath);
            ImageViewer.ImageSource = ConvertDrawingBitmapToBitmapFrame(bitmap);
            var colorThief = new ColorThief();
            List<QuantizedColor> NewColor = colorThief.GetPalette(bitmap, 8);
            List<int> Scolor = new List<int>();
            foreach (QuantizedColor color in NewColor)
                {
                int A = color.Color.A;
                int R = color.Color.R;
                int G = color.Color.G * 256;
                int B = color.Color.B * 256 * 256;
                int newcolor = R + G + B;
                Scolor.Add(newcolor);
                }
            return Scolor;
            }

        public void DelShpe(string Name)
            {
            Slide slide = app.ActiveWindow.View.Slide;
            int shapesCount = slide.Shapes.Count;
            for (int i = shapesCount ; i > 0 ; i--)
                {
                PowerPoint.Shape shape = slide.Shapes[i];
                if (shape.Tags["配色"] == Name)
                    {
                    shape.Delete();
                    }
                }
            }

        public System.Drawing.Color BrushToColor(Brush brush)
            {
            SolidColorBrush solidBrush = brush as SolidColorBrush;
            if (solidBrush != null)
                {
                return System.Drawing.Color.FromArgb(
                    solidBrush.Color.A,
                    solidBrush.Color.R,
                    solidBrush.Color.G,
                    solidBrush.Color.B);
                }
            else
                {
                throw new InvalidOperationException("The provided brush is not a SolidColorBrush.");
                }
            }

        private void Button_Click_1(object sender, RoutedEventArgs e)
            {
            DelShpe("快速配色");
            try
                {
                Slide slide = app.ActiveWindow.View.Slide;
                float left = -50;
                string PicPath = Properties.Settings.Default.PicPath;

                // Create the shapes
                var shapes = Enumerable.Range(0, 8)
                    .Select(i => slide.Shapes.AddShape(MsoAutoShapeType.msoShapeRectangle, left += 50, 0, 50, 50))
                    .ToArray();

                // Set common properties
                foreach (var shape in shapes)
                    {
                    shape.Tags.Add("配色", "快速配色");
                    shape.Line.Visible = MsoTriState.msoFalse;
                    }

                // Set the fill colors
                var colors = new[]
                {
                ToggleButton1.Background,
                ToggleButton2.Background,
                ToggleButton3.Background,
                ToggleButton4.Background,
                ToggleButton5.Background,
                ToggleButton6.Background,
                ToggleButton7.Background,
                ToggleButton8.Background,
                    };
                for (int i = 0 ; i < colors.Length ; i++)
                    {
                    var color = BrushToColor(colors[i]);
                    int A = color.A;
                    int R = color.R;
                    int G = color.G;
                    int B = color.B;
                    int rgb = RGB2Int(R, G, B);

                    shapes[i].Fill.ForeColor.RGB = rgb;
                    float tra = 1.0f - A / 255.0f;
                    //System.Windows.Forms.MessageBox.Show(tra.ToString());
                    shapes[i].Fill.Transparency = tra;
                    }
                }
            catch (Exception)
                {
                // Handle exceptions
                }
            }

        public int RGB2Int(int R, int G, int B)
            {
            int PPTRGB = R + G * 256 + B * 256 * 256;
            return PPTRGB;
            }

        private void Button_Click_2(object sender, RoutedEventArgs e)
            {
            var app = Globals.ThisAddIn.Application;
            var slide = app.ActiveWindow.View.Slide;
            var themeColorScheme = slide.ThemeColorScheme;
            var colors = new[]
               {
                ToggleButton1.Background,
                ToggleButton2.Background,
                ToggleButton3.Background,
                ToggleButton4.Background,
                ToggleButton5.Background,
                ToggleButton6.Background,
                ToggleButton7.Background,
                ToggleButton8.Background,
                    };

            for (int i = 0 ; i < colors.Length ; i++)
                {
                var color = BrushToColor(colors[i]);
                var rgb = color.R + color.G * 256 + color.B * 256 * 256;
                var themeIndex = (MsoThemeColorSchemeIndex)(i + 5);
                themeColorScheme.Colors(themeIndex).RGB = rgb;
                }

            this.Close();
            app.CommandBars.ExecuteMso("ThemeColorsCreateNew");
            }

        private void Shield_Click(object sender, RoutedEventArgs e)
            {
            Button_Click_1(sender, e);
            }

        private void Shield_Click_1(object sender, RoutedEventArgs e)
            {
            DelShpe("辅助配色");
            try
                {
                Slide slide = app.ActiveWindow.View.Slide;
                float left = -50;
                string PicPath = Properties.Settings.Default.PicPath;

                // Create the shapes
                var shapes = Enumerable.Range(0, 8)
                    .Select(i => slide.Shapes.AddShape(MsoAutoShapeType.msoShapeRectangle, left += 50, 60, 50, 50))
                    .ToArray();

                // Set common properties
                foreach (var shape in shapes)
                    {
                    shape.Tags.Add("配色", "辅助配色");
                    shape.Line.Visible = MsoTriState.msoFalse;
                    }

                // Set the fill colors
                var colors = new[]
                {
                ToggleButton9.Background,
                ToggleButton10.Background,
                ToggleButton11.Background,
                ToggleButton12.Background,
                ToggleButton13.Background,
                ToggleButton14.Background,
                ToggleButton15.Background,
                ToggleButton16.Background,
                    };
                for (int i = 0 ; i < colors.Length ; i++)
                    {
                    var color = BrushToColor(colors[i]);
                    int A = color.A;
                    int R = color.R;
                    int G = color.G;
                    int B = color.B;
                    int rgb = RGB2Int(R, G, B);

                    shapes[i].Fill.ForeColor.RGB = rgb;
                    float tra = 1.0f - A / 255.0f;
                    //System.Windows.Forms.MessageBox.Show(tra.ToString());
                    shapes[i].Fill.Transparency = tra;
                    }
                }
            catch (Exception)
                {
                // Handle exceptions
                }
            }

        /// <summary>
        /// 获取PPT主题色，MsoThemeColorSchemeIndex
        /// </summary>
        /// <param name="num"></param>
        /// <returns></returns>
        //获取主题色的函数
        public System.Drawing.Color GetColor(int num)
            {
            MsoThemeColorSchemeIndex num2 = new MsoThemeColorSchemeIndex();
            num2 = (MsoThemeColorSchemeIndex)num;
            Slide slide = app.ActiveWindow.View.Slide;
            int rgb = slide.ThemeColorScheme.Colors(num2).RGB;//输入代表主题的数字
            int r = rgb % 256;
            int g = rgb / 256 % 256;
            int b = rgb / 256 / 256 % 256;
            System.Drawing.Color Tcolor = System.Drawing.Color.FromArgb(255, (byte)r, (byte)g, (byte)b);
            return Tcolor;
            }

        public void GetThemeBtn()
            {
            app = Globals.ThisAddIn.Application;

            //获主题色
            ToggleButton17.Background = ConvertToWpfColor(GetColor(2));
            ToggleButton18.Background = ConvertToWpfColor(GetColor(1));
            ToggleButton19.Background = ConvertToWpfColor(GetColor(4));
            ToggleButton20.Background = ConvertToWpfColor(GetColor(3));
            ToggleButton21.Background = ConvertToWpfColor(GetColor(5));
            ToggleButton22.Background = ConvertToWpfColor(GetColor(6));
            ToggleButton23.Background = ConvertToWpfColor(GetColor(7));
            ToggleButton24.Background = ConvertToWpfColor(GetColor(8));
            ToggleButton25.Background = ConvertToWpfColor(GetColor(9));
            ToggleButton26.Background = ConvertToWpfColor(GetColor(10));
            ToggleButton27.Background = ConvertToWpfColor(GetColor(11));
            ToggleButton28.Background = ConvertToWpfColor(GetColor(12));
            //ToggleButton28.Background = ConvertToWpfColor(GetColor(13));
            //ToggleButton29.Background = ConvertToWpfColor(GetColor(14));
            //ToggleButton30.Background = ConvertToWpfColor(GetColor(15));
            }

        private void importTheme_Click(object sender, RoutedEventArgs e)
            {
            System.Windows.Forms.SaveFileDialog saveFileDialog = new System.Windows.Forms.SaveFileDialog();
            //设置文件类型
            saveFileDialog.Filter = "主题文件（*.thmx）|*.thmx";
            //设置默认文件类型显示顺序
            saveFileDialog.FilterIndex = 1;
            //保存对话框是否记忆上次打开的目录
            saveFileDialog.RestoreDirectory = true;
            saveFileDialog.FileName = System.IO.Path.GetFileNameWithoutExtension(app.ActivePresentation.Name) + "-主题方案";
            if (saveFileDialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                string path = saveFileDialog.FileName;
                app.Presentations[1].SaveCopyAs(path, PpSaveAsFileType.ppSaveAsOpenXMLTheme, MsoTriState.msoFalse);//导出配色
                }
            else
                {
                return;
                }
            }

        private void importColor_Click(object sender, RoutedEventArgs e)
            {
            Slide slide = app.ActiveWindow.View.Slide;
            SaveFileDialog saveFileDialog = new SaveFileDialog();
            //设置文件类型
            saveFileDialog.Filter = "配色方案（*.xml）|*.xml";
            //设置默认文件类型显示顺序
            saveFileDialog.FilterIndex = 1;
            //保存对话框是否记忆上次打开的目录
            saveFileDialog.RestoreDirectory = true;
            saveFileDialog.FileName = saveFileDialog.FileName = System.IO.Path.GetFileNameWithoutExtension(app.ActivePresentation.Name) + "-配色方案"; ;
            if (saveFileDialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                string cPath = saveFileDialog.FileName;
                slide.ThemeColorScheme.Save(cPath);
                }
            }

        private void GetThemeColor_Click(object sender, RoutedEventArgs e)
            {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            Presentation pre = app.ActivePresentation;
            //设置文件类型
            openFileDialog.Filter = "主题文件（*.thmx;*.xml）|*.thmx;*.xml";
            //设置默认文件类型显示顺序
            openFileDialog.Title = "请选��主题/配色文件";

            openFileDialog.FilterIndex = 1;
            //保存对话框是否记忆上次打开的目
            openFileDialog.RestoreDirectory = true;
            if (openFileDialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                string cPath = openFileDialog.FileName;
                if (System.IO.Path.GetExtension(cPath) == ".thmx")
                    {
                    pre.SlideMaster.ApplyTheme(cPath);//导入主题
                    }
                else
                    {
                    pre.SlideMaster.Theme.ThemeColorScheme.Load(cPath);//导入配色
                    }
                }
            else
                {
                }
            }

        private void Shield_Click_2(object sender, RoutedEventArgs e)
            {
            DelShpe("主题深色");
            try
                {
                Slide slide = app.ActiveWindow.View.Slide;
                float left = -50;
                string PicPath = Properties.Settings.Default.PicPath;

                // Create the shapes
                var shapes = Enumerable.Range(0, 4)
                    .Select(i => slide.Shapes.AddShape(MsoAutoShapeType.msoShapeRectangle, left += 50, 120, 50, 50))
                    .ToArray();

                // Set common properties
                foreach (var shape in shapes)
                    {
                    shape.Tags.Add("配色", "主题深色");
                    shape.Line.Visible = MsoTriState.msoFalse;
                    }

                // Set the fill colors
                var colors = new[]
                {
                ToggleButton17.Background,
                ToggleButton18.Background,
                ToggleButton19.Background,
                ToggleButton20.Background,
                    };
                for (int i = 0 ; i < colors.Length ; i++)
                    {
                    var color = BrushToColor(colors[i]);
                    int A = color.A;
                    int R = color.R;
                    int G = color.G;
                    int B = color.B;
                    int rgb = RGB2Int(R, G, B);

                    shapes[i].Fill.ForeColor.RGB = rgb;
                    float tra = 1.0f - A / 255.0f;
                    //System.Windows.Forms.MessageBox.Show(tra.ToString());
                    shapes[i].Fill.Transparency = tra;
                    }
                }
            catch (Exception)
                {
                // Handle exceptions
                }
            }

        private void Shield_Click_3(object sender, RoutedEventArgs e)
            {
            DelShpe("主题着色");
            try
                {
                Slide slide = app.ActiveWindow.View.Slide;
                float left = -50;
                string PicPath = Properties.Settings.Default.PicPath;

                // Create the shapes
                var shapes = Enumerable.Range(0, 8)
                    .Select(i => slide.Shapes.AddShape(MsoAutoShapeType.msoShapeRectangle, left += 50, 180, 50, 50))
                    .ToArray();

                // Set common properties
                foreach (var shape in shapes)
                    {
                    shape.Tags.Add("配色", "主题着色");
                    shape.Line.Visible = MsoTriState.msoFalse;
                    }

                // Set the fill colors
                var colors = new[]
                {
                 ToggleButton21.Background,
                 ToggleButton22.Background,
                 ToggleButton23.Background,
                 ToggleButton24.Background,
                 ToggleButton25.Background,
                 ToggleButton26.Background,
                 ToggleButton27.Background,
                 ToggleButton28.Background,
                    };
                for (int i = 0 ; i < colors.Length ; i++)
                    {
                    var color = BrushToColor(colors[i]);
                    int A = color.A;
                    int R = color.R;
                    int G = color.G;
                    int B = color.B;
                    int rgb = RGB2Int(R, G, B);

                    shapes[i].Fill.ForeColor.RGB = rgb;
                    float tra = 1.0f - A / 255.0f;
                    //System.Windows.Forms.MessageBox.Show(tra.ToString());
                    shapes[i].Fill.Transparency = tra;
                    }
                }
            catch (Exception)
                {
                // Handle exceptions
                }
            }

        // 保存配色方案
        private void SaveColorScheme_Click(object sender, RoutedEventArgs e)
            {
            try
                {
                var saveFileDialog = new SaveFileDialog
                    {
                    Filter = "配色方案文件(*.ppio)|*.ppio",
                    Title = "保存配色方案",
                    FileName = "ColorScheme_" + DateTime.Now.ToString("yyyyMMdd")
                    };

                if (saveFileDialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                    {
                    // 收集所有颜色按钮的颜色信息
                    var colorScheme = new Dictionary<string, List<string>>
                    {
                        {
                            "MainColors", new List<string>
                            {
                                GetColorHex(ToggleButton1),
                                GetColorHex(ToggleButton2),
                                GetColorHex(ToggleButton3),
                                GetColorHex(ToggleButton4),
                                GetColorHex(ToggleButton5),
                                GetColorHex(ToggleButton6),
                                GetColorHex(ToggleButton7),
                                GetColorHex(ToggleButton8)
                            }
                        },
                        {
                            "AuxColors", new List<string>
                            {
                                GetColorHex(ToggleButton9),
                                GetColorHex(ToggleButton10),
                                GetColorHex(ToggleButton11),
                                GetColorHex(ToggleButton12),
                                GetColorHex(ToggleButton13),
                                GetColorHex(ToggleButton14),
                                GetColorHex(ToggleButton15),
                                GetColorHex(ToggleButton16)
                            }
                        },
                        {
                            "ThemeColors", new List<string>
                            {
                                GetColorHex(ToggleButton17),
                                GetColorHex(ToggleButton18),
                                GetColorHex(ToggleButton19),
                                GetColorHex(ToggleButton20),
                                GetColorHex(ToggleButton21),
                                GetColorHex(ToggleButton22),
                                GetColorHex(ToggleButton23),
                                GetColorHex(ToggleButton24),
                                GetColorHex(ToggleButton25),
                                GetColorHex(ToggleButton26),
                                GetColorHex(ToggleButton27),
                                GetColorHex(ToggleButton28)
                            }
                        }
                    };

                    // 将配色方案序列化为JSON
                    string json = JsonConvert.SerializeObject(colorScheme, Formatting.Indented);

                    // 保存到文件
                    File.WriteAllText(saveFileDialog.FileName, json);
                    Growl.Success("配色方案保存成功！");
                    }
                }
            catch (Exception ex)
                {
                Growl.Error($"保存配色方案时出错: {ex.Message}");
                }
            }

        // 获取按钮颜色的十六进制表示
        private string GetColorHex(ToggleButton button)
            {
            if (button?.Background is SolidColorBrush brush)
                {
                var color = brush.Color;
                return $"#{color.R:X2}{color.G:X2}{color.B:X2}";
                }
            return "#808080"; // 默认灰色
            }

        // 导出调色板
        private void ExportPalette_Click(object sender, RoutedEventArgs e)
            {
            try
                {
                var saveFileDialog = new SaveFileDialog
                    {
                    Filter = "PNG图片(*.png)|*.png",
                    Title = "导出调色板",
                    FileName = "ColorPalette_" + DateTime.Now.ToString("yyyyMMdd")
                    };

                if (saveFileDialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                    {
                    ExportAsPng(saveFileDialog.FileName);
                    Growl.Success("调色板��出成功！");
                    }
                }
            catch (Exception ex)
                {
                Growl.Error($"导出调色板时出错: {ex.Message}");
                }
            }

        private void ExportAsPng(string fileName)
            {
            // 创建一个DrawingVisual���绘制调色板
            var drawingVisual = new DrawingVisual();

            // 设置画布大小
            int width = 1000; // 增加宽度以适应左右布局
            int height = 600;
            int margin = 40;
            int colorBlockSize = 50;
            int spacing = 10;
            int titleHeight = 50;
            int imageWidth = 400; // 左侧图片宽度

            using (DrawingContext dc = drawingVisual.RenderOpen())
                {
                // 绘制白色背景
                dc.DrawRectangle(Brushes.White, null, new Rect(0, 0, width, height));

                // 绘制标题
                var titleText = new FormattedText(
                    $"PresPio色卡 - {Path.GetFileNameWithoutExtension(fileName)}",
                    System.Globalization.CultureInfo.CurrentCulture,
                    FlowDirection.LeftToRight,
                    new Typeface("Microsoft YaHei"),
                    20,
                    Brushes.Black,
                    VisualTreeHelper.GetDpi(this).PixelsPerDip);

                dc.DrawText(titleText,
                    new Point((width - titleText.Width) / 2, 20));

                // 左侧图片区域
                if (ImageViewer.ImageSource is BitmapSource bitmapSource)
                    {
                    dc.DrawImage(bitmapSource,
                        new Rect(margin, titleHeight + margin, imageWidth, height - titleHeight - margin * 2));
                    }
                else
                    {
                    // 如果没有图片则绘制灰色背景
                    dc.DrawRectangle(Brushes.LightGray,
                        new Pen(Brushes.Gray, 1),
                        new Rect(margin, titleHeight + margin, imageWidth, height - titleHeight - margin * 2));
                    }

                // 右侧色块区域起始位置
                int colorStartX = margin + imageWidth + margin;
                int colorStartY = titleHeight + margin;
                int rowSpacing = 90;

                // 绘制主要配色
                var mainColors = new[] { ToggleButton1, ToggleButton2, ToggleButton3, ToggleButton4,
                               ToggleButton5, ToggleButton6, ToggleButton7, ToggleButton8 };
                DrawColorRow(dc, "主要配色", mainColors, colorStartX, colorStartY, colorBlockSize, spacing);

                // 绘制辅助配色
                var auxColors = new[] { ToggleButton9, ToggleButton10, ToggleButton11, ToggleButton12,
                              ToggleButton13, ToggleButton14, ToggleButton15, ToggleButton16 };
                DrawColorRow(dc, "辅助配色", auxColors, colorStartX, colorStartY + rowSpacing, colorBlockSize, spacing);

                // 绘制主题配色
                var themeColors = new[] { ToggleButton17, ToggleButton18, ToggleButton19, ToggleButton20,
                                ToggleButton21, ToggleButton22, ToggleButton23, ToggleButton24,
                                ToggleButton25, ToggleButton26, ToggleButton27, ToggleButton28 };

                // 主题配色两行布局
                var themeColorsRow1 = themeColors.Take(6).ToArray();
                DrawColorRow(dc, "主题配色", themeColorsRow1, colorStartX, colorStartY + rowSpacing * 2, colorBlockSize, spacing);

                var themeColorsRow2 = themeColors.Skip(6).Take(6).ToArray();
                DrawColorRow(dc, "", themeColorsRow2, colorStartX, colorStartY + rowSpacing * 2 + colorBlockSize + spacing * 2, colorBlockSize, spacing);

                // 绘制底部装饰线
                var decorativeLine = new Pen(Brushes.LightGray, 1);
                dc.DrawLine(decorativeLine,
                    new Point(margin, height - 30),
                    new Point(width - margin, height - 30));
                }

            // 创建RenderTargetBitmap
            var renderBitmap = new RenderTargetBitmap(width, height, 96, 96, PixelFormats.Pbgra32);
            renderBitmap.Render(drawingVisual);

            // 创建PNG编码器
            PngBitmapEncoder encoder = new PngBitmapEncoder();
            encoder.Frames.Add(BitmapFrame.Create(renderBitmap));

            // 保存文件
            using (var stream = File.Create(fileName))
                {
                encoder.Save(stream);
                }
            }

        private void DrawColorRow(DrawingContext dc, string title, ToggleButton[] buttons,
            int x, int y, int size, int spacing)
            {
            // 绘制标题
            var titleText = new FormattedText(
                title,
                System.Globalization.CultureInfo.CurrentCulture,
                FlowDirection.LeftToRight,
                new Typeface("Microsoft YaHei"),
                14,
                System.Windows.Media.Brushes.Black,
                VisualTreeHelper.GetDpi(this).PixelsPerDip);

            dc.DrawText(titleText, new Point(x, y));

            // 绘制颜色块
            int startX = x;
            int startY = y + 25;

            for (int i = 0 ; i < buttons.Length ; i++)
                {
                var brush = buttons[i].Background;
                var rect = new Rect(startX + (size + spacing) * i, startY, size, size);

                // 绘制颜色块
                dc.DrawRectangle(brush, new System.Windows.Media.Pen(System.Windows.Media.Brushes.LightGray, 1), rect);

                // 绘制颜色值
                if (brush is SolidColorBrush solidBrush)
                    {
                    var colorText = new FormattedText(
                        GetColorHex(buttons[i]),
                        System.Globalization.CultureInfo.CurrentCulture,
                        FlowDirection.LeftToRight,
                        new Typeface("Consolas"),
                        12,
                        Brushes.Black,
                        VisualTreeHelper.GetDpi(this).PixelsPerDip);

                    dc.DrawText(colorText,
                        new Point(rect.X + (size - colorText.Width) / 2,
                                 rect.Y + size + 5));
                    }
                }
            }

        // 提取主色调
        private void ExtractMainColor_Click(object sender, RoutedEventArgs e)
            {
            // TODO: 实现提取主色调的逻辑
            }

        // 分析配色方案
        private void AnalyzeColorScheme_Click(object sender, RoutedEventArgs e)
            {
            try
                {
                // 获取主要配色区域的颜色
                var mainButtons = new[]
                {
                    ToggleButton1, ToggleButton2, ToggleButton3, ToggleButton4,
                    ToggleButton5, ToggleButton6, ToggleButton7, ToggleButton8
                };

                var colors = new List<Color>();
                foreach (var button in mainButtons)
                    {
                    if (button.Background is SolidColorBrush brush)
                        {
                        colors.Add(brush.Color);
                        }
                    }

                if (colors.Count == 0)
                    {
                    Growl.Warning("没有找到需要分析的颜色");
                    return;
                    }

                // 分析颜色特征
                var analysis = new System.Text.StringBuilder();
                analysis.AppendLine("配色方案分析结果:");

                // 计算平均亮度
                double avgBrightness = colors.Average(c => (c.R + c.G + c.B) / 3.0);
                analysis.AppendLine($"整体亮度: {avgBrightness:F2}");

                // 分析色调分布
                foreach (var color in colors)
                    {
                    GetHSL(color, out double h, out double s, out double l);
                    string colorType = "";
                    if (h >= 0 && h < 30 || h >= 330) colorType = "红色系";
                    else if (h >= 30 && h < 90) colorType = "黄色系";
                    else if (h >= 90 && h < 150) colorType = "绿色系";
                    else if (h >= 150 && h < 210) colorType = "青色系";
                    else if (h >= 210 && h < 270) colorType = "蓝色系";
                    else if (h >= 270 && h < 330) colorType = "紫色系";

                    analysis.AppendLine($"颜色 #{color.R:X2}{color.G:X2}{color.B:X2} - {colorType}");
                    analysis.AppendLine($"  色相: {h:F0}°, 饱和度: {s:P0}, 亮度: {l:P0}");
                    }

                // 显示分析结果
                HandyControl.Controls.MessageBox.Show(analysis.ToString(), "配色方案分析", MessageBoxButton.OK, MessageBoxImage.Information);
                }
            catch (Exception ex)
                {
                Growl.Error($"分析配色方案时出错: {ex.Message}");
                }
            }

        // 复制色值
        private void CopyColorValue_Click(object sender, RoutedEventArgs e)
            {
            if (sender is Button button && button.Tag is ToggleButton colorButton)
                {
                if (colorButton.Background is SolidColorBrush brush)
                    {
                    var color = brush.Color;
                    string hexColor = $"#{color.R:X2}{color.G:X2}{color.B:X2}";
                    Clipboard.SetText(hexColor);
                    Growl.Success($"已复制颜色值: {hexColor}");
                    }
                else
                    {
                    Growl.Warning("无法获取颜色值");
                    }
                }
            }

        // 粘贴色值
        private void PasteColorValue_Click(object sender, RoutedEventArgs e)
            {
            if (sender is Button button && button.Tag is ToggleButton colorButton)
                {
                try
                    {
                    string clipboardText = Clipboard.GetText();
                    if (string.IsNullOrWhiteSpace(clipboardText))
                        {
                        Growl.Warning("剪贴板中没有可用的颜色值");
                        return;
                        }

                    // 尝试转换颜色值
                    var color = (Color)System.Windows.Media.ColorConverter.ConvertFromString(clipboardText);
                    colorButton.Background = new SolidColorBrush(color);
                    Growl.Success($"已粘贴颜色值: {clipboardText}");
                    }
                catch (Exception)
                    {
                    Growl.Error("无效的颜色值格式");
                    }
                }
            }

        // 清空选择
        private void ClearSelection_Click(object sender, RoutedEventArgs e)
            {
            // TODO: 实现清空选择的逻辑
            }

        // 显示颜色选择器
        private void ShowColorPicker(ToggleButton colorButton)
            {
            if (colorButton.IsChecked == true)
                {
                ColorPicker.SelectedBrush = (SolidColorBrush)colorButton.Background;
                ColorPickerPopup.IsOpen = true;
                }
            }

        // 应用预设方案
        private void ApplyPreset_Click(object sender, RoutedEventArgs e)
            {
            try
                {
                // 获取当前选中的预设方案
                var comboBox = PresetComboBox;
                if (comboBox == null || comboBox.SelectedItem == null)
                    {
                    Growl.Warning("请先选择一个预设方案");
                    return;
                    }

                // 获取选中的预设方案名称
                var selectedItem = comboBox.SelectedItem as ComboBoxItem;
                string selectedScheme = selectedItem?.Content?.ToString();
                if (string.IsNullOrEmpty(selectedScheme))
                    {
                    Growl.Warning("无效的预设方案");
                    return;
                    }

                // 根据选择的预设方案获取对应的颜色数组
                string[] colorScheme;
                switch (selectedScheme)
                    {
                    case "经典配色":
                        colorScheme = ClassicScheme;
                        break;

                    case "现代简约":
                        colorScheme = ModernScheme;
                        break;

                    case "自然清新":
                        colorScheme = NatureScheme;
                        break;

                    case "科技感":
                        colorScheme = TechScheme;
                        break;

                    case "商务专业":
                        colorScheme = BusinessScheme;
                        break;

                    default:
                        Growl.Warning("未知的预设方案");
                        return;
                    }

                // 应用颜色方案
                ApplyColorScheme(colorScheme);

                // 显示成功提示
                Growl.Success($"已成功应用{selectedScheme}预设方案");
                }
            catch (Exception ex)
                {
                Growl.Error($"应用预设方案时出错: {ex.Message}");
                }
            }

        // 预设配色方案
        private readonly string[] ClassicScheme = new[]
        {
            "#2C3E50", "#E74C3C", "#ECF0F1", "#3498DB",
            "#2980B9", "#27AE60", "#16A085", "#F1C40F"
        };

        private readonly string[] ModernScheme = new[]
        {
            "#000000", "#FFFFFF", "#FF4081", "#3F51B5",
            "#2196F3", "#009688", "#4CAF50", "#FFC107"
        };

        private readonly string[] NatureScheme = new[]
        {
            "#8BC34A", "#CDDC39", "#4CAF50", "#009688",
            "#00BCD4", "#03A9F4", "#5C6BC0", "#7E57C2"
        };

        private readonly string[] TechScheme = new[]
        {
            "#212121", "#00BCD4", "#E0E0E0", "#2196F3",
            "#607D8B", "#00E5FF", "#448AFF", "#40C4FF"
        };

        private readonly string[] BusinessScheme = new[]
        {
            "#1F2937", "#374151", "#4B5563", "#6B7280",
            "#9CA3AF", "#D1D5DB", "#E5E7EB", "#F3F4F6"
        };

        // 应用颜色方案到按钮
        private void ApplyColorScheme(string[] colors)
            {
            var buttons = new[]
            {
                ToggleButton1, ToggleButton2, ToggleButton3, ToggleButton4,
                ToggleButton5, ToggleButton6, ToggleButton7, ToggleButton8
            };

            for (int i = 0 ; i < Math.Min(buttons.Length, colors.Length) ; i++)
                {
                try
                    {
                    var color = (System.Windows.Media.Color)System.Windows.Media.ColorConverter.ConvertFromString(colors[i]);
                    buttons[i].Background = new SolidColorBrush(color);
                    }
                catch (Exception)
                    {
                    HandyControl.Controls.Growl.Error($"颜色转换错误: {colors[i]}");
                    }
                }
            }

        // 导入配色方案
        private void ImportColorScheme_Click(object sender, RoutedEventArgs e)
            {
            try
                {
                var openFileDialog = new OpenFileDialog
                    {
                    Filter = "配色方案文件(*.ppio)|*.ppio|所有文件(*.*)|*.*",
                    Title = "导入配色方案",
                    CheckFileExists = true
                    };

                if (openFileDialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                    {
                    if (!File.Exists(openFileDialog.FileName))
                        {
                        Growl.Warning("所选文件不存在");
                        return;
                        }

                    string json = File.ReadAllText(openFileDialog.FileName);
                    if (string.IsNullOrWhiteSpace(json))
                        {
                        Growl.Warning("配色方案文件为空");
                        return;
                        }

                    var colorScheme = JsonConvert.DeserializeObject<Dictionary<string, List<string>>>(json);
                    if (colorScheme == null || !colorScheme.Any())
                        {
                        Growl.Warning("无效的配色方案格式");
                        return;
                        }

                    bool hasChanges = false;

                    // 应用主要配色
                    if (colorScheme.TryGetValue("MainColors", out var mainColors) && mainColors?.Count > 0)
                        {
                        var mainButtons = new[]
                        {
                            ToggleButton1, ToggleButton2, ToggleButton3, ToggleButton4,
                            ToggleButton5, ToggleButton6, ToggleButton7, ToggleButton8
                        };
                        ApplyColors(mainButtons, mainColors);
                        hasChanges = true;
                        }

                    // 应用辅助配色
                    if (colorScheme.TryGetValue("AuxColors", out var auxColors) && auxColors?.Count > 0)
                        {
                        var auxButtons = new[]
                        {
                            ToggleButton9, ToggleButton10, ToggleButton11, ToggleButton12,
                            ToggleButton13, ToggleButton14, ToggleButton15, ToggleButton16
                        };
                        ApplyColors(auxButtons, auxColors);
                        hasChanges = true;
                        }

                    // 应用主题配色
                    if (colorScheme.TryGetValue("ThemeColors", out var themeColors) && themeColors?.Count > 0)
                        {
                        var themeButtons = new[]
                        {
                            ToggleButton17, ToggleButton18, ToggleButton19, ToggleButton20,
                            ToggleButton21, ToggleButton22, ToggleButton23, ToggleButton24,
                            ToggleButton25, ToggleButton26, ToggleButton27, ToggleButton28
                        };
                        ApplyColors(themeButtons, themeColors);
                        hasChanges = true;
                        }

                    if (hasChanges)
                        {
                        Growl.Success($"成功导入配色方案：{Path.GetFileNameWithoutExtension(openFileDialog.FileName)}");
                        }
                    else
                        {
                        Growl.Warning("配色方案中没有有效的颜色数据");
                        }
                    }
                }
            catch (JsonException)
                {
                Growl.Error("配色方案文件格式错误");
                }
            catch (IOException ex)
                {
                Growl.Error($"读取文件时出错: {ex.Message}");
                }
            catch (Exception ex)
                {
                Growl.Error($"导入��色方案时出错: {ex.Message}");
                }
            }

        // 优化后的ApplyColors方法
        private void ApplyColors(ToggleButton[] buttons, List<string> colors)
            {
            if (buttons == null || colors == null)
                return;

            int count = Math.Min(buttons.Length, colors.Count);
            for (int i = 0 ; i < count ; i++)
                {
                try
                    {
                    if (buttons[i] != null && !string.IsNullOrWhiteSpace(colors[i]))
                        {
                        var color = (System.Windows.Media.Color)System.Windows.Media.ColorConverter.ConvertFromString(colors[i]);
                        buttons[i].Background = new SolidColorBrush(color);
                        }
                    }
                catch (FormatException)
                    {
                    Growl.Warning($"无效的颜色值: {colors[i]}");
                    }
                catch (Exception ex)
                    {
                    Growl.Warning($"应用颜色时出错: {ex.Message}");
                    }
                }
            }

        // 智能配色
        private void SmartColorMatch_Click(object sender, RoutedEventArgs e)
            {
            try
                {
                // 获取主要配色的第一个颜色作为基准色
                var baseBrush = ToggleButton1.Background as SolidColorBrush;
                if (baseBrush == null)
                    {
                    Growl.Warning("请先生成或选择一个基准色");
                    return;
                    }

                // 生成配色方案
                var baseColor = baseBrush.Color;
                var colorScheme = GenerateColorScheme(baseColor);

                // 应用配色方案到辅助配色区域
                var buttons = new[]
                {
                    ToggleButton9, ToggleButton10, ToggleButton11, ToggleButton12,
                    ToggleButton13, ToggleButton14, ToggleButton15, ToggleButton16
                };

                for (int i = 0 ; i < Math.Min(buttons.Length, colorScheme.Count) ; i++)
                    {
                    buttons[i].Background = new SolidColorBrush(colorScheme[i]);
                    }

                Growl.Success("智能配色方案已生成到辅助配色区域");
                }
            catch (Exception ex)
                {
                Growl.Error($"生成配色方案时出错: {ex.Message}");
                }
            }

        // 修改生成配色方案的逻辑
        private List<Color> GenerateColorScheme(Color baseColor)
            {
            var colors = new List<Color>();

            // 转换为HSL
            ColorToHSL(baseColor, out double h, out double s, out double l);

            // 生成配色方案
            colors.Add(baseColor);  // 基准色
            colors.Add(HSLToColor(AdjustHue(h + 180), s, l));  // 互补色
            colors.Add(HSLToColor(AdjustHue(h + 120), s, l));  // 三角色1
            colors.Add(HSLToColor(AdjustHue(h - 120), s, l));  // 三角色2

            // 类似色（稍微调整饱和度和亮度）
            colors.Add(HSLToColor(AdjustHue(h + 30), Math.Min(1, s * 1.1), Math.Min(1, l * 1.1)));
            colors.Add(HSLToColor(AdjustHue(h - 30), Math.Min(1, s * 1.1), Math.Min(1, l * 0.9)));

            // 明暗变化
            colors.Add(HSLToColor(h, s, AdjustLightness(l + 0.2)));  // 更亮
            colors.Add(HSLToColor(h, s, AdjustLightness(l - 0.2)));  // 更暗

            return colors;
            }

        // 颜色转换辅助方法
        private void ColorToHSL(Color color, out double hue, out double saturation, out double lightness)
            {
            double r = color.R / 255.0;
            double g = color.G / 255.0;
            double b = color.B / 255.0;

            double max = Math.Max(Math.Max(r, g), b);
            double min = Math.Min(Math.Min(r, g), b);
            double delta = max - min;

            lightness = (max + min) / 2.0;

            if (delta == 0)
                {
                hue = 0;
                saturation = 0;
                }
            else
                {
                saturation = lightness < 0.5 ? delta / (max + min) : delta / (2.0 - max - min);

                if (r == max)
                    hue = (g - b) / delta + (g < b ? 6 : 0);
                else if (g == max)
                    hue = (b - r) / delta + 2;
                else
                    hue = (r - g) / delta + 4;

                hue *= 60;
                }
            }

        private Color HSLToColor(double hue, double saturation, double lightness)
            {
            double r, g, b;

            if (saturation == 0)
                {
                r = g = b = lightness;
                }
            else
                {
                double q = lightness < 0.5 ?
                    lightness * (1 + saturation) :
                    lightness + saturation - lightness * saturation;
                double p = 2 * lightness - q;

                r = HueToRGB(p, q, hue + 120);
                g = HueToRGB(p, q, hue);
                b = HueToRGB(p, q, hue - 120);
                }

            return Color.FromRgb(
                (byte)(r * 255),
                (byte)(g * 255),
                (byte)(b * 255));
            }

        private double HueToRGB(double p, double q, double t)
            {
            if (t < 0) t += 360;
            if (t > 360) t -= 360;

            if (t < 60) return p + (q - p) * t / 60;
            if (t < 180) return q;
            if (t < 240) return p + (q - p) * (240 - t) / 60;
            return p;
            }

        private double AdjustHue(double hue)
            {
            while (hue >= 360) hue -= 360;
            while (hue < 0) hue += 360;
            return hue;
            }

        private double AdjustLightness(double lightness)
            {
            return Math.Max(0, Math.Min(1, lightness));
            }

        // 色彩平衡
        private void ColorBalance_Click(object sender, RoutedEventArgs e)
            {
            try
                {
                // 获取主要配色的所有颜色
                var buttons = new[]
                {
                    ToggleButton1, ToggleButton2, ToggleButton3, ToggleButton4,
                    ToggleButton5, ToggleButton6, ToggleButton7, ToggleButton8
                };

                // 收集所有颜色
                var colors = new List<Color>();
                foreach (var button in buttons)
                    {
                    if (button.Background is SolidColorBrush brush)
                        {
                        colors.Add(brush.Color);
                        }
                    }

                if (colors.Count == 0)
                    {
                    Growl.Warning("没有找到需要平衡的颜色");
                    return;
                    }

                // 计算RGB通道的平均值
                double avgR = colors.Average(c => c.R);
                double avgG = colors.Average(c => c.G);
                double avgB = colors.Average(c => c.B);

                // 计算整体亮度
                double avgBrightness = colors.Average(c => (c.R + c.G + c.B) / 3.0);

                // 平衡后的颜色
                var balancedColors = new List<Color>();
                foreach (var color in colors)
                    {
                    // 计算当前颜色的亮度
                    double brightness = (color.R + color.G + color.B) / 3.0;

                    // 计算亮度调整系数
                    double brightnessFactor = avgBrightness / Math.Max(1, brightness);
                    brightnessFactor = Math.Max(0.7, Math.Min(1.3, brightnessFactor)); // 限制调整范围

                    // 计算RGB平衡系数
                    double rFactor = color.R > 0 ? avgR / color.R : 1.0;
                    double gFactor = color.G > 0 ? avgG / color.G : 1.0;
                    double bFactor = color.B > 0 ? avgB / color.B : 1.0;

                    // 限制调整范围
                    rFactor = Math.Max(0.8, Math.Min(1.2, rFactor));
                    gFactor = Math.Max(0.8, Math.Min(1.2, gFactor));
                    bFactor = Math.Max(0.8, Math.Min(1.2, bFactor));

                    // 应用平衡
                    byte newR = (byte)Math.Min(255, Math.Max(0, color.R * rFactor * brightnessFactor));
                    byte newG = (byte)Math.Min(255, Math.Max(0, color.G * gFactor * brightnessFactor));
                    byte newB = (byte)Math.Min(255, Math.Max(0, color.B * bFactor * brightnessFactor));

                    balancedColors.Add(Color.FromRgb(newR, newG, newB));
                    }

                // 应用平衡后的颜色到辅助配色区域
                var auxButtons = new[]
                {
                    ToggleButton9, ToggleButton10, ToggleButton11, ToggleButton12,
                    ToggleButton13, ToggleButton14, ToggleButton15, ToggleButton16
                };

                for (int i = 0 ; i < Math.Min(auxButtons.Length, balancedColors.Count) ; i++)
                    {
                    auxButtons[i].Background = new SolidColorBrush(balancedColors[i]);
                    }

                Growl.Success("色彩平衡已完成，结果显示在辅助配色区域");
                }
            catch (Exception ex)
                {
                Growl.Error($"色彩平衡时出错: {ex.Message}");
                }
            }

        // 辅助方法：计算颜色的HSL值
        private void GetHSL(Color color, out double h, out double s, out double l)
            {
            double r = color.R / 255.0;
            double g = color.G / 255.0;
            double b = color.B / 255.0;

            double max = Math.Max(Math.Max(r, g), b);
            double min = Math.Min(Math.Min(r, g), b);

            h = s = l = (max + min) / 2.0;

            if (max == min)
                {
                h = s = 0;
                }
            else
                {
                double d = max - min;
                s = l > 0.5 ? d / (2.0 - max - min) : d / (max + min);

                if (max == r)
                    h = (g - b) / d + (g < b ? 6 : 0);
                else if (max == g)
                    h = (b - r) / d + 2;
                else
                    h = (r - g) / d + 4;

                h *= 60;
                }
            }

        // 撤销操作
        private void UndoOperation_Click(object sender, RoutedEventArgs e)
            {
            // TODO: 实现撤销操作的逻辑
            }

        // 重做操作
        private void RedoOperation_Click(object sender, RoutedEventArgs e)
            {
            // TODO: 实现重做操作的逻辑
            }
        }
    }