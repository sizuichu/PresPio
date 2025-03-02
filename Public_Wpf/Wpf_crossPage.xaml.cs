using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media.Imaging;
using Microsoft.Office.Interop.PowerPoint;
using Path = System.IO.Path;

using PowerPoint = Microsoft.Office.Interop.PowerPoint;

using Shape = Microsoft.Office.Interop.PowerPoint.Shape;

namespace PresPio
    {
    public partial class Wpf_crossPage
        {
        public PowerPoint.Application app { get; set; }
        private const float POSITION_TOLERANCE = 1.0f;
        private List<int> previewPages = new List<int>();

        public Wpf_crossPage()
            {
            app = Globals.ThisAddIn.Application;
            InitializeComponent();
            InitializeControls();
            SetupEventHandlers();
            }

        private void InitializeControls()
            {
            Presentation pre = app.ActivePresentation;
            int slideCount = pre.Slides.Count;
            int currentSlide = app.ActiveWindow.Selection.SlideRange.SlideIndex;

            // 设置页码范围
            NumericUpDown1.Minimum = 1;
            NumericUpDown1.Maximum = slideCount;
            NumericUpDown1.Value = currentSlide;

            NumericUpDown2.Maximum = slideCount;
            NumericUpDown2.Minimum = 1;
            NumericUpDown2.Value = slideCount;

            // 初始化复制选项
            IntervalUpDown.Visibility = Visibility.Collapsed;
            CustomRangeBox.Visibility = Visibility.Collapsed;
            }

        private void SetupEventHandlers()
            {
            app.WindowSelectionChange += Application_ShapeSelectionChange;
            app.AfterShapeSizeChange += Application_AfterShapeSizeChange;
            Loaded += CrossWindow_Loaded;
            CopyModeCombo.SelectionChanged += CopyModeCombo_SelectionChanged;
            }

        public void Application_ShapeSelectionChange(Selection Sel)
            {
            LoadImg();
            }

        private void Application_AfterShapeSizeChange(Shape Shape)
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
                if (sel == null || sel.Type != PpSelectionType.ppSelectionShapes || sel.ShapeRange.Count != 1)
                    {
                    imageBox.Source = new BitmapImage(new Uri("/PresPio;component/Images/Icons/Lucency.png", UriKind.Relative));
                    NoSelectionText.Visibility = Visibility.Visible;
                    return;
                    }

                NoSelectionText.Visibility = Visibility.Collapsed;
                string tempFolderName = Guid.NewGuid().ToString("N");
                string tempDir = Path.Combine(Path.GetTempPath(), "MyAppTemp", tempFolderName);
                Directory.CreateDirectory(tempDir);

                try
                    {
                    string tempFile = Path.Combine(tempDir, $"temp_{Guid.NewGuid():N}.png");
                    sel.ShapeRange[1].Export(tempFile, PpShapeFormat.ppShapeFormatPNG);

                    if (File.Exists(tempFile))
                        {
                        BitmapImage bitmapImage = new BitmapImage();
                        bitmapImage.BeginInit();
                        bitmapImage.UriSource = new Uri(tempFile, UriKind.Absolute);
                        bitmapImage.CacheOption = BitmapCacheOption.OnLoad;
                        bitmapImage.DecodePixelWidth = 160; // 限制解码大小以提高性能
                        bitmapImage.EndInit();
                        bitmapImage.Freeze(); // 提高性能

                        imageBox.Source = bitmapImage;
                        }
                    }
                finally
                    {
                    if (Directory.Exists(tempDir))
                        {
                        Directory.Delete(tempDir, true);
                        }
                    }
                }
            catch (Exception ex)
                {
                System.Diagnostics.Debug.WriteLine($"LoadImg error: {ex.Message}");
                imageBox.Source = new BitmapImage(new Uri("/PresPio;component/Images/Icons/Lucency.png", UriKind.Relative));
                NoSelectionText.Visibility = Visibility.Visible;
                }
            }

        private void DelBtn_Click(object sender, RoutedEventArgs e)
            {
            try
                {
                Selection sel = app.ActiveWindow.Selection;
                if (sel.Type != PpSelectionType.ppSelectionShapes || sel.ShapeRange.Count != 1)
                    {
                    MessageBox.Show("请选择单个形状后再试！", "提示");
                    return;
                    }

                Shape selectedShape = sel.ShapeRange[1];
                float targetTop = selectedShape.Top;
                float targetLeft = selectedShape.Left;
                float targetWidth = selectedShape.Width;
                float targetHeight = selectedShape.Height;

                Presentation pre = app.ActivePresentation;
                int startPage = (int)NumericUpDown1.Value;
                int endPage = (int)NumericUpDown2.Value;
                List<int> targetPages = GetTargetPages(startPage, endPage);
                int deletedCount = 0;

                foreach (int pageNum in targetPages)
                    {
                    if (pageNum > 0 && pageNum <= pre.Slides.Count)
                        {
                        Slide slide = pre.Slides[pageNum];
                        List<Shape> shapesToDelete = new List<Shape>();
                        foreach (Shape shape in slide.Shapes)
                            {
                            if (IsShapeMatch(shape, targetTop, targetLeft, targetWidth, targetHeight))
                                {
                                shapesToDelete.Add(shape);
                                }
                            }

                        foreach (Shape shape in shapesToDelete)
                            {
                            shape.Delete();
                            deletedCount++;
                            }
                        }
                    }

                MessageBox.Show($"删除完成！共删除 {deletedCount} 个对象。", "提示");
                }
            catch (Exception ex)
                {
                MessageBox.Show($"操作失败：{ex.Message}", "错误");
                }
            }

        private bool IsShapeMatch(Shape shape, float targetTop, float targetLeft, float targetWidth, float targetHeight)
            {
            return Math.Abs(shape.Top - targetTop) < POSITION_TOLERANCE &&
                   Math.Abs(shape.Left - targetLeft) < POSITION_TOLERANCE &&
                   Math.Abs(shape.Width - targetWidth) < POSITION_TOLERANCE &&
                   Math.Abs(shape.Height - targetHeight) < POSITION_TOLERANCE;
            }

        private void CopyModeCombo_SelectionChanged(object sender, SelectionChangedEventArgs e)
            {
            string selectedMode = ((ComboBoxItem)CopyModeCombo.SelectedItem).Content.ToString();
            IntervalUpDown.Visibility = selectedMode == "指定间隔复制" ? Visibility.Visible : Visibility.Collapsed;
            CustomRangeBox.Visibility = selectedMode == "自定义范围复制" ? Visibility.Visible : Visibility.Collapsed;
            }

        private void QuickSelectBtn_Click(object sender, RoutedEventArgs e)
            {
            try
                {
                Selection sel = app.ActiveWindow.Selection;
                if (sel.Type == PpSelectionType.ppSelectionSlides)
                    {
                    SlideRange slideRange = sel.SlideRange;
                    NumericUpDown1.Value = slideRange.SlideIndex;
                    NumericUpDown2.Value = slideRange.SlideIndex + slideRange.Count - 1;
                    }
                }
            catch (Exception ex)
                {
                MessageBox.Show($"快速选择失败：{ex.Message}", "错误");
                }
            }

        private void PreviewBtn_Click(object sender, RoutedEventArgs e)
            {
            try
                {
                int startPage = (int)NumericUpDown1.Value;
                int endPage = (int)NumericUpDown2.Value;
                previewPages = GetTargetPages(startPage, endPage);

                if (previewPages.Count == 0)
                    {
                    MessageBox.Show("当前设置下没有符合条件的页面。", "提示");
                    return;
                    }

                string pageList = string.Join(", ", previewPages);
                MessageBox.Show($"将会在以下页面执行操作：\n{pageList}", "预览");
                }
            catch (Exception ex)
                {
                MessageBox.Show($"预览失败：{ex.Message}", "错误");
                }
            }

        private List<int> GetTargetPages(int startPage, int endPage)
            {
            List<int> pages = new List<int>();
            string copyMode = ((ComboBoxItem)CopyModeCombo.SelectedItem).Content.ToString();

            if (copyMode == "自定义范围复制")
                {
                pages = ParseCustomRange(CustomRangeBox.Text, startPage, endPage);
                }
            else
                {
                for (int i = startPage ; i <= endPage ; i++)
                    {
                    switch (copyMode)
                        {
                        case "单页复制":
                            pages.Add(i);
                            break;

                        case "奇数页复制":
                            if (i % 2 != 0) pages.Add(i);
                            break;

                        case "偶数页复制":
                            if (i % 2 == 0) pages.Add(i);
                            break;

                        case "指定间隔复制":
                            if ((i - startPage) % (int)IntervalUpDown.Value == 0)
                                pages.Add(i);
                            break;
                        }
                    }
                }
            return pages;
            }

        private List<int> ParseCustomRange(string range, int defaultStart, int defaultEnd)
            {
            var result = new HashSet<int>();
            if (string.IsNullOrWhiteSpace(range))
                {
                return new List<int> { defaultStart };
                }

            foreach (var part in range.Split(','))
                {
                if (part.Contains("-"))
                    {
                    var bounds = part.Split('-');
                    if (bounds.Length == 2 && int.TryParse(bounds[0], out int start) && int.TryParse(bounds[1], out int end))
                        {
                        for (int i = start ; i <= end ; i++)
                            {
                            result.Add(i);
                            }
                        }
                    }
                else if (int.TryParse(part, out int page))
                    {
                    result.Add(page);
                    }
                }

            return result.Where(p => p >= defaultStart && p <= defaultEnd).OrderBy(p => p).ToList();
            }

        private void CopBtn_Click(object sender, RoutedEventArgs e)
            {
            if (app == null)
                {
                MessageBox.Show("PowerPoint应用程序未初始化！", "错误");
                return;
                }

            try
                {
                Selection sel = app.ActiveWindow.Selection;
                if (sel == null)
                    {
                    MessageBox.Show("无法获取当前选择！", "提示");
                    return;
                    }

                if (sel.Type != PpSelectionType.ppSelectionShapes)
                    {
                    MessageBox.Show("请先选择要复制的形状！", "提示");
                    return;
                    }

                if (sel.ShapeRange.Count != 1)
                    {
                    MessageBox.Show("请只选择一个形状！", "提示");
                    return;
                    }

                Presentation pre = app.ActivePresentation;
                if (pre == null)
                    {
                    MessageBox.Show("请先打开一个PowerPoint文档！", "提示");
                    return;
                    }

                int startPage = (int)NumericUpDown1.Value;
                int endPage = (int)NumericUpDown2.Value;

                if (startPage < 1 || endPage > pre.Slides.Count)
                    {
                    MessageBox.Show($"页码范围无效！请输入1到{pre.Slides.Count}之间的数字。", "提示");
                    return;
                    }

                if (startPage > endPage)
                    {
                    MessageBox.Show("起始页不能大于结束页！", "提示");
                    return;
                    }

                sel.ShapeRange.Copy();
                List<int> targetPages = GetTargetPages(startPage, endPage);
                int copyCount = 0;

                foreach (int pageNum in targetPages)
                    {
                    if (pageNum > 0 && pageNum <= pre.Slides.Count)
                        {
                        try
                            {
                            pre.Slides[pageNum].Shapes.Paste();
                            copyCount++;
                            }
                        catch (Exception ex)
                            {
                            System.Diagnostics.Debug.WriteLine($"复制到页面 {pageNum} 失败：{ex.Message}");
                            continue;
                            }
                        }
                    }

                if (copyCount > 0)
                    {
                    MessageBox.Show($"复制完成！共复制 {copyCount} 个对象。", "提示");
                    }
                else
                    {
                    MessageBox.Show("复制失败，未能复制任何对象。", "提示");
                    }
                }
            catch (Exception ex)
                {
                MessageBox.Show($"操作失败：{ex.Message}", "错误");
                System.Diagnostics.Debug.WriteLine($"CopBtn_Click error: {ex}");
                }
            }
        }
    }