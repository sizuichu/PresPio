using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using System.Windows.Media.Effects;
using System.Windows.Shapes;
using HandyControl.Controls;
using HandyControl.Data;
using Microsoft.Office.Interop.PowerPoint;

using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace PresPio
    {
    public partial class Wpf__Pagesize
        {
        private readonly Dictionary<PpSlideSizeType, string> pageDictionary = new Dictionary<PpSlideSizeType, string>
        {
            { PpSlideSizeType.ppSlideSizeOnScreen, "全屏显示" },
            { PpSlideSizeType.ppSlideSizeLetterPaper, "信纸" },
            { PpSlideSizeType.ppSlideSizeA4Paper, "A4纸张" },
            { PpSlideSizeType.ppSlideSizeA3Paper, "A3纸张" },
            { PpSlideSizeType.ppSlideSizeB4ISOPaper, "B4 ISO纸张" },
            { PpSlideSizeType.ppSlideSizeB5ISOPaper, "B5 ISO纸张" },
            { PpSlideSizeType.ppSlideSizeB5JISPaper, "B5 JIS纸张" },
            { PpSlideSizeType.ppSlideSizeBanner, "横幅" },
            { PpSlideSizeType.ppSlideSizeCustom, "自定义" },
            { PpSlideSizeType.ppSlideSizeHagakiCard, "Hagaki卡片" },
            { PpSlideSizeType.ppSlideSizeLedgerPaper, "分类帐纸张" },
            { PpSlideSizeType.ppSlideSize35MM, "35MM" },
            { PpSlideSizeType.ppSlideSizeOverhead, "顶置" }
        };

        public PowerPoint.Application app;

        // 添加新的字段来存储当前PPT尺寸
        private float currentHeight;

        private float currentWidth;

        public Wpf__Pagesize()
            {
            InitializeComponent();
            Loaded += PageWindow_Loaded;

            // PowerPoint中1英寸=72磅
            Num1.Value = 540;  // 默认高度 7.5英寸
            Num2.Value = 720;  // 默认宽度 10英寸
            Num1.Minimum = 72;  // 最小1英寸
            Num2.Minimum = 72;
            Num1.Maximum = 5184;  // 最大72英寸
            Num2.Maximum = 5184;

            // 设置小数位数为2位
            Num1.DecimalPlaces = 2;
            Num2.DecimalPlaces = 2;
            }

        private void PageWindow_Loaded(object sender, RoutedEventArgs e)
            {
            app = Globals.ThisAddIn.Application;
            if (app?.ActivePresentation != null)
                {
                currentHeight = (float)app.ActivePresentation.PageSetup.SlideHeight;
                currentWidth = (float)app.ActivePresentation.PageSetup.SlideWidth;

                // 设置当前尺寸为默认值
                Num1.Value = Math.Round(currentHeight, 2);
                Num2.Value = Math.Round(currentWidth, 2);

                UpdateCurrentSizeDisplay();
                }
            ListPages();
            UpdatePreview();
            }

        // 添加显示当前PPT尺寸的方法
        private void UpdateCurrentSizeDisplay()
            {
            if (app?.ActivePresentation != null)
                {
                var presentation = app.ActivePresentation;
                var orientation = HtogBtn.IsChecked == true ? "横向" : "纵向";
                var heightInInches = Math.Round(currentHeight / 72.0, 2);
                var widthInInches = Math.Round(currentWidth / 72.0, 2);

                CurrentSizeText.Text = $"当前页面尺寸: {Math.Round(currentWidth, 2)} × {Math.Round(currentHeight, 2)} 磅" +
                                      $" ({widthInInches} × {heightInInches} 英寸)  |  " +
                                      $"方向: {orientation}  |  起始页码: {presentation.PageSetup.FirstSlideNumber}";
                }
            }

        private void PageComBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
            {
            if (PageComBox.SelectedItem is MyPageSize selectedSize)
                {
                // 根据当前方向设置尺寸
                if (HtogBtn.IsChecked == true)
                    {
                    Num1.Value = Math.Round(selectedSize.Width, 2);
                    Num2.Value = Math.Round(selectedSize.Height, 2);
                    }
                else
                    {
                    Num1.Value = Math.Round(selectedSize.Height, 2);
                    Num2.Value = Math.Round(selectedSize.Width, 2);
                    }

                UpdatePreview();
                }
            }

        private void HtogBtn_Click(object sender, RoutedEventArgs e)
            {
            if (HtogBtn.IsChecked == true)
                {
                VtogBtn.IsChecked = false;
                SwapDimensions();
                }
            }

        private void VtogBtn_Click(object sender, RoutedEventArgs e)
            {
            if (VtogBtn.IsChecked == true)
                {
                HtogBtn.IsChecked = false;
                SwapDimensions();
                }
            }

        private void SwapDimensions()
            {
            var temp = Num1.Value;
            Num1.Value = Num2.Value;
            Num2.Value = temp;
            UpdatePreview();
            }

        // 修改预览方法，增加更多细节
        private void UpdatePreview()
            {
            if (PreviewCanvas == null) return;

            PreviewCanvas.Children.Clear();

            var height = (float)Num1.Value;
            var width = (float)Num2.Value;

            // 计算缩放比例以适应预览区域
            var scale = Math.Min(
                (PreviewBorder.ActualWidth - 40) / width,
                (PreviewBorder.ActualHeight - 40) / height
            ) * 0.9;

            // 创建页���矩形
            var rect = new Rectangle
                {
                Width = Math.Max(1, width * scale),  // 确保最小宽度为1
                Height = Math.Max(1, height * scale), // 确保最小高度为1
                Stroke = new SolidColorBrush(Colors.Gray),
                StrokeThickness = 1,
                Fill = new SolidColorBrush(Colors.White)
                };

            // 添加阴影效果
            var effect = new DropShadowEffect
                {
                Color = Colors.Gray,
                Direction = 315,
                ShadowDepth = 3,
                Opacity = 0.3
                };
            rect.Effect = effect;

            // 居中放置
            Canvas.SetLeft(rect, (PreviewCanvas.ActualWidth - rect.Width) / 2);
            Canvas.SetTop(rect, (PreviewCanvas.ActualHeight - rect.Height) / 2);

            PreviewCanvas.Children.Add(rect);

            // 添加尺寸标注
            AddSizeLabel(rect, width, height, scale);

            // 添加方向指示
            AddOrientationIndicator(rect);
            }

        // 添加尺寸标注的方法
        private void AddSizeLabel(Rectangle rect, float width, float height, double scale)
            {
            var label = new TextBlock
                {
                Text = $"{Math.Round(width, 2)} × {Math.Round(height, 2)}",
                Foreground = new SolidColorBrush(Colors.Gray),
                FontSize = 12
                };

            Canvas.SetLeft(label, Canvas.GetLeft(rect));
            Canvas.SetTop(label, Canvas.GetTop(rect) - 20);

            PreviewCanvas.Children.Add(label);
            }

        // 添加方向指示器
        private void AddOrientationIndicator(Rectangle rect)
            {
            var orientation = HtogBtn.IsChecked == true ? "横向" : "纵向";
            var label = new TextBlock
                {
                Text = orientation,
                Foreground = new SolidColorBrush(Colors.Gray),
                FontSize = 12
                };

            Canvas.SetLeft(label, Canvas.GetLeft(rect));
            Canvas.SetTop(label, Canvas.GetTop(rect) - 35);

            PreviewCanvas.Children.Add(label);
            }

        public class MyPageSize
            {
            public string Name { get; set; }
            public float Height { get; set; }
            public float Width { get; set; }

            // 添加尺寸显示属性
            public string SizeDisplay
                {
                get
                    {
                    var heightInInches = Math.Round(Height / 72.0, 2);
                    var widthInInches = Math.Round(Width / 72.0, 2);
                    return $"{Math.Round(Width, 0)} × {Math.Round(Height, 0)} 磅 ({widthInInches}\" × {heightInInches}\")";
                    }
                }

            public static MyPageSize CreatePageSize(string name, float height, float width)
                {
                return new MyPageSize { Name = name, Height = height, Width = width };
                }
            }

        private MyPageSize[] myPageSizes;

        public void ListPages()
            {
            try
                {
                var pageSizeList = new List<MyPageSize>();

                // 添加常用尺寸（单位：磅）
                pageSizeList.Add(MyPageSize.CreatePageSize("标准(4:3)", 540, 720));     // 7.5" × 10"
                pageSizeList.Add(MyPageSize.CreatePageSize("宽屏(16:9)", 540, 960));    // 7.5" × 13.33"
                pageSizeList.Add(MyPageSize.CreatePageSize("A4纸张", 841.89f, 595.28f)); // A4尺寸
                pageSizeList.Add(MyPageSize.CreatePageSize("信纸", 792, 612));          // US Letter
                pageSizeList.Add(MyPageSize.CreatePageSize("A3纸张", 1190.55f, 841.89f)); // A3尺寸
                pageSizeList.Add(MyPageSize.CreatePageSize("B4纸张", 1031.81f, 728.5f));  // B4尺寸
                pageSizeList.Add(MyPageSize.CreatePageSize("B5纸张", 728.5f, 515.91f));   // B5尺寸

                // 添加当前尺寸（如果与预设不同）
                var currentSize = MyPageSize.CreatePageSize("当前尺寸", currentHeight, currentWidth);
                if (!pageSizeList.Any(p => Math.Abs(p.Height - currentHeight) < 1 && Math.Abs(p.Width - currentWidth) < 1))
                    {
                    pageSizeList.Insert(0, currentSize);
                    }

                myPageSizes = pageSizeList.ToArray();
                PageComBox.ItemsSource = myPageSizes;
                PageComBox.SelectedIndex = 0;
                }
            catch (Exception ex)
                {
                Growl.Warning($"加载预设尺寸时出错：{ex.Message}");
                }
            }

        // 修改事件处理方法的签名
        private void DimensionValueChanged(object sender, FunctionEventArgs<double> e)
            {
            if (e.Info > 0)
                {
                UpdatePreview();
                }
            }

        private void StartNumberChanged(object sender, FunctionEventArgs<double> e)
            {
            if (app?.ActivePresentation != null && e.Info > 0)
                {
                try
                    {
                    app.ActivePresentation.PageSetup.FirstSlideNumber = (int)e.Info;
                    UpdateCurrentSizeDisplay();
                    }
                catch (Exception ex)
                    {
                    Growl.Warning($"设置起始页码时出错：{ex.Message}");
                    }
                }
            }

        // 添加新的按钮事件处理方法
        private void ApplyButton_Click(object sender, RoutedEventArgs e)
            {
            var pre = app?.ActivePresentation;
            if (pre != null)
                {
                try
                    {
                    // 设置页面尺寸
                    var pageHeight = (float)Math.Round(Num1.Value, 2);
                    var pageWidth = (float)Math.Round(Num2.Value, 2);

                    if (pageHeight <= 0 || pageWidth <= 0)
                        {
                        Growl.Warning("页面尺寸必须大于0");
                        return;
                        }

                    // 先设置为自定义尺寸类型
                    pre.PageSetup.SlideSize = PpSlideSizeType.ppSlideSizeCustom;

                    // 根据方向设置尺寸
                    if (HtogBtn.IsChecked == true)
                        {
                        pre.PageSetup.SlideHeight = pageWidth;
                        pre.PageSetup.SlideWidth = pageHeight;
                        }
                    else
                        {
                        pre.PageSetup.SlideHeight = pageHeight;
                        pre.PageSetup.SlideWidth = pageWidth;
                        }

                    // 更新当前尺寸
                    currentHeight = (float)pre.PageSetup.SlideHeight;
                    currentWidth = (float)pre.PageSetup.SlideWidth;
                    UpdateCurrentSizeDisplay();

                    Growl.SuccessGlobal("页面设置已应用");
                    Close();
                    }
                catch (Exception ex)
                    {
                    Growl.Warning($"设置页面时出错：{ex.Message}");
                    }
                }
            }

        private void CancelButton_Click(object sender, RoutedEventArgs e)
            {
            Close();
            }
        }
    }