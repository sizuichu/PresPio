using HandyControl.Controls;
using System;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Windows;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.IO;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using Application = Microsoft.Office.Interop.PowerPoint.Application;

namespace PresPio
{
    public partial class Wpf_ColorAdjust : INotifyPropertyChanged
    {
        public event PropertyChangedEventHandler PropertyChanged;

        private ObservableCollection<ColorCard> _colorCards;
        public ObservableCollection<ColorCard> ColorCards
        {
            get => _colorCards;
            set
            {
                _colorCards = value;
                OnPropertyChanged(nameof(ColorCards));
            }
        }

        private string _currentScheme = "";
        private readonly Application _powerPoint;

        public Wpf_ColorAdjust(Application powerPoint)
        {
            _powerPoint = powerPoint;
            InitializeComponent();
            ColorCards = new ObservableCollection<ColorCard>();
            ColorCardList.ItemsSource = ColorCards;

            MainColorPicker.SelectedBrush = new SolidColorBrush(Colors.DodgerBlue);
            
            MainColorPicker.SelectedColorChanged += (s, e) =>
            {
                if (ColorCards.Count > 0)
                {
                    RegenerateCurrentScheme();
                }
            };
        }

        private void RegenerateCurrentScheme()
        {
            switch (_currentScheme)
            {
                case "Monochrome":
                    OnMonochrome(null, null);
                    break;
                case "Complementary":
                    OnComplementary(null, null);
                    break;
                case "Triadic":
                    OnTriadic(null, null);
                    break;
                case "Analogous":
                    OnAnalogous(null, null);
                    break;
                case "SplitComplementary":
                    OnSplitComplementary(null, null);
                    break;
            }
        }

        #region 色卡生成方法
        private void OnMonochrome(object sender, RoutedEventArgs e)
        {
            _currentScheme = "Monochrome";
            var baseColor = ((SolidColorBrush)MainColorPicker.SelectedBrush).Color;
            ColorCards.Clear();

            var hsv = ColorToHSV(baseColor);
            
            // 生成15个不同亮度的色卡
            for (int i = 0; i < 15; i++)
            {
                double value = Math.Min(Math.Max(hsv.v + (i - 7) * 0.067, 0), 1);
                var color = HSVToColor(hsv.h, hsv.s, value);
                AddColorCard($"色调 {i + 1}", color);
            }
        }

        private void OnComplementary(object sender, RoutedEventArgs e)
        {
            _currentScheme = "Complementary";
            var baseColor = ((SolidColorBrush)MainColorPicker.SelectedBrush).Color;
            ColorCards.Clear();

            var hsv = ColorToHSV(baseColor);
            
            // 生成主色系列（8个）
            for (int i = 0; i < 8; i++)
            {
                double value = Math.Min(Math.Max(hsv.v + (i - 3.5) * 0.125, 0), 1);
                var color = HSVToColor(hsv.h, hsv.s, value);
                AddColorCard($"主色 {i + 1}", color);
            }

            // 生成互补色系列（7个）
            double complementaryHue = (hsv.h + 180) % 360;
            for (int i = 0; i < 7; i++)
            {
                double value = Math.Min(Math.Max(hsv.v + (i - 3) * 0.125, 0), 1);
                var color = HSVToColor(complementaryHue, hsv.s, value);
                AddColorCard($"互补 {i + 1}", color);
            }
        }

        private void OnTriadic(object sender, RoutedEventArgs e)
        {
            _currentScheme = "Triadic";
            var baseColor = ((SolidColorBrush)MainColorPicker.SelectedBrush).Color;
            ColorCards.Clear();

            var hsv = ColorToHSV(baseColor);
            
            // 生成三组颜色，每组5个不同亮度
            double[] hues = { hsv.h, (hsv.h + 120) % 360, (hsv.h + 240) % 360 };
            string[] names = { "主色", "三色1", "三色2" };

            for (int i = 0; i < 3; i++)
            {
                for (int j = 0; j < 5; j++)
                {
                    double value = Math.Min(Math.Max(hsv.v + (j - 2) * 0.2, 0), 1);
                    var color = HSVToColor(hues[i], hsv.s, value);
                    AddColorCard($"{names[i]} {j + 1}", color);
                }
            }
        }

        private void OnAnalogous(object sender, RoutedEventArgs e)
        {
            _currentScheme = "Analogous";
            var baseColor = ((SolidColorBrush)MainColorPicker.SelectedBrush).Color;
            ColorCards.Clear();

            var hsv = ColorToHSV(baseColor);
            
            // 生成15个类比色
            for (int i = -7; i <= 7; i++)
            {
                double hue = (hsv.h + i * 10 + 360) % 360;
                double saturation = Math.Min(Math.Max(hsv.s + i * 0.03, 0), 1);
                var color = HSVToColor(hue, saturation, hsv.v);
                AddColorCard(i == 0 ? "主色" : $"类比色 {Math.Abs(i)}", color);
            }
        }

        private void OnSplitComplementary(object sender, RoutedEventArgs e)
        {
            _currentScheme = "SplitComplementary";
            var baseColor = ((SolidColorBrush)MainColorPicker.SelectedBrush).Color;
            ColorCards.Clear();

            var hsv = ColorToHSV(baseColor);

            // 主色系列（5个）
            for (int i = 0; i < 5; i++)
            {
                double value = Math.Min(Math.Max(hsv.v + (i - 2) * 0.2, 0), 1);
                AddColorCard($"主色 {i + 1}", HSVToColor(hsv.h, hsv.s, value));
            }

            // 分裂色1系列（5个）
            double splitHue1 = (hsv.h + 150) % 360;
            for (int i = 0; i < 5; i++)
            {
                double value = Math.Min(Math.Max(hsv.v + (i - 2) * 0.2, 0), 1);
                AddColorCard($"分裂1-{i + 1}", HSVToColor(splitHue1, hsv.s, value));
            }

            // 分裂色2系列（5个）
            double splitHue2 = (hsv.h + 210) % 360;
            for (int i = 0; i < 5; i++)
            {
                double value = Math.Min(Math.Max(hsv.v + (i - 2) * 0.2, 0), 1);
                AddColorCard($"分裂2-{i + 1}", HSVToColor(splitHue2, hsv.s, value));
            }
        }
        #endregion

        #region 辅助方法
        private void AddColorCard(string name, Color color)
        {
            ColorCards.Add(new ColorCard
            {
                ColorName = name,
                ColorBrush = new SolidColorBrush(color),
                ColorCode = $"#{color.R:X2}{color.G:X2}{color.B:X2}",
                RgbValue = $"RGB({color.R}, {color.G}, {color.B})",
                TextBrush = new SolidColorBrush(GetContrastColor(color))
            });
        }

        private (double h, double s, double v) ColorToHSV(Color color)
        {
            double r = color.R / 255.0;
            double g = color.G / 255.0;
            double b = color.B / 255.0;

            double max = Math.Max(r, Math.Max(g, b));
            double min = Math.Min(r, Math.Min(g, b));
            double delta = max - min;

            double hue = 0;
            if (delta != 0)
            {
                if (max == r)
                    hue = 60 * ((g - b) / delta % 6);
                else if (max == g)
                    hue = 60 * ((b - r) / delta + 2);
                else
                    hue = 60 * ((r - g) / delta + 4);
            }
            if (hue < 0) hue += 360;

            double saturation = max == 0 ? 0 : delta / max;
            double value = max;

            return (hue, saturation, value);
        }

        private Color HSVToColor(double h, double s, double v)
        {
            double c = v * s;
            double x = c * (1 - Math.Abs((h / 60) % 2 - 1));
            double m = v - c;

            double r = 0, g = 0, b = 0;
            if (h < 60) { r = c; g = x; }
            else if (h < 120) { r = x; g = c; }
            else if (h < 180) { g = c; b = x; }
            else if (h < 240) { g = x; b = c; }
            else if (h < 300) { r = x; b = c; }
            else { r = c; b = x; }

            return Color.FromRgb(
                (byte)((r + m) * 255),
                (byte)((g + m) * 255),
                (byte)((b + m) * 255));
        }

        private Color GetContrastColor(Color backgroundColor)
        {
            double luminance = (0.299 * backgroundColor.R + 
                              0.587 * backgroundColor.G + 
                              0.114 * backgroundColor.B) / 255;
            return luminance > 0.5 ? Colors.Black : Colors.White;
        }
        #endregion

        #region 导出功能
        private void OnExport(object sender, RoutedEventArgs e)
        {
            if (ColorCards.Count == 0)
            {
                Growl.Warning("请先生成色卡");
                return;
            }

            try
            {
                var dialog = new Microsoft.Win32.SaveFileDialog
                {
                    Filter = "PNG图片|*.png|所有文件|*.*",
                    DefaultExt = ".png",
                    FileName = "色卡导出_" + DateTime.Now.ToString("yyyyMMdd_HHmmss")
                };

                if (dialog.ShowDialog() == true)
                {
                    var panel = new WrapPanel();
                    foreach (var card in ColorCards)
                    {
                        var border = CreateColorCardElement(card);
                        panel.Children.Add(border);
                    }

                    SaveColorCardsImage(panel, dialog.FileName);
                    Growl.Success("色卡导出成功");
                }
            }
            catch (Exception ex)
            {
                Growl.Error($"导出失败: {ex.Message}");
            }
        }

        private Border CreateColorCardElement(ColorCard card)
        {
            var border = new Border
            {
                Width = 150,
                Height = 90,
                Margin = new Thickness(5),
                Background = card.ColorBrush,
                BorderBrush = new SolidColorBrush(Colors.LightGray),
                BorderThickness = new Thickness(1),
                CornerRadius = new CornerRadius(4)
            };

            var grid = new Grid();
            grid.RowDefinitions.Add(new RowDefinition { Height = new GridLength(1, GridUnitType.Star) });
            grid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });

            var nameText = new TextBlock
            {
                Text = card.ColorName,
                Foreground = card.TextBrush,
                HorizontalAlignment = HorizontalAlignment.Center,
                VerticalAlignment = VerticalAlignment.Center,
                FontSize = 14
            };
            Grid.SetRow(nameText, 0);

            var infoBorder = new Border
            {
                Background = new SolidColorBrush(Color.FromArgb(128, 0, 0, 0)),
                Padding = new Thickness(5)
            };
            Grid.SetRow(infoBorder, 1);

            var infoPanel = new StackPanel();
            infoPanel.Children.Add(new TextBlock
            {
                Text = card.ColorCode,
                Foreground = new SolidColorBrush(Colors.White),
                HorizontalAlignment = HorizontalAlignment.Center
            });
            infoPanel.Children.Add(new TextBlock
            {
                Text = card.RgbValue,
                Foreground = new SolidColorBrush(Colors.White),
                HorizontalAlignment = HorizontalAlignment.Center,
                FontSize = 11
            });

            infoBorder.Child = infoPanel;
            grid.Children.Add(nameText);
            grid.Children.Add(infoBorder);
            border.Child = grid;

            return border;
        }

        private void SaveColorCardsImage(WrapPanel panel, string fileName)
        {
            panel.Measure(new Size(double.PositiveInfinity, double.PositiveInfinity));
            panel.Arrange(new Rect(new Point(0, 0), panel.DesiredSize));

            var renderBitmap = new RenderTargetBitmap(
                (int)panel.ActualWidth,
                (int)panel.ActualHeight,
                96, 96, PixelFormats.Pbgra32);

            renderBitmap.Render(panel);

            var encoder = new PngBitmapEncoder();
            encoder.Frames.Add(BitmapFrame.Create(renderBitmap));

            using (var stream = File.Create(fileName))
            {
                encoder.Save(stream);
            }
        }
        #endregion

        private void OnColorCardClick(object sender, System.Windows.Input.MouseButtonEventArgs e)
        {
            if (sender is Border border && border.Background is SolidColorBrush brush)
            {
                var colorCode = $"#{brush.Color.R:X2}{brush.Color.G:X2}{brush.Color.B:X2}";
                Clipboard.SetText(colorCode);
                
                Growl.Info($"已复制颜色代码: {colorCode}");
            }
        }

        protected void OnPropertyChanged(string propertyName)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }

        private void OnApplyColors(object sender, RoutedEventArgs e)
        {
            if (ColorCards.Count == 0)
            {
                Growl.Warning("请先生成色卡");
                return;
            }

            try
            {
                var slide = _powerPoint.ActiveWindow.View.Slide;
                float slideWidth = slide.Design.SlideMaster.Width;
                float slideHeight = slide.Design.SlideMaster.Height;

                // 计算形状大小和位置
                float shapeSize = Math.Min(slideWidth, slideHeight) / 8; // 形状大小
                float startX = 50; // 起始X坐标
                float startY = 50; // 起始Y坐标
                float spacing = shapeSize * 1.2f; // 形状间距
                int maxShapesPerRow = (int)((slideWidth - startX * 2) / spacing); // 每行最大形状数

                // 清除现有形状
                var existingShapes = slide.Shapes.Cast<Microsoft.Office.Interop.PowerPoint.Shape>()
                    .Where(s => s.Name.StartsWith("ColorShape_"))
                    .ToList();
                foreach (var shape in existingShapes)
                {
                    shape.Delete();
                }

                // 创建新形状
                for (int i = 0; i < ColorCards.Count; i++)
                {
                    int row = i / maxShapesPerRow;
                    int col = i % maxShapesPerRow;

                    float left = startX + col * spacing;
                    float top = startY + row * spacing;

                    var shape = slide.Shapes.AddShape(
                        MsoAutoShapeType.msoShapeRectangle,
                        left, top, shapeSize, shapeSize);

                    // 设置形状属性
                    var color = ((SolidColorBrush)ColorCards[i].ColorBrush).Color;
                    shape.Fill.ForeColor.RGB = (color.R) | (color.G << 8) | (color.B << 16);
                    shape.Line.Visible = MsoTriState.msoTrue;
                    shape.Line.ForeColor.RGB = 0x808080; // 灰色边框
                    shape.Name = $"ColorShape_{i}"; // 设置形状名称以便识别

                    // 添加颜色代码标签
                    var textShape = slide.Shapes.AddTextbox(
                        MsoTextOrientation.msoTextOrientationHorizontal,
                        left, top + shapeSize, shapeSize, 20);
                    textShape.TextFrame.TextRange.Text = ColorCards[i].ColorCode;
                    textShape.TextFrame.TextRange.Font.Size = 10;
                    textShape.TextFrame.TextRange.Font.Color.RGB = 0x000000; // 黑色文字
                    textShape.Name = $"ColorLabel_{i}";
                }

                Growl.Success("配色已应用到当前幻灯片");
            }
            catch (Exception ex)
            {
                Growl.Error($"应用配色失败: {ex.Message}");
            }
        }
    }

    public class ColorCard
    {
        public string ColorName { get; set; }
        public SolidColorBrush ColorBrush { get; set; }
        public string ColorCode { get; set; }
        public string RgbValue { get; set; }
        public SolidColorBrush TextBrush { get; set; }
    }
} 