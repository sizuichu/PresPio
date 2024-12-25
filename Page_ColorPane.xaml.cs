using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Collections.Specialized;
using System.ComponentModel;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using HandyControl.Controls;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using System.Windows.Media.Animation;
using HandyControl.Tools;
using System.Linq;

namespace PresPio
{
    public class ColorClass
    {
        public string ColorName { get; set; }
        public string HexValue { get; set; }
        public float Hue { get; set; }
        public float Saturation { get; set; }
        public float Lightness { get; set; }
        public SolidColorBrush BackgroundColor
        {
            get
            {
                var color = HSLToRGB(Hue, Saturation, Lightness);
                return new SolidColorBrush(color);
            }
        }

        public ColorClass(string colorName, string hexValue, float hue, float saturation, float lightness)
        {
            ColorName = colorName;
            HexValue = hexValue;
            Hue = hue;
            Saturation = saturation;
            Lightness = lightness;
        }

        private Color HSLToRGB(float h, float s, float l)
        {
            float c = (1 - Math.Abs(2 * l - 1)) * s;
            float x = c * (1 - Math.Abs((h / 60) % 2 - 1));
            float m = l - c / 2;

            float r, g, b;

            if (h >= 0 && h < 60) { r = c; g = x; b = 0; }
            else if (h >= 60 && h < 120) { r = x; g = c; b = 0; }
            else if (h >= 120 && h < 180) { r = 0; g = c; b = x; }
            else if (h >= 180 && h < 240) { r = 0; g = x; b = c; }
            else if (h >= 240 && h < 300) { r = x; g = 0; b = c; }
            else { r = c; g = 0; b = x; }

            r += m;
            g += m;
            b += m;

            return Color.FromArgb(255, (byte)(r * 255), (byte)(g * 255), (byte)(b * 255));
        }
    }

    public partial class Page_ColorPane : UserControl
    {
        public PowerPoint.Application app;
        private ColorViewModel _viewModel;
        private const int MAX_RECENT_COLORS = 12;
        public ObservableCollection<ColorGroup> ColorSchemeGroups { get; set; }

        public Page_ColorPane()
        {
            InitializeComponent();
            _viewModel = new ColorViewModel();
            ColorSchemeGroups = new ObservableCollection<ColorGroup>();
            
            // 初始化颜色列表
            _viewModel.Colors = InitializeColors();
            _viewModel.RecentColors = LoadRecentColors();
            _viewModel.FavoriteColors = new ObservableCollection<SolidColorBrush>();
            _viewModel.CurrentOpacity = 100;
            
            this.DataContext = _viewModel;
            InitializeColorSchemes();
        }

        private ObservableCollection<SolidColorBrush> LoadRecentColors()
        {
            var recentColors = new ObservableCollection<SolidColorBrush>();
            try
            {
                if (Properties.Settings.Default.RecentColors == null)
                {
                    Properties.Settings.Default.RecentColors = new StringCollection();
                    Properties.Settings.Default.Save();
                }

                foreach (string colorStr in Properties.Settings.Default.RecentColors)
                {
                    try
                    {
                        var color = (Color)ColorConverter.ConvertFromString(colorStr);
                        recentColors.Add(new SolidColorBrush(color));
                    }
                    catch { }
                }
            }
            catch (Exception ex)
            {
                Growl.ErrorGlobal($"加载最近使用的颜色时出错: {ex.Message}");
            }
            return recentColors;
        }

        private void SaveRecentColors()
        {
            try
            {
                var colorList = _viewModel.RecentColors.Select(brush => brush.Color.ToString()).ToList();
                Properties.Settings.Default.RecentColors = new StringCollection();
                Properties.Settings.Default.RecentColors.AddRange(colorList.ToArray());
                Properties.Settings.Default.Save();
            }
            catch (Exception ex)
            {
                Growl.ErrorGlobal($"保存最近使用的颜色时出错: {ex.Message}");
            }
        }

        private void AddToRecentColors(Color color)
        {
            var brush = new SolidColorBrush(color);
            if (_viewModel.RecentColors.Contains(brush))
            {
                _viewModel.RecentColors.Remove(brush);
            }
            _viewModel.RecentColors.Insert(0, brush);

            while (_viewModel.RecentColors.Count > MAX_RECENT_COLORS)
            {
                _viewModel.RecentColors.RemoveAt(_viewModel.RecentColors.Count - 1);
            }

            SaveRecentColors();
        }

        private List<ColorClass> InitializeColors()
        {
            var colors = new List<ColorClass>();

            // 生成红色系列（R变化，GB为0）
            for (int r = 0; r <= 255; r += 5)
            {
                colors.Add(new ColorClass($"R{r}", "", 0, 1.0f, r / 255.0f));
            }

            // 生成绿色系列（G变化，RB为0）
            for (int g = 0; g <= 255; g += 5)
            {
                colors.Add(new ColorClass($"G{g}", "", 120, 1.0f, g / 255.0f));
            }

            // 生成蓝色系列（B变化，RG为0）
            for (int b = 0; b <= 255; b += 5)
            {
                colors.Add(new ColorClass($"B{b}", "", 240, 1.0f, b / 255.0f));
            }

            // 生成黄色系列（RG变化，B为0）
            for (int y = 0; y <= 255; y += 5)
            {
                colors.Add(new ColorClass($"Y{y}", "", 60, 1.0f, y / 255.0f));
            }

            // 生成青色系列（GB变化，R为0）
            for (int c = 0; c <= 255; c += 5)
            {
                colors.Add(new ColorClass($"C{c}", "", 180, 1.0f, c / 255.0f));
            }

            // 生成紫色系列（RB变化，G为0）
            for (int m = 0; m <= 255; m += 5)
            {
                colors.Add(new ColorClass($"M{m}", "", 300, 1.0f, m / 255.0f));
            }

            // 生成灰度系列（RGB同值）
            for (int gray = 0; gray <= 255; gray += 5)
            {
                float value = gray / 255.0f;
                colors.Add(new ColorClass($"Gray{gray}", "", 0, 0, value));
            }

            // 更新每个颜色的十六进制值
            foreach (var color in colors)
            {
                var brush = color.BackgroundColor;
                color.HexValue = brush.Color.ToString();
            }

            return colors;
        }

        private void OnColorClick(object sender, MouseButtonEventArgs e)
        {
            if (sender is Border border && border.Background is SolidColorBrush brush)
            {
                _viewModel.CurrentColor = brush;
                _viewModel.CurrentColorHex = brush.Color.ToString();
                AddToRecentColors(brush.Color);
                GenerateSubColors(brush.Color);
                ApplyColor(brush.Color);
            }
        }

        private void GenerateSubColors(Color baseColor)
        {
            _viewModel.SubColors.Clear();
            float h, s, v;
            ColorToHSV(baseColor, out h, out s, out v);

            // 生成12个子颜色
            // 第一行：6个颜色，保持色相不变，调整饱和度
            for (int i = 0; i < 6; i++)
            {
                float newS = Math.Min(1.0f, (i + 1) * 0.2f);
                Color newColor = HSVToColor(h, newS, v);
                _viewModel.SubColors.Add(new ColorClass(
                    $"饱和度 {i + 1}",
                    $"#{newColor.R:X2}{newColor.G:X2}{newColor.B:X2}",
                    h,
                    newS,
                    v
                ));
            }

            // 第二行：6个颜色，保持饱和度不变，调整色相
            for (int i = 0; i < 6; i++)
            {
                float newH = (h + i * 30) % 360; // 每次增加30度
                Color newColor = HSVToColor(newH, s, v);
                _viewModel.SubColors.Add(new ColorClass(
                    $"色相 {i + 1}",
                    $"#{newColor.R:X2}{newColor.G:X2}{newColor.B:X2}",
                    newH,
                    s,
                    v
                ));
            }
        }

        private void ColorToHSV(Color color, out float h, out float s, out float v)
        {
            float r = color.R / 255f;
            float g = color.G / 255f;
            float b = color.B / 255f;

            float max = Math.Max(r, Math.Max(g, b));
            float min = Math.Min(r, Math.Min(g, b));
            float delta = max - min;

            // Hue
            if (delta == 0)
            {
                h = 0;
            }
            else if (max == r)
            {
                h = 60 * ((g - b) / delta % 6);
            }
            else if (max == g)
            {
                h = 60 * ((b - r) / delta + 2);
            }
            else
            {
                h = 60 * ((r - g) / delta + 4);
            }

            if (h < 0)
                h += 360;

            // Saturation
            s = max == 0 ? 0 : delta / max;

            // Value
            v = max;
        }

        private Color HSVToColor(float h, float s, float v)
        {
            float c = v * s;
            float x = c * (1 - Math.Abs((h / 60) % 2 - 1));
            float m = v - c;

            float r, g, b;

            if (h >= 0 && h < 60)
            {
                r = c; g = x; b = 0;
            }
            else if (h >= 60 && h < 120)
            {
                r = x; g = c; b = 0;
            }
            else if (h >= 120 && h < 180)
            {
                r = 0; g = c; b = x;
            }
            else if (h >= 180 && h < 240)
            {
                r = 0; g = x; b = c;
            }
            else if (h >= 240 && h < 300)
            {
                r = x; g = 0; b = c;
            }
            else
            {
                r = c; g = 0; b = x;
            }

            byte red = (byte)((r + m) * 255);
            byte green = (byte)((g + m) * 255);
            byte blue = (byte)((b + m) * 255);

            return Color.FromRgb(red, green, blue);
        }

        private void ApplyColor(Color color)
        {
            try
            {
                // 创建新的颜色对象，保留RGB值但使用当前的透明度
                color = Color.FromArgb(
                    (byte)(_viewModel.CurrentOpacity * 255 / 100),
                    color.R,
                    color.G,
                    color.B
                );

                int colorRgb = color.R | (color.G << 8) | (color.B << 16);
                var selection = Globals.ThisAddIn.Application.ActiveWindow.Selection;
                bool isCtrlPressed = (System.Windows.Forms.Control.ModifierKeys & System.Windows.Forms.Keys.Control) == System.Windows.Forms.Keys.Control;

                if (selection.Type == PowerPoint.PpSelectionType.ppSelectionShapes)
                {
                    foreach (PowerPoint.Shape shape in selection.ShapeRange)
                    {
                        if (isCtrlPressed)
                        {
                            // 设置边框颜色
                            shape.Line.ForeColor.RGB = colorRgb;
                            shape.Line.Visible = Microsoft.Office.Core.MsoTriState.msoTrue;
                            shape.Line.Transparency = (float)(1 - (_viewModel.CurrentOpacity / 100.0));
                        }
                        else
                        {
                            // 设置填充颜色
                            shape.Fill.ForeColor.RGB = colorRgb;
                            shape.Fill.Visible = Microsoft.Office.Core.MsoTriState.msoTrue;
                            shape.Fill.Transparency = (float)(1 - (_viewModel.CurrentOpacity / 100.0));
                        }
                    }
                }
                else if (selection.Type == PowerPoint.PpSelectionType.ppSelectionText)
                {
                    // 设置文字颜色
                    selection.TextRange.Font.Color.RGB = colorRgb;
                }
                else
                {
                    Growl.WarningGlobal("请先在PPT中选择要填充颜色的形状或文字");
                    return;
                }

                // 更新当前颜色和最近使用的颜色
                AddToRecentColors(color);
                _viewModel.CurrentColor = new SolidColorBrush(color);
                _viewModel.CurrentColorHex = $"#{color.R:X2}{color.G:X2}{color.B:X2}";
            }
            catch (Exception ex)
            {
                Growl.ErrorGlobal($"应用颜色时出错: {ex.Message}");
            }
        }

        private void OnColorRightClick(object sender, MouseButtonEventArgs e)
        {
            var border = sender as Border;
            if (border?.Background is SolidColorBrush brush)
            {
                Clipboard.SetText(brush.Color.ToString());
            }
        }

        private void Border_MouseRightButtonDown(object sender, MouseButtonEventArgs e)
        {
            var border = sender as Border;
            if (border?.Background is SolidColorBrush brush)
            {
                Clipboard.SetText(brush.Color.ToString());
            }
        }

        private void OnSubColorClick(object sender, MouseButtonEventArgs e)
        {
            if (sender is Border border && border.Background is SolidColorBrush brush)
            {
                _viewModel.CurrentColor = brush;
                _viewModel.CurrentColorHex = brush.Color.ToString();
                AddToRecentColors(brush.Color);
                ApplyColor(brush.Color);
            }
        }

        private void OnSchemeColorClick(object sender, MouseButtonEventArgs e)
        {
            if (sender is Border border && border.Background is SolidColorBrush brush)
            {
                _viewModel.CurrentColor = brush;
                _viewModel.CurrentColorHex = brush.Color.ToString();
                AddToRecentColors(brush.Color);
                GenerateSubColors(brush.Color);
                ApplyColor(brush.Color);
            }
        }

        private void OnRecentColorClick(object sender, MouseButtonEventArgs e)
        {
            if (sender is Border border && border.Background is SolidColorBrush brush)
            {
                _viewModel.CurrentColor = brush;
                _viewModel.CurrentColorHex = brush.Color.ToString();
                AddToRecentColors(brush.Color);
                GenerateSubColors(brush.Color);
                ApplyColor(brush.Color);
            }
        }

        private void OnFavoriteColorClick(object sender, MouseButtonEventArgs e)
        {
            if (sender is Border border && border.Background is SolidColorBrush brush)
            {
                _viewModel.CurrentColor = brush;
                _viewModel.CurrentColorHex = brush.Color.ToString();
                AddToRecentColors(brush.Color);
                GenerateSubColors(brush.Color);
                ApplyColor(brush.Color);
            }
        }

        private void OnAddToFavorites(object sender, RoutedEventArgs e)
        {
            var button = sender as Button;
            var border = VisualTreeHelper.GetParent(button) as Border;
            if (border != null && border.Background is SolidColorBrush brush)
            {
                if (!_viewModel.FavoriteColors.Contains(brush))
                {
                    _viewModel.FavoriteColors.Add(brush);
                }
            }
        }

        private void OnMorandiColorClick(object sender, MouseButtonEventArgs e)
        {
            if (sender is Border border && border.Background is SolidColorBrush brush)
            {
                _viewModel.CurrentColor = brush;
                _viewModel.CurrentColorHex = brush.Color.ToString();
                AddToRecentColors(brush.Color);
                GenerateSubColors(brush.Color);
                ApplyColor(brush.Color);
            }
        }

        private void InitializeColorSchemes()
        {
            // 添加配色方案1-100
            for (int i = 1; i <= 100; i++)
            {
                var colors = GetColorScheme(i);
                if (colors != null)
                {
                    ColorSchemeGroups.Add(new ColorGroup
                    {
                        GroupName = $"配色方案 {i}",
                        Colors = new ObservableCollection<ColorInfo>(colors)
                    });
                }
            }
        }

        private List<ColorInfo> GetColorScheme(int index)
        {
            switch (index)
            {
                case 1:
                    return new List<ColorInfo>
                    {
                        new ColorInfo { ColorName = "浅灰1", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#eeeeee")) },
                        new ColorInfo { ColorName = "浅灰2", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#eeeeee")) },
                        new ColorInfo { ColorName = "青绿1", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#00adb5")) },
                        new ColorInfo { ColorName = "青绿2", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#00adb5")) },
                        new ColorInfo { ColorName = "深灰1", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#393e46")) },
                        new ColorInfo { ColorName = "深灰2", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#393e46")) },
                        new ColorInfo { ColorName = "暗黑1", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#222831")) },
                        new ColorInfo { ColorName = "暗黑2", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#222831")) },
                        new ColorInfo { ColorName = "占位1", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#222831")) },
                        new ColorInfo { ColorName = "占位2", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#222831")) },
                        new ColorInfo { ColorName = "占位3", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#222831")) },
                        new ColorInfo { ColorName = "占位4", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#222831")) }
                    };
                case 2:
                    return new List<ColorInfo>
                    {
                        new ColorInfo { ColorName = "深紫1", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#6a2c70")) },
                        new ColorInfo { ColorName = "深紫2", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#6a2c70")) },
                        new ColorInfo { ColorName = "玫红1", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#b83b5e")) },
                        new ColorInfo { ColorName = "玫红2", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#b83b5e")) },
                        new ColorInfo { ColorName = "橙色1", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#f08a5d")) },
                        new ColorInfo { ColorName = "橙色2", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#f08a5d")) },
                        new ColorInfo { ColorName = "明黄1", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#f9ed69")) },
                        new ColorInfo { ColorName = "明黄2", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#f9ed69")) },
                        new ColorInfo { ColorName = "占位1", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#f9ed69")) },
                        new ColorInfo { ColorName = "占位2", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#f9ed69")) },
                        new ColorInfo { ColorName = "占位3", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#f9ed69")) },
                        new ColorInfo { ColorName = "占位4", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#f9ed69")) }
                    };
                case 3:
                    return new List<ColorInfo>
                    {
                        new ColorInfo { ColorName = "薄荷绿1", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#95e1d3")) },
                        new ColorInfo { ColorName = "薄荷绿2", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#95e1d3")) },
                        new ColorInfo { ColorName = "浅绿1", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#eaffd0")) },
                        new ColorInfo { ColorName = "浅绿2", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#eaffd0")) },
                        new ColorInfo { ColorName = "淡黄1", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#fce38a")) },
                        new ColorInfo { ColorName = "淡黄2", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#fce38a")) },
                        new ColorInfo { ColorName = "珊瑚红1", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#f38181")) },
                        new ColorInfo { ColorName = "珊瑚红2", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#f38181")) },
                        new ColorInfo { ColorName = "占位1", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#f38181")) },
                        new ColorInfo { ColorName = "占位2", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#f38181")) },
                        new ColorInfo { ColorName = "占位3", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#f38181")) },
                        new ColorInfo { ColorName = "占位4", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#f38181")) }
                    };
                case 4:
                    return new List<ColorInfo>
                    {
                        new ColorInfo { ColorName = "浅灰1", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#eaeaea")) },
                        new ColorInfo { ColorName = "浅灰2", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#eaeaea")) },
                        new ColorInfo { ColorName = "玫红1", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#ff2e63")) },
                        new ColorInfo { ColorName = "玫红2", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#ff2e63")) },
                        new ColorInfo { ColorName = "深灰1", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#252a34")) },
                        new ColorInfo { ColorName = "深灰2", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#252a34")) },
                        new ColorInfo { ColorName = "绿1", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#08d9d6")) },
                        new ColorInfo { ColorName = "青绿2", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#08d9d6")) },
                        new ColorInfo { ColorName = "占位1", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#08d9d6")) },
                        new ColorInfo { ColorName = "占位2", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#08d9d6")) },
                        new ColorInfo { ColorName = "占位3", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#08d9d6")) },
                        new ColorInfo { ColorName = "占位4", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#08d9d6")) }
                    };
                case 5:
                    return new List<ColorInfo>
                    {
                        new ColorInfo { ColorName = "玫红1", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#fc5185")) },
                        new ColorInfo { ColorName = "玫红2", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#fc5185")) },
                        new ColorInfo { ColorName = "浅灰1", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#f5f5f5")) },
                        new ColorInfo { ColorName = "浅灰2", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#f5f5f5")) },
                        new ColorInfo { ColorName = "青1", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#3fc1c9")) },
                        new ColorInfo { ColorName = "青绿2", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#3fc1c9")) },
                        new ColorInfo { ColorName = "深蓝1", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#364f6b")) },
                        new ColorInfo { ColorName = "深蓝2", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#364f6b")) },
                        new ColorInfo { ColorName = "占位1", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#364f6b")) },
                        new ColorInfo { ColorName = "占位2", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#364f6b")) },
                        new ColorInfo { ColorName = "占位3", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#364f6b")) },
                        new ColorInfo { ColorName = "占位4", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#364f6b")) }
                    };
                case 6:
                    return new List<ColorInfo>
                    {
                        new ColorInfo { ColorName = "淡黄1", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#ffffd2")) },
                        new ColorInfo { ColorName = "淡黄2", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#ffffd2")) },
                        new ColorInfo { ColorName = "粉红1", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#fcbad3")) },
                        new ColorInfo { ColorName = "粉红2", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#fcbad3")) },
                        new ColorInfo { ColorName = "淡紫1", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#aa96da")) },
                        new ColorInfo { ColorName = "淡紫2", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#aa96da")) },
                        new ColorInfo { ColorName = "淡蓝1", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#a8d8ea")) },
                        new ColorInfo { ColorName = "淡蓝2", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#a8d8ea")) },
                        new ColorInfo { ColorName = "占位1", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#a8d8ea")) },
                        new ColorInfo { ColorName = "占位2", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#a8d8ea")) },
                        new ColorInfo { ColorName = "占位3", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#a8d8ea")) },
                        new ColorInfo { ColorName = "占位4", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#a8d8ea")) }
                    };
                case 7:
                    return new List<ColorInfo>
                    {
                        new ColorInfo { ColorName = "青绿1", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#71c9ce")) },
                        new ColorInfo { ColorName = "青绿2", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#71c9ce")) },
                        new ColorInfo { ColorName = "浅青1", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#a6e3e9")) },
                        new ColorInfo { ColorName = "浅青2", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#a6e3e9")) },
                        new ColorInfo { ColorName = "淡青1", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#cbf1f5")) },
                        new ColorInfo { ColorName = "淡青2", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#cbf1f5")) },
                        new ColorInfo { ColorName = "极淡青1", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#e3fdfd")) },
                        new ColorInfo { ColorName = "极淡青2", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#e3fdfd")) },
                        new ColorInfo { ColorName = "占位1", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#e3fdfd")) },
                        new ColorInfo { ColorName = "占位2", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#e3fdfd")) },
                        new ColorInfo { ColorName = "占位3", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#e3fdfd")) },
                        new ColorInfo { ColorName = "占位4", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#e3fdfd")) }
                    };
                case 8:
                    return new List<ColorInfo>
                    {
                        new ColorInfo { ColorName = "深灰1", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#40514e")) },
                        new ColorInfo { ColorName = "深灰2", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#40514e")) },
                        new ColorInfo { ColorName = "青绿1", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#11999e")) },
                        new ColorInfo { ColorName = "青绿2", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#11999e")) },
                        new ColorInfo { ColorName = "亮青1", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#30e3ca")) },
                        new ColorInfo { ColorName = "亮青2", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#30e3ca")) },
                        new ColorInfo { ColorName = "淡青1", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#e4f9f5")) },
                        new ColorInfo { ColorName = "淡青2", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#e4f9f5")) },
                        new ColorInfo { ColorName = "占位1", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#e4f9f5")) },
                        new ColorInfo { ColorName = "占位2", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#e4f9f5")) },
                        new ColorInfo { ColorName = "占位3", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#e4f9f5")) },
                        new ColorInfo { ColorName = "占位4", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#e4f9f5")) }
                    };
                case 9:
                    return new List<ColorInfo>
                    {
                        new ColorInfo { ColorName = "灰紫1", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#8785a2")) },
                        new ColorInfo { ColorName = "灰紫2", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#8785a2")) },
                        new ColorInfo { ColorName = "浅灰1", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#f6f6f6")) },
                        new ColorInfo { ColorName = "浅灰2", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#f6f6f6")) },
                        new ColorInfo { ColorName = "淡粉1", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#ffe2e2")) },
                        new ColorInfo { ColorName = "淡粉2", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#ffe2e2")) },
                        new ColorInfo { ColorName = "浅粉1", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#ffc7c7")) },
                        new ColorInfo { ColorName = "浅粉2", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#ffc7c7")) },
                        new ColorInfo { ColorName = "占位1", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#ffc7c7")) },
                        new ColorInfo { ColorName = "占位2", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#ffc7c7")) },
                        new ColorInfo { ColorName = "占位3", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#ffc7c7")) },
                        new ColorInfo { ColorName = "占位4", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#ffc7c7")) }
                    };
                case 10:
                    return new List<ColorInfo>
                    {
                        new ColorInfo { ColorName = "深蓝1", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#112d4e")) },
                        new ColorInfo { ColorName = "深蓝2", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#112d4e")) },
                        new ColorInfo { ColorName = "蓝色1", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#3f72af")) },
                        new ColorInfo { ColorName = "蓝色2", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#3f72af")) },
                        new ColorInfo { ColorName = "淡蓝1", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#dbe2ef")) },
                        new ColorInfo { ColorName = "淡蓝2", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#dbe2ef")) },
                        new ColorInfo { ColorName = "浅灰1", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#f9f7f7")) },
                        new ColorInfo { ColorName = "浅灰2", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#f9f7f7")) },
                        new ColorInfo { ColorName = "占位1", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#f9f7f7")) },
                        new ColorInfo { ColorName = "占位2", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#f9f7f7")) },
                        new ColorInfo { ColorName = "占位3", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#f9f7f7")) },
                        new ColorInfo { ColorName = "占位4", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#f9f7f7")) }
                    };
                // ... 继续添加其他配色方案 ...
                default:
                    return null;
            }
        }

        private void ColorButton_Click(object sender, RoutedEventArgs e)
        {
            if (sender is Button button && button.ToolTip is string colorName)
            {
                var colorClass = _viewModel.Colors.FirstOrDefault(c => c.ColorName == colorName);
                if (colorClass != null)
                {
                    var color = colorClass.BackgroundColor.Color;
                    _viewModel.CurrentColor = new SolidColorBrush(color);
                    _viewModel.CurrentColorHex = color.ToString();
                    AddToRecentColors(color);
                    GenerateSubColors(color);
                    ApplyColor(color);
                }
            }
        }
    }

    public class ColorInfo
    {
        public string ColorName { get; set; }
        public SolidColorBrush ColorBrush { get; set; }
    }

    public class ColorGroup
    {
        public string GroupName { get; set; }
        public ObservableCollection<ColorInfo> Colors { get; set; }
    }

    public class ColorViewModel : INotifyPropertyChanged
    {
        private List<ColorClass> _colors;
        public List<ColorClass> Colors
        {
            get => _colors;
            set
            {
                _colors = value;
                OnPropertyChanged(nameof(Colors));
            }
        }

        private ObservableCollection<ColorClass> _subColors;
        public ObservableCollection<ColorClass> SubColors
        {
            get => _subColors;
            set
            {
                _subColors = value;
                OnPropertyChanged(nameof(SubColors));
            }
        }

        public ObservableCollection<SolidColorBrush> RecentColors { get; set; }
        public ObservableCollection<SolidColorBrush> FavoriteColors { get; set; }

        private SolidColorBrush _currentColor;
        public SolidColorBrush CurrentColor
        {
            get => _currentColor;
            set
            {
                _currentColor = value;
                OnPropertyChanged(nameof(CurrentColor));
            }
        }

        private string _currentColorHex;
        public string CurrentColorHex
        {
            get => _currentColorHex;
            set
            {
                _currentColorHex = value;
                OnPropertyChanged(nameof(CurrentColorHex));
            }
        }

        private double _currentOpacity = 100;
        public double CurrentOpacity
        {
            get => _currentOpacity;
            set
            {
                if (_currentOpacity != value)
                {
                    _currentOpacity = value;
                    OnPropertyChanged(nameof(CurrentOpacity));
                    UpdateCurrentColorWithOpacity();
                }
            }
        }

        private void UpdateCurrentColorWithOpacity()
        {
            if (_currentColor != null)
            {
                var color = _currentColor.Color;
                color.A = (byte)(_currentOpacity * 255 / 100);
                _currentColor = new SolidColorBrush(color);
                OnPropertyChanged(nameof(CurrentColor));
            }
        }

        public event PropertyChangedEventHandler PropertyChanged;

        protected virtual void OnPropertyChanged(string propertyName)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }

        private ObservableCollection<ColorGroup> _morandiColorGroups;
        public ObservableCollection<ColorGroup> MorandiColorGroups
        {
            get => _morandiColorGroups;
            set
            {
                _morandiColorGroups = value;
                OnPropertyChanged(nameof(MorandiColorGroups));
            }
        }

        private ObservableCollection<ColorGroup> _chineseColorGroups;
        public ObservableCollection<ColorGroup> ChineseColorGroups
        {
            get => _chineseColorGroups;
            set
            {
                _chineseColorGroups = value;
                OnPropertyChanged(nameof(ChineseColorGroups));
            }
        }

        private ObservableCollection<ColorGroup> _macaronColorGroups;
        public ObservableCollection<ColorGroup> MacaronColorGroups
        {
            get => _macaronColorGroups;
            set
            {
                _macaronColorGroups = value;
                OnPropertyChanged(nameof(MacaronColorGroups));
            }
        }

        private ObservableCollection<ColorGroup> _colorSchemeGroups;
        public ObservableCollection<ColorGroup> ColorSchemeGroups
        {
            get => _colorSchemeGroups;
            set
            {
                _colorSchemeGroups = value;
                OnPropertyChanged(nameof(ColorSchemeGroups));
            }
        }

        public ColorViewModel()
        {
            SubColors = new ObservableCollection<ColorClass>();
            RecentColors = new ObservableCollection<SolidColorBrush>();
            FavoriteColors = new ObservableCollection<SolidColorBrush>();
            ColorSchemeGroups = new ObservableCollection<ColorGroup>();
            InitializeMorandiColors();
            InitializeChineseColors();
            InitializeMacaronColors();
            InitializeColorSchemes();
        }

        private void InitializeMorandiColors()
        {
            MorandiColorGroups = new ObservableCollection<ColorGroup>
            {
                new ColorGroup
                {
                    GroupName = "暖色系",
                    Colors = new ObservableCollection<ColorInfo>
                    {
                        new ColorInfo { ColorName = "暖粉", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#E1C6BC")) },
                        new ColorInfo { ColorName = "浅褐", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#D7B2A8")) },
                        new ColorInfo { ColorName = "灰棕", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#C4AEA1")) },
                        new ColorInfo { ColorName = "浅粉", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#E9D7CD")) },
                        new ColorInfo { ColorName = "浅灰棕", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#D1BCAF")) },
                        new ColorInfo { ColorName = "米白", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#EBDED6")) },
                        new ColorInfo { ColorName = "暖灰", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#D4B8B0")) },
                        new ColorInfo { ColorName = "深暖粉", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#C5A99D")) },
                        new ColorInfo { ColorName = "珊瑚粉", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#E6C3BA")) },
                        new ColorInfo { ColorName = "玫瑰灰", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#DDCBC4")) },
                        new ColorInfo { ColorName = "暖沙", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#E4D4CB")) },
                        new ColorInfo { ColorName = "奶茶色", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#DBC8BE")) }
                    }
                },
                new ColorGroup
                {
                    GroupName = "暖棕系",
                    Colors = new ObservableCollection<ColorInfo>
                    {
                        new ColorInfo { ColorName = "奶茶棕", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#C8B4A4")) },
                        new ColorInfo { ColorName = "焦糖棕", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#BE9B7B")) },
                        new ColorInfo { ColorName = "杏仁棕", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#D4C4B7")) },
                        new ColorInfo { ColorName = "驼色", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#C1A894")) },
                        new ColorInfo { ColorName = "沙褐", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#CEB6A2")) },
                        new ColorInfo { ColorName = "咖啡", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#B49C8C")) },
                        new ColorInfo { ColorName = "浅驼", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#D6C6B9")) },
                        new ColorInfo { ColorName = "卡其", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#C8B8A9")) },
                        new ColorInfo { ColorName = "深棕", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#B4A294")) },
                        new ColorInfo { ColorName = "暖棕", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#CCBCAF")) },
                        new ColorInfo { ColorName = "沙棕", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#D1C2B3")) },
                        new ColorInfo { ColorName = "米棕", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#E0D3C5")) }
                    }
                },
                new ColorGroup
                {
                    GroupName = "冷色系",
                    Colors = new ObservableCollection<ColorInfo>
                    {
                        new ColorInfo { ColorName = "灰绿", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#9DAFA7")) },
                        new ColorInfo { ColorName = "浅灰蓝", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#C7CDD1")) },
                        new ColorInfo { ColorName = "灰蓝", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#B7BABE")) },
                        new ColorInfo { ColorName = "浅灰绿", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#B2BEB5")) },
                        new ColorInfo { ColorName = "淡灰", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#D3D4D6")) },
                        new ColorInfo { ColorName = "浅灰白", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#EAEAEA")) },
                        new ColorInfo { ColorName = "深灰绿", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#A4B5AE")) },
                        new ColorInfo { ColorName = "深灰蓝", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#B8C4C9")) },
                        new ColorInfo { ColorName = "青灰", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#B8C4C1")) },
                        new ColorInfo { ColorName = "蓝灰", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#C4CCD2")) },
                        new ColorInfo { ColorName = "冷灰", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#D1D5D8")) },
                        new ColorInfo { ColorName = "银灰", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#E2E4E6")) }
                    }
                },
                new ColorGroup
                {
                    GroupName = "蓝灰系",
                    Colors = new ObservableCollection<ColorInfo>
                    {
                        new ColorInfo { ColorName = "雾蓝", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#B0C4DE")) },
                        new ColorInfo { ColorName = "烟灰蓝", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#A2B5CD")) },
                        new ColorInfo { ColorName = "淡蓝灰", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#C6D4E2")) },
                        new ColorInfo { ColorName = "深蓝灰", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#8CA6B5")) },
                        new ColorInfo { ColorName = "钢青", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#B2BCC0")) },
                        new ColorInfo { ColorName = "雨灰", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#D1DCE3")) },
                        new ColorInfo { ColorName = "青灰", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#A4B9C4")) },
                        new ColorInfo { ColorName = "蓝铅", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#98A8B5")) },
                        new ColorInfo { ColorName = "湖蓝灰", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#B4C5D4")) },
                        new ColorInfo { ColorName = "天蓝灰", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#C5D5E0")) },
                        new ColorInfo { ColorName = "浅蓝灰", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#D5E0E6")) },
                        new ColorInfo { ColorName = "银蓝", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#E0E6EC")) }
                    }
                },
                new ColorGroup
                {
                    GroupName = "绿灰系",
                    Colors = new ObservableCollection<ColorInfo>
                    {
                        new ColorInfo { ColorName = "薄荷灰", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#A8C1B4")) },
                        new ColorInfo { ColorName = "苔绿灰", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#95A5A0")) },
                        new ColorInfo { ColorName = "橄榄灰", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#A3B4A2")) },
                        new ColorInfo { ColorName = "深绿灰", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#8B9E8E")) },
                        new ColorInfo { ColorName = "青瓷灰", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#B4C4BC")) },
                        new ColorInfo { ColorName = "草灰", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#A6B5A3")) },
                        new ColorInfo { ColorName = "银杏灰", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#BEC8B7")) },
                        new ColorInfo { ColorName = "豆绿灰", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#9CAB98")) },
                        new ColorInfo { ColorName = "抹茶灰", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#B5C4B1")) },
                        new ColorInfo { ColorName = "青草灰", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#C4D0BE")) },
                        new ColorInfo { ColorName = "嫩绿灰", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#D0DBCA")) },
                        new ColorInfo { ColorName = "银绿", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#DCE4D7")) }
                    }
                },
                new ColorGroup
                {
                    GroupName = "紫色系",
                    Colors = new ObservableCollection<ColorInfo>
                    {
                        new ColorInfo { ColorName = "浅紫", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#C8B6C1")) },
                        new ColorInfo { ColorName = "灰紫", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#B4A3B1")) },
                        new ColorInfo { ColorName = "淡紫", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#D6CBD3")) },
                        new ColorInfo { ColorName = "浅灰紫", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#E4DBE0")) },
                        new ColorInfo { ColorName = "深紫", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#A18EA4")) },
                        new ColorInfo { ColorName = "中紫", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#BFB1C1")) },
                        new ColorInfo { ColorName = "藕紫", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#DCD3D8")) },
                        new ColorInfo { ColorName = "紫灰", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#B8A9B9")) },
                        new ColorInfo { ColorName = "丁香紫", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#C5B6C6")) },
                        new ColorInfo { ColorName = "紫灰", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#D2C5D3")) },
                        new ColorInfo { ColorName = "淡雅紫", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#DED4DF")) },
                        new ColorInfo { ColorName = "银紫", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#E7E0E8")) }
                    }
                },
                new ColorGroup
                {
                    GroupName = "粉色系",
                    Colors = new ObservableCollection<ColorInfo>
                    {
                        new ColorInfo { ColorName = "浅粉", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#E8C6C6")) },
                        new ColorInfo { ColorName = "珊瑚粉", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#E8D3D1")) },
                        new ColorInfo { ColorName = "贝壳粉", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#EED8D3")) },
                        new ColorInfo { ColorName = "淡粉", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#F2E2DE")) },
                        new ColorInfo { ColorName = "浅珊瑚", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#F5E6E6")) },
                        new ColorInfo { ColorName = "雾粉", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#F8EFEF")) },
                        new ColorInfo { ColorName = "深粉", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#D9B5B5")) },
                        new ColorInfo { ColorName = "玫瑰粉", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#E6CCCC")) },
                        new ColorInfo { ColorName = "蜜桃粉", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#EBD4D4")) },
                        new ColorInfo { ColorName = "樱花粉", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#F0DCDC")) },
                        new ColorInfo { ColorName = "淡雅粉", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#F5E5E5")) },
                        new ColorInfo { ColorName = "银粉", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#F9EDED")) }
                    }
                },
                new ColorGroup
                {
                    GroupName = "灰色系",
                    Colors = new ObservableCollection<ColorInfo>
                    {
                        new ColorInfo { ColorName = "深灰", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#B7B1A5")) },
                        new ColorInfo { ColorName = "中灰", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#C5C0B7")) },
                        new ColorInfo { ColorName = "浅灰", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#D0CCC4")) },
                        new ColorInfo { ColorName = "珍珠", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#DCD8D3")) },
                        new ColorInfo { ColorName = "银灰", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#E6E3E0")) },
                        new ColorInfo { ColorName = "雾灰", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#F0EEEC")) },
                        new ColorInfo { ColorName = "烟灰", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#AAA49D")) },
                        new ColorInfo { ColorName = "石灰", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#C2BDB6")) },
                        new ColorInfo { ColorName = "暖灰", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#D1CDC7")) },
                        new ColorInfo { ColorName = "珠光灰", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#DFDCD8")) },
                        new ColorInfo { ColorName = "贝壳灰", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#E8E6E3")) },
                        new ColorInfo { ColorName = "月光灰", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#F2F1EF")) }
                    }
                },
                new ColorGroup
                {
                    GroupName = "米白系",
                    Colors = new ObservableCollection<ColorInfo>
                    {
                        new ColorInfo { ColorName = "象牙白", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#FFFBF0")) },
                        new ColorInfo { ColorName = "珍珠白", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#F7F4ED")) },
                        new ColorInfo { ColorName = "米白", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#F5F2E9")) },
                        new ColorInfo { ColorName = "乳白", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#F3EFE1")) },
                        new ColorInfo { ColorName = "灰白", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#F1EDE4")) },
                        new ColorInfo { ColorName = "浅米", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#F9F4DC")) },
                        new ColorInfo { ColorName = "奶白", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#F8F4E9")) },
                        new ColorInfo { ColorName = "素白", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#F6F5EC")) },
                        new ColorInfo { ColorName = "贝壳白", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#F9F6F0")) },
                        new ColorInfo { ColorName = "月光白", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#F8F7F2")) },
                        new ColorInfo { ColorName = "银白", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#F7F6F4")) },
                        new ColorInfo { ColorName = "雾白", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#F6F6F6")) }
                    }
                }
            };
        }

        private void InitializeChineseColors()
        {
            ChineseColorGroups = new ObservableCollection<ColorGroup>
            {
                new ColorGroup
                {
                    GroupName = "纯白系",
                    Colors = new ObservableCollection<ColorInfo>
                    {
                        new ColorInfo { ColorName = "精白", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#FFFFFF")) },
                        new ColorInfo { ColorName = "银白", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#E9E7EF")) },
                        new ColorInfo { ColorName = "铅白", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#F0F0F4")) },
                        new ColorInfo { ColorName = "霜色", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#E9F1F6")) },
                        new ColorInfo { ColorName = "雪白", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#F0FCFF")) },
                        new ColorInfo { ColorName = "莹白", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#E3F9FD")) },
                        new ColorInfo { ColorName = "月白", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#D6ECF0")) },
                        new ColorInfo { ColorName = "象牙白", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#FFFBF0")) },
                        new ColorInfo { ColorName = "缟", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#F2ECDE")) },
                        new ColorInfo { ColorName = "鱼肚白", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#FCEFE8")) },
                        new ColorInfo { ColorName = "白粉", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#FFF2DF")) },
                        new ColorInfo { ColorName = "荼白", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#F3F9F1")) }
                    }
                },
                new ColorGroup
                {
                    GroupName = "灰色系",
                    Colors = new ObservableCollection<ColorInfo>
                    {
                        new ColorInfo { ColorName = "鸭卵青", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#E0EEE8")) },
                        new ColorInfo { ColorName = "素", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#E0F0E9")) },
                        new ColorInfo { ColorName = "青白", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#C0EBD7")) },
                        new ColorInfo { ColorName = "蟹壳青", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#BBCDC5")) },
                        new ColorInfo { ColorName = "花白", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#C2CCD0")) },
                        new ColorInfo { ColorName = "老银", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#BACAC6")) },
                        new ColorInfo { ColorName = "灰色", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#808080")) },
                        new ColorInfo { ColorName = "苍色", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#75878A")) },
                        new ColorInfo { ColorName = "水色", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#88ADA6")) },
                        new ColorInfo { ColorName = "黝", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#6B6882")) },
                        new ColorInfo { ColorName = "乌色", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#725E82")) },
                        new ColorInfo { ColorName = "玄青", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#3D3B4F")) }
                    }
                },
                new ColorGroup
                {
                    GroupName = "黑色系",
                    Colors = new ObservableCollection<ColorInfo>
                    {
                        new ColorInfo { ColorName = "乌黑", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#392F41")) },
                        new ColorInfo { ColorName = "黎", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#75664D")) },
                        new ColorInfo { ColorName = "黧", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#5D513C")) },
                        new ColorInfo { ColorName = "黝黑", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#665757")) },
                        new ColorInfo { ColorName = "缁色", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#493131")) },
                        new ColorInfo { ColorName = "煤黑", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#312520")) },
                        new ColorInfo { ColorName = "漆黑", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#161823")) },
                        new ColorInfo { ColorName = "黑色", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#000000")) },
                        new ColorInfo { ColorName = "玄色", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#622A1D")) },
                        new ColorInfo { ColorName = "墨灰", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#758A99")) },
                        new ColorInfo { ColorName = "��色", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#50616D")) },
                        new ColorInfo { ColorName = "鸦青", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#424C50")) }
                    }
                },
                new ColorGroup
                {
                    GroupName = "黄色系",
                    Colors = new ObservableCollection<ColorInfo>
                    {
                        new ColorInfo { ColorName = "樱草色", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#EAFF56")) },
                        new ColorInfo { ColorName = "鹅黄", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#FFF143")) },
                        new ColorInfo { ColorName = "鸭黄", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#FAFF72")) },
                        new ColorInfo { ColorName = "杏黄", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#FFA631")) },
                        new ColorInfo { ColorName = "橙黄", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#FFA400")) },
                        new ColorInfo { ColorName = "橙色", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#FA8C35")) },
                        new ColorInfo { ColorName = "杏红", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#FF8C31")) },
                        new ColorInfo { ColorName = "橘黄", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#FF8936")) },
                        new ColorInfo { ColorName = "橘红", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#FF7500")) },
                        new ColorInfo { ColorName = "藤黄", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#FFB61E")) },
                        new ColorInfo { ColorName = "姜黄", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#FFC773")) },
                        new ColorInfo { ColorName = "雌黄", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#FFC64B")) }
                    }
                },
                new ColorGroup
                {
                    GroupName = "金色系",
                    Colors = new ObservableCollection<ColorInfo>
                    {
                        new ColorInfo { ColorName = "赤金", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#F2BE45")) },
                        new ColorInfo { ColorName = "缃色", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#F0C239")) },
                        new ColorInfo { ColorName = "雄黄", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#E9BB1D")) },
                        new ColorInfo { ColorName = "秋香色", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#D9B611")) },
                        new ColorInfo { ColorName = "金色", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#EACD76")) },
                        new ColorInfo { ColorName = "牙色", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#EEDEB0")) },
                        new ColorInfo { ColorName = "枯黄", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#D3B17D")) },
                        new ColorInfo { ColorName = "黄栌", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#E29C45")) },
                        new ColorInfo { ColorName = "乌金", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#A78E44")) },
                        new ColorInfo { ColorName = "昏黄", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#C89B40")) },
                        new ColorInfo { ColorName = "棕黄", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#AE7000")) },
                        new ColorInfo { ColorName = "琥珀", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#CA6924")) }
                    }
                },
                new ColorGroup
                {
                    GroupName = "棕色系",
                    Colors = new ObservableCollection<ColorInfo>
                    {
                        new ColorInfo { ColorName = "棕色", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#B25D25")) },
                        new ColorInfo { ColorName = "茶色", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#B35C44")) },
                        new ColorInfo { ColorName = "棕红", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#9B4400")) },
                        new ColorInfo { ColorName = "赭", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#9C5333")) },
                        new ColorInfo { ColorName = "色", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#A88462")) },
                        new ColorInfo { ColorName = "秋色", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#896C39")) },
                        new ColorInfo { ColorName = "棕绿", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#827100")) },
                        new ColorInfo { ColorName = "褐色", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#6E511E")) },
                        new ColorInfo { ColorName = "棕黑", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#7C4B00")) },
                        new ColorInfo { ColorName = "赭色", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#955539")) },
                        new ColorInfo { ColorName = "赭石", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#845A33")) },
                        new ColorInfo { ColorName = "黯", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#41555D")) }
                    }
                },
                new ColorGroup
                {
                    GroupName = "绿色系一",
                    Colors = new ObservableCollection<ColorInfo>
                    {
                        new ColorInfo { ColorName = "松花色", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#BCE672")) },
                        new ColorInfo { ColorName = "柳黄", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#C9DD22")) },
                        new ColorInfo { ColorName = "嫩绿", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#BDDD22")) },
                        new ColorInfo { ColorName = "绿", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#AFDD22")) },
                        new ColorInfo { ColorName = "葱黄", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#A3D900")) },
                        new ColorInfo { ColorName = "葱绿", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#9ED900")) },
                        new ColorInfo { ColorName = "豆绿", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#9ED048")) },
                        new ColorInfo { ColorName = "豆青", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#96CE54")) },
                        new ColorInfo { ColorName = "油绿", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#00BC12")) },
                        new ColorInfo { ColorName = "葱倩", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#0EB83A")) },
                        new ColorInfo { ColorName = "葱青", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#0EB83A")) },
                        new ColorInfo { ColorName = "青葱", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#0AA344")) }
                    }
                },
                new ColorGroup
                {
                    GroupName = "绿色系二",
                    Colors = new ObservableCollection<ColorInfo>
                    {
                        new ColorInfo { ColorName = "石绿", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#16A951")) },
                        new ColorInfo { ColorName = "松柏绿", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#21A675")) },
                        new ColorInfo { ColorName = "松花绿", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#057748")) },
                        new ColorInfo { ColorName = "绿沈", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#0C8918")) },
                        new ColorInfo { ColorName = "绿色", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#00E500")) },
                        new ColorInfo { ColorName = "草绿", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#40DE5A")) },
                        new ColorInfo { ColorName = "青翠", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#00E079")) },
                        new ColorInfo { ColorName = "青色", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#00E09E")) },
                        new ColorInfo { ColorName = "翡翠色", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#3DE1AD")) },
                        new ColorInfo { ColorName = "碧绿", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#2ADD9C")) },
                        new ColorInfo { ColorName = "玉色", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#2EDFA3")) },
                        new ColorInfo { ColorName = "缥", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#7FECAD")) }
                    }
                },
                new ColorGroup
                {
                    GroupName = "青色系",
                    Colors = new ObservableCollection<ColorInfo>
                    {
                        new ColorInfo { ColorName = "艾绿", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#A4E2C6")) },
                        new ColorInfo { ColorName = "石青", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#7BCFA6")) },
                        new ColorInfo { ColorName = "碧色", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#1BD1A5")) },
                        new ColorInfo { ColorName = "青碧", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#48C0A3")) },
                        new ColorInfo { ColorName = "铜绿", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#549688")) },
                        new ColorInfo { ColorName = "竹青", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#789262")) },
                        new ColorInfo { ColorName = "墨灰", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#758A99")) },
                        new ColorInfo { ColorName = "墨色", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#50616D")) },
                        new ColorInfo { ColorName = "鸦青", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#424C50")) },
                        new ColorInfo { ColorName = "黯", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#41555D")) },
                        new ColorInfo { ColorName = "玄青", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#3D3B4F")) },
                        new ColorInfo { ColorName = "玄色", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#622A1D")) }
                    }
                },
                new ColorGroup
                {
                    GroupName = "红色系一",
                    Colors = new ObservableCollection<ColorInfo>
                    {
                        new ColorInfo { ColorName = "朱砂", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#FF461F")) },
                        new ColorInfo { ColorName = "火红", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#FF2D51")) },
                        new ColorInfo { ColorName = "朱膘", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#F36838")) },
                        new ColorInfo { ColorName = "妃色", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#ED5736")) },
                        new ColorInfo { ColorName = "洋红", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#FF4777")) },
                        new ColorInfo { ColorName = "品红", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#F00056")) },
                        new ColorInfo { ColorName = "粉红", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#FFB3A7")) },
                        new ColorInfo { ColorName = "桃红", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#F47983")) },
                        new ColorInfo { ColorName = "海棠红", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#DB5A6B")) },
                        new ColorInfo { ColorName = "樱桃色", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#C93756")) },
                        new ColorInfo { ColorName = "酡颜", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#F9906F")) },
                        new ColorInfo { ColorName = "银红", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#F05654")) }
                    }
                },
                new ColorGroup
                {
                    GroupName = "红色系二",
                    Colors = new ObservableCollection<ColorInfo>
                    {
                        new ColorInfo { ColorName = "大红", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#FF2121")) },
                        new ColorInfo { ColorName = "石榴红", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#F20C00")) },
                        new ColorInfo { ColorName = "绛紫", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#8C4356")) },
                        new ColorInfo { ColorName = "绯红", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#C83C23")) },
                        new ColorInfo { ColorName = "胭脂", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#9D2933")) },
                        new ColorInfo { ColorName = "朱红", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#FF4C00")) },
                        new ColorInfo { ColorName = "丹", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#FF4E20")) },
                        new ColorInfo { ColorName = "彤", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#F35336")) },
                        new ColorInfo { ColorName = "酡红", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#DC3023")) },
                        new ColorInfo { ColorName = "炎", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#FF3300")) },
                        new ColorInfo { ColorName = "茜色", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#CB3A56")) },
                        new ColorInfo { ColorName = "绾", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#A98175")) }
                    }
                },
                new ColorGroup
                {
                    GroupName = "红色系三",
                    Colors = new ObservableCollection<ColorInfo>
                    {
                        new ColorInfo { ColorName = "檀", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#B36D61")) },
                        new ColorInfo { ColorName = "嫣红", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#EF7A82")) },
                        new ColorInfo { ColorName = "洋红", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#FF0097")) },
                        new ColorInfo { ColorName = "枣红", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#C32136")) },
                        new ColorInfo { ColorName = "殷红", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#BE002F")) },
                        new ColorInfo { ColorName = "赫赤", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#C91F37")) },
                        new ColorInfo { ColorName = "银朱", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#BF242A")) },
                        new ColorInfo { ColorName = "赤", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#C3272B")) },
                        new ColorInfo { ColorName = "胭脂", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#9D2933")) },
                        new ColorInfo { ColorName = "栗色", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#60281E")) },
                        new ColorInfo { ColorName = "玄色", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#622A1D")) },
                        new ColorInfo { ColorName = "黑色", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#000000")) }
                    }
                }
            };
        }

        private void InitializeMacaronColors()
        {
            MacaronColorGroups = new ObservableCollection<ColorGroup>
            {
                new ColorGroup
                {
                    GroupName = "甜美系",
                    Colors = new ObservableCollection<ColorInfo>
                    {
                        new ColorInfo { ColorName = "草莓奶昔", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#FFB5C5")) },
                        new ColorInfo { ColorName = "蜜桃乌龙", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#FADADD")) },
                        new ColorInfo { ColorName = "奶油粉", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#FFE4E1")) },
                        new ColorInfo { ColorName = "樱花粉", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#FFD1DC")) },
                        new ColorInfo { ColorName = "蔓越莓", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#FFB6C1")) },
                        new ColorInfo { ColorName = "覆盆子", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#FFB0B0")) },
                        new ColorInfo { ColorName = "棉花糖", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#FFE1E8")) },
                        new ColorInfo { ColorName = "果冻粉", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#FFCAD4")) },
                        new ColorInfo { ColorName = "珊瑚粉", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#FFB7B9")) },
                        new ColorInfo { ColorName = "水蜜桃", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#FFCDC4")) },
                        new ColorInfo { ColorName = "玫瑰粉", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#FFBEC8")) },
                        new ColorInfo { ColorName = "浅粉红", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#FFD9E0")) }
                    }
                },
                new ColorGroup
                {
                    GroupName = "清新系",
                    Colors = new ObservableCollection<ColorInfo>
                    {
                        new ColorInfo { ColorName = "薄荷绿", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#C5E7D1")) },
                        new ColorInfo { ColorName = "抹茶拿铁", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#D8E4D4")) },
                        new ColorInfo { ColorName = "青柠", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#E8F3D6")) },
                        new ColorInfo { ColorName = "开心果", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#DDF1D6")) },
                        new ColorInfo { ColorName = "牛油果", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#C7E5C8")) },
                        new ColorInfo { ColorName = "青瓜", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#DEF1DD")) },
                        new ColorInfo { ColorName = "绿茶", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#E6F3E3")) },
                        new ColorInfo { ColorName = "柠檬", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#F1F6D0")) },
                        new ColorInfo { ColorName = "抹茶冰", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#D4E5D6")) },
                        new ColorInfo { ColorName = "青草", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#E1ECD0")) },
                        new ColorInfo { ColorName = "豆绿", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#DCE8D5")) },
                        new ColorInfo { ColorName = "嫩芽绿", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#E8F3E2")) }
                    }
                },
                new ColorGroup
                {
                    GroupName = "梦幻系",
                    Colors = new ObservableCollection<ColorInfo>
                    {
                        new ColorInfo { ColorName = "蓝莓慕斯", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#B5C4E3")) },
                        new ColorInfo { ColorName = "薰衣草", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#D8D1E3")) },
                        new ColorInfo { ColorName = "葡萄奶昔", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#E3D7E8")) },
                        new ColorInfo { ColorName = "蝶豆花", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#C9D7E8")) },
                        new ColorInfo { ColorName = "星空蓝", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#BCD4E6")) },
                        new ColorInfo { ColorName = "海盐", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#D6E4E8")) },
                        new ColorInfo { ColorName = "蓝风铃", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#CCE2E8")) },
                        new ColorInfo { ColorName = "天空蓝", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#D4E5E9")) },
                        new ColorInfo { ColorName = "紫藤花", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#E2D9E8")) },
                        new ColorInfo { ColorName = "梦幻紫", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#E5DEE9")) },
                        new ColorInfo { ColorName = "幻彩蓝", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#D9E5EC")) },
                        new ColorInfo { ColorName = "云雾蓝", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#E1EBF0")) }
                    }
                },
                new ColorGroup
                {
                    GroupName = "温暖系",
                    Colors = new ObservableCollection<ColorInfo>
                    {
                        new ColorInfo { ColorName = "香草奶油", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#FFF0D9")) },
                        new ColorInfo { ColorName = "焦糖布丁", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#FFE4C4")) },
                        new ColorInfo { ColorName = "蜂蜜", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#FFE5B4")) },
                        new ColorInfo { ColorName = "杏仁", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#FFEFD5")) },
                        new ColorInfo { ColorName = "奶茶", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#F3E5D0")) },
                        new ColorInfo { ColorName = "拿铁", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#E8D3BB")) },
                        new ColorInfo { ColorName = "榛果", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#F4DFC8")) },
                        new ColorInfo { ColorName = "太妃糖", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#F1E3D3")) },
                        new ColorInfo { ColorName = "牛奶糖", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#F8E8D7")) },
                        new ColorInfo { ColorName = "奶油", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#FFF2E2")) },
                        new ColorInfo { ColorName = "椰奶", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#FFF4E8")) },
                        new ColorInfo { ColorName = "米糖", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#FFF6EC")) }
                    }
                }
            };
        }

        private void InitializeColorSchemes()
        {
            // 添加配色方案1-100
            for (int i = 1; i <= 100; i++)
            {
                var colors = GetColorScheme(i);
                if (colors != null)
                {
                    ColorSchemeGroups.Add(new ColorGroup
                    {
                        GroupName = $"配色方案 {i}",
                        Colors = new ObservableCollection<ColorInfo>(colors)
                    });
                }
            }
        }

        private List<ColorInfo> GetColorScheme(int index)
        {
            switch (index)
            {
                case 1:
                    return new List<ColorInfo>
                    {
                        new ColorInfo { ColorName = "浅灰1", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#eeeeee")) },
                        new ColorInfo { ColorName = "浅灰2", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#eeeeee")) },
                        new ColorInfo { ColorName = "青绿1", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#00adb5")) },
                        new ColorInfo { ColorName = "青绿2", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#00adb5")) },
                        new ColorInfo { ColorName = "深灰1", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#393e46")) },
                        new ColorInfo { ColorName = "深灰2", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#393e46")) },
                        new ColorInfo { ColorName = "暗黑1", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#222831")) },
                        new ColorInfo { ColorName = "暗黑2", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#222831")) },
                        new ColorInfo { ColorName = "占位1", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#222831")) },
                        new ColorInfo { ColorName = "占位2", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#222831")) },
                        new ColorInfo { ColorName = "占位3", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#222831")) },
                        new ColorInfo { ColorName = "占位4", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#222831")) }
                    };
                case 2:
                    return new List<ColorInfo>
                    {
                        new ColorInfo { ColorName = "深紫1", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#6a2c70")) },
                        new ColorInfo { ColorName = "深紫2", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#6a2c70")) },
                        new ColorInfo { ColorName = "玫红1", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#b83b5e")) },
                        new ColorInfo { ColorName = "玫红2", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#b83b5e")) },
                        new ColorInfo { ColorName = "橙色1", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#f08a5d")) },
                        new ColorInfo { ColorName = "橙色2", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#f08a5d")) },
                        new ColorInfo { ColorName = "明黄1", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#f9ed69")) },
                        new ColorInfo { ColorName = "明黄2", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#f9ed69")) },
                        new ColorInfo { ColorName = "占位1", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#f9ed69")) },
                        new ColorInfo { ColorName = "占位2", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#f9ed69")) },
                        new ColorInfo { ColorName = "占位3", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#f9ed69")) },
                        new ColorInfo { ColorName = "占位4", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#f9ed69")) }
                    };
                case 3:
                    return new List<ColorInfo>
                    {
                        new ColorInfo { ColorName = "薄荷绿1", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#95e1d3")) },
                        new ColorInfo { ColorName = "薄荷绿2", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#95e1d3")) },
                        new ColorInfo { ColorName = "浅绿1", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#eaffd0")) },
                        new ColorInfo { ColorName = "浅绿2", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#eaffd0")) },
                        new ColorInfo { ColorName = "淡黄1", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#fce38a")) },
                        new ColorInfo { ColorName = "淡黄2", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#fce38a")) },
                        new ColorInfo { ColorName = "珊瑚红1", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#f38181")) },
                        new ColorInfo { ColorName = "珊瑚红2", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#f38181")) },
                        new ColorInfo { ColorName = "占位1", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#f38181")) },
                        new ColorInfo { ColorName = "占位2", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#f38181")) },
                        new ColorInfo { ColorName = "占位3", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#f38181")) },
                        new ColorInfo { ColorName = "占位4", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#f38181")) }
                    };
                case 4:
                    return new List<ColorInfo>
                    {
                        new ColorInfo { ColorName = "浅灰1", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#eaeaea")) },
                        new ColorInfo { ColorName = "浅灰2", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#eaeaea")) },
                        new ColorInfo { ColorName = "玫红1", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#ff2e63")) },
                        new ColorInfo { ColorName = "玫红2", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#ff2e63")) },
                        new ColorInfo { ColorName = "深灰1", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#252a34")) },
                        new ColorInfo { ColorName = "深灰2", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#252a34")) },
                        new ColorInfo { ColorName = "绿1", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#08d9d6")) },
                        new ColorInfo { ColorName = "青绿2", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#08d9d6")) },
                        new ColorInfo { ColorName = "占位1", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#08d9d6")) },
                        new ColorInfo { ColorName = "占位2", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#08d9d6")) },
                        new ColorInfo { ColorName = "占位3", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#08d9d6")) },
                        new ColorInfo { ColorName = "占位4", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#08d9d6")) }
                    };
                case 5:
                    return new List<ColorInfo>
                    {
                        new ColorInfo { ColorName = "玫红1", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#fc5185")) },
                        new ColorInfo { ColorName = "玫红2", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#fc5185")) },
                        new ColorInfo { ColorName = "浅灰1", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#f5f5f5")) },
                        new ColorInfo { ColorName = "浅灰2", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#f5f5f5")) },
                        new ColorInfo { ColorName = "青1", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#3fc1c9")) },
                        new ColorInfo { ColorName = "青绿2", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#3fc1c9")) },
                        new ColorInfo { ColorName = "深蓝1", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#364f6b")) },
                        new ColorInfo { ColorName = "深蓝2", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#364f6b")) },
                        new ColorInfo { ColorName = "占位1", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#364f6b")) },
                        new ColorInfo { ColorName = "占位2", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#364f6b")) },
                        new ColorInfo { ColorName = "占位3", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#364f6b")) },
                        new ColorInfo { ColorName = "占位4", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#364f6b")) }
                    };
                case 6:
                    return new List<ColorInfo>
                    {
                        new ColorInfo { ColorName = "淡黄1", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#ffffd2")) },
                        new ColorInfo { ColorName = "淡黄2", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#ffffd2")) },
                        new ColorInfo { ColorName = "粉红1", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#fcbad3")) },
                        new ColorInfo { ColorName = "粉红2", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#fcbad3")) },
                        new ColorInfo { ColorName = "淡紫1", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#aa96da")) },
                        new ColorInfo { ColorName = "淡紫2", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#aa96da")) },
                        new ColorInfo { ColorName = "淡蓝1", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#a8d8ea")) },
                        new ColorInfo { ColorName = "淡蓝2", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#a8d8ea")) },
                        new ColorInfo { ColorName = "占位1", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#a8d8ea")) },
                        new ColorInfo { ColorName = "占位2", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#a8d8ea")) },
                        new ColorInfo { ColorName = "占位3", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#a8d8ea")) },
                        new ColorInfo { ColorName = "占位4", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#a8d8ea")) }
                    };
                case 7:
                    return new List<ColorInfo>
                    {
                        new ColorInfo { ColorName = "青绿1", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#71c9ce")) },
                        new ColorInfo { ColorName = "青绿2", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#71c9ce")) },
                        new ColorInfo { ColorName = "浅青1", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#a6e3e9")) },
                        new ColorInfo { ColorName = "浅青2", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#a6e3e9")) },
                        new ColorInfo { ColorName = "淡青1", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#cbf1f5")) },
                        new ColorInfo { ColorName = "淡青2", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#cbf1f5")) },
                        new ColorInfo { ColorName = "极淡青1", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#e3fdfd")) },
                        new ColorInfo { ColorName = "极淡青2", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#e3fdfd")) },
                        new ColorInfo { ColorName = "占位1", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#e3fdfd")) },
                        new ColorInfo { ColorName = "占位2", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#e3fdfd")) },
                        new ColorInfo { ColorName = "占位3", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#e3fdfd")) },
                        new ColorInfo { ColorName = "占位4", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#e3fdfd")) }
                    };
                case 8:
                    return new List<ColorInfo>
                    {
                        new ColorInfo { ColorName = "深灰1", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#40514e")) },
                        new ColorInfo { ColorName = "深灰2", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#40514e")) },
                        new ColorInfo { ColorName = "青绿1", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#11999e")) },
                        new ColorInfo { ColorName = "青绿2", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#11999e")) },
                        new ColorInfo { ColorName = "亮青1", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#30e3ca")) },
                        new ColorInfo { ColorName = "亮青2", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#30e3ca")) },
                        new ColorInfo { ColorName = "淡青1", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#e4f9f5")) },
                        new ColorInfo { ColorName = "淡青2", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#e4f9f5")) },
                        new ColorInfo { ColorName = "占位1", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#e4f9f5")) },
                        new ColorInfo { ColorName = "占位2", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#e4f9f5")) },
                        new ColorInfo { ColorName = "占位3", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#e4f9f5")) },
                        new ColorInfo { ColorName = "占位4", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#e4f9f5")) }
                    };
                case 9:
                    return new List<ColorInfo>
                    {
                        new ColorInfo { ColorName = "灰紫1", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#8785a2")) },
                        new ColorInfo { ColorName = "灰紫2", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#8785a2")) },
                        new ColorInfo { ColorName = "浅灰1", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#f6f6f6")) },
                        new ColorInfo { ColorName = "浅灰2", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#f6f6f6")) },
                        new ColorInfo { ColorName = "淡粉1", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#ffe2e2")) },
                        new ColorInfo { ColorName = "淡粉2", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#ffe2e2")) },
                        new ColorInfo { ColorName = "浅粉1", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#ffc7c7")) },
                        new ColorInfo { ColorName = "浅粉2", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#ffc7c7")) },
                        new ColorInfo { ColorName = "占位1", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#ffc7c7")) },
                        new ColorInfo { ColorName = "占位2", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#ffc7c7")) },
                        new ColorInfo { ColorName = "占位3", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#ffc7c7")) },
                        new ColorInfo { ColorName = "占位4", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#ffc7c7")) }
                    };
                case 10:
                    return new List<ColorInfo>
                    {
                        new ColorInfo { ColorName = "深蓝1", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#112d4e")) },
                        new ColorInfo { ColorName = "深蓝2", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#112d4e")) },
                        new ColorInfo { ColorName = "蓝色1", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#3f72af")) },
                        new ColorInfo { ColorName = "蓝色2", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#3f72af")) },
                        new ColorInfo { ColorName = "淡蓝1", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#dbe2ef")) },
                        new ColorInfo { ColorName = "淡蓝2", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#dbe2ef")) },
                        new ColorInfo { ColorName = "浅灰1", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#f9f7f7")) },
                        new ColorInfo { ColorName = "浅灰2", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#f9f7f7")) },
                        new ColorInfo { ColorName = "占位1", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#f9f7f7")) },
                        new ColorInfo { ColorName = "占位2", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#f9f7f7")) },
                        new ColorInfo { ColorName = "占位3", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#f9f7f7")) },
                        new ColorInfo { ColorName = "占位4", ColorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#f9f7f7")) }
                    };
                // ... 继续添加其他配色方案 ...
                default:
                    return null;
            }
        }
    }

    public class ColorScheme
    {
        public string Name { get; set; }
        public List<SolidColorBrush> Colors { get; set; }
    }
}
