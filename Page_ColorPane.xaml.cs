using System;
using System.Collections.Generic;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using HandyControl.Controls;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
namespace PresPio
    {
    public partial class Page_ColorPane : UserControl
        {
          public PowerPoint.Application app; //加载PPT项目
        public Page_ColorPane()
            {
            InitializeComponent();


            List<ColorClass> colors = new List<ColorClass>
{
    new ColorClass("Red", "#F44336", 0, 1.00f, 0.50f),
    new ColorClass("Pink", "#E91E63", 330, 1.00f, 0.50f),
    new ColorClass("Purple", "#9C27B0", 270, 1.00f, 0.50f),
    new ColorClass("Deep Purple", "#673AB7", 240, 1.00f, 0.50f),
    new ColorClass("Indigo", "#3F51B5", 240, 1.00f, 0.50f),
    new ColorClass("Blue", "#2196F3", 210, 1.00f, 0.50f),
    new ColorClass("Light Blue", "#03A9F4", 195, 1.00f, 0.50f),
    new ColorClass("Cyan", "#00BCD4", 180, 1.00f, 0.50f),
    new ColorClass("Teal", "#009688", 160, 1.00f, 0.50f),
    new ColorClass("Green", "#4CAF50", 150, 1.00f, 0.50f),
    new ColorClass("Light Green", "#8BC34A", 120, 1.00f, 0.50f),
    new ColorClass("Lime", "#CDDC39", 60, 1.00f, 0.50f),
    new ColorClass("Yellow", "#FFEB3B", 60, 1.00f, 0.50f),
    new ColorClass("Amber", "#FFC107", 45, 1.00f, 0.50f),
    new ColorClass("Orange", "#FF9800", 30, 1.00f, 0.50f),
    new ColorClass("Deep Orange", "#FF5722", 15, 1.00f, 0.50f),
    new ColorClass("Brown", "#795548", 30, 0.50f, 0.30f),
    new ColorClass("Grey", "#9E9E9E", 0, 0.00f, 0.60f),
    new ColorClass("Blue Grey", "#607D8B", 195, 0.30f, 0.40f),
    new ColorClass("Black", "#000000", 0, 0.00f, 0.00f),
    new ColorClass("White", "#FFFFFF", 0, 0.00f, 1.00f),
    new ColorClass("Light Grey", "#BDBDBD", 0, 0.00f, 0.75f),
    new ColorClass("Dark Grey", "#616161", 0, 0.00f, 0.38f),
    new ColorClass("Light Pink", "#F8BBD0", 340, 0.70f, 0.80f),
    new ColorClass("Dark Pink", "#C2185B", 330, 0.80f, 0.45f),
    new ColorClass("Violet", "#8E24AA", 270, 0.80f, 0.45f),
    new ColorClass("Lavender", "#E1BEE7", 270, 0.50f, 0.90f),
    new ColorClass("Lilac", "#CE93D8", 270, 0.60f, 0.75f),
    new ColorClass("Magenta", "#D500F9", 300, 1.00f, 0.50f),
    new ColorClass("Deep Pink", "#E91E63", 330, 1.00f, 0.50f),
    new ColorClass("Turquoise", "#1DE9B6", 170, 1.00f, 0.50f),
    new ColorClass("Emerald", "#2E7D32", 150, 0.75f, 0.45f),
    new ColorClass("Forest Green", "#388E3C", 120, 0.75f, 0.30f),
    new ColorClass("Olive", "#8D6E63", 30, 0.50f, 0.30f),
    new ColorClass("Peach", "#FFAB91", 6, 0.80f, 0.70f),
    new ColorClass("Coral", "#FF7043", 14, 1.00f, 0.60f),
    new ColorClass("Gold", "#FFD54F", 51, 1.00f, 0.60f),
    new ColorClass("Sunset", "#FF7043", 20, 1.00f, 0.60f),
    new ColorClass("Peach Puff", "#FFCCBC", 20, 0.85f, 0.90f),
    new ColorClass("Blush", "#F48FB1", 330, 0.60f, 0.80f),
    new ColorClass("Salmon", "#FF8A65", 15, 1.00f, 0.50f),
    new ColorClass("Crimson", "#D32F2F", 0, 0.70f, 0.50f),
    new ColorClass("Slate Blue", "#5E35B1", 240, 0.60f, 0.60f),
    new ColorClass("Electric Blue", "#00B0FF", 200, 1.00f, 0.50f),
    new ColorClass("Aqua", "#00FFFF", 180, 1.00f, 0.50f),
    new ColorClass("Spring Green", "#00FF7F", 150, 1.00f, 0.50f),
    new ColorClass("Sea Green", "#4CAF50", 145, 1.00f, 0.45f),
    new ColorClass("Lime Green", "#32CD32", 120, 1.00f, 0.50f),
    new ColorClass("Cobalt Blue", "#3D5AFE", 220, 0.85f, 0.55f),
    new ColorClass("Electric Purple", "#D500F9", 300, 1.00f, 0.50f),
    new ColorClass("Grape", "#6A1B9A", 270, 0.75f, 0.50f),
    new ColorClass("Magenta", "#FF00FF", 300, 1.00f, 0.50f),
    new ColorClass("Sky Blue", "#00B0FF", 200, 1.00f, 0.50f),
    new ColorClass("Moss Green", "#558B2F", 80, 0.50f, 0.30f),
    new ColorClass("Sapphire", "#1A237E", 210, 0.70f, 0.30f),
    new ColorClass("Ruby", "#D50000", 0, 1.00f, 0.45f),
    new ColorClass("Topaz", "#FFD600", 51, 1.00f, 0.50f),
    new ColorClass("Charcoal", "#424242", 0, 0.00f, 0.30f),
    new ColorClass("Jet Black", "#343434", 0, 0.00f, 0.20f),
    new ColorClass("Cotton Candy", "#FFB3E3", 340, 0.80f, 0.90f),
    new ColorClass("Honey", "#FFEB3B", 54, 1.00f, 0.60f),
    new ColorClass("Emerald Green", "#2E7D32", 150, 0.75f, 0.45f),
    new ColorClass("Teal Blue", "#00796B", 180, 0.60f, 0.40f)
};






            // 为每个颜色生成子颜色
            foreach (var color in colors)
                {
                color.GenerateSubColors();
                }

            // 将列表绑定到 DataContext
            this.DataContext = new ColorViewModel { Colors = colors, SubColors = new List<ColorClass>() };
            }

        // 处理点击主颜色的事件
        private void OnColorClick(object sender, MouseButtonEventArgs e)
            {
            ColorClass color;
            if (sender is Border border)
                {
                color = border.DataContext as ColorClass;
                }
            else
                {
                color = sender as ColorClass;
                }

            if (color != null)
                {
                // 生成子颜色列表
                var subColors = color.GenerateSubColors();

                // 更新绑定的子颜色列表，触发 UI 更新
                var currentColors = (this.DataContext as ColorViewModel)?.Colors;

                // 更新 DataContext
                this.DataContext = new ColorViewModel { Colors = currentColors, SubColors = subColors };
                }
            }


        // 颜色类
        public class ColorClass
            {
            public string ColorName { get; set; }
            public string HexValue { get; set; }
            public float Hue { get; set; }
            public float Saturation { get; set; }
            public float Lightness { get; set; }
            public int SubColorCount { get; set; }
            public List<ColorClass> SubColors { get; set; }

            // 子颜色的背景颜色（用于 UI 绑定）
            public SolidColorBrush BackgroundColor
                {
                get
                    {
                    var color = HSLToRGB(Hue, Saturation, Lightness);
                    return new SolidColorBrush(color);
                    }
                }

            // 构造函数
            public ColorClass(string colorName, string hexValue, float hue, float saturation, float lightness, int subColorCount = 5)
                {
                ColorName = colorName;
                HexValue = hexValue;
                Hue = hue;
                Saturation = saturation;
                Lightness = lightness;
                SubColorCount = subColorCount;
                SubColors = new List<ColorClass>();
                }

            // 将 HSL 转换为 RGB
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

            // 生成子颜色列表
            public List<ColorClass> GenerateSubColors()
                {
                var subColors = new List<ColorClass>();

                // 生成 12 个子颜色（从深到浅）
                for (int i = 0 ; i < 12 ; i++)
                    {
                    // 保持色调在同一范围，避免跳跃
                    float newHue = Hue; // 保持主色调不变

                    // 渐变变化饱和度，逐步增强饱和度
                    float saturationChange = 0.2f * (i + 1); // 逐渐增加饱和度
                    float newSaturation = Saturation + saturationChange; // 逐步增加饱和度
                    newSaturation = Clamp(newSaturation, 0.2f, 1.0f); // 使用自定义的 Clamp 方法，确保饱和度在合理范围

                    // 使亮度逐渐增加，从深色过渡到浅色
                    float newLightness = Lightness + (i * 0.05f); // 增加亮度，使颜色从深到浅
                    newLightness = Clamp(newLightness, 0.3f, 0.8f); // 确保亮度在合理范围

                    // 创建新的子颜色
                    var subColor = new ColorClass($"{ColorName} - Sub {i + 1}", HexValue, newHue, newSaturation, newLightness);
                    subColors.Add(subColor); // 添加到子颜色列表
                    }

                return subColors; // 返回子颜色列表
                }

            public static float Clamp(float value, float min, float max)
                {
                if (value < min)
                    return min;
                if (value > max)
                    return max;
                return value;
                }


            }


        // 用于传递颜色数据和子颜色数据的 ViewModel
        public class ColorViewModel
            {
            public List<ColorClass> Colors { get; set; }
            public List<ColorClass> SubColors { get; set; }
            }

        private void OnSubColorClick(object sender, MouseButtonEventArgs e)
            {
            // 获取点击的 Border 控件
            var border = sender as Border;
            if (border != null && border.Background is SolidColorBrush brush)
                {
                try
                    {
                    // 获取颜色值并转换
                    var color = brush.Color;
                    int colorRgb = (color.R) | (color.G << 8) | (color.B << 16);

                    // 获取当前PPT中选中的内容
                    var selection = Globals.ThisAddIn.Application.ActiveWindow.Selection;
                    bool isCtrlPressed = (System.Windows.Forms.Control.ModifierKeys & System.Windows.Forms.Keys.Control) == System.Windows.Forms.Keys.Control;

                    if (selection.Type == PowerPoint.PpSelectionType.ppSelectionShapes)
                        {
                        // 遍历所有选中的形状
                        foreach (PowerPoint.Shape shape in selection.ShapeRange)
                            {
                            if (isCtrlPressed)
                                {
                                // 设置形状边框颜色
                                shape.Line.ForeColor.RGB = colorRgb;
                                shape.Line.Visible = Microsoft.Office.Core.MsoTriState.msoTrue;
                                }
                            else
                                {
                                // 设置形状填充色
                                shape.Fill.ForeColor.RGB = colorRgb;
                                shape.Fill.Visible = Microsoft.Office.Core.MsoTriState.msoTrue;
                                }
                            }
                        //Growl.SuccessGlobal($"已成功应用颜色 #{color.R:X2}{color.G:X2}{color.B:X2} 到选中{(isCtrlPressed ? "形状边框" : "形状")}");
                        }
                    else if (selection.Type == PowerPoint.PpSelectionType.ppSelectionText)
                        {
                        // 处理文本选择
                        var textRange = selection.TextRange;
                        if (isCtrlPressed)
                            {
                            }
                        else
                            {
                            // 设置文字颜色
                            textRange.Font.Color.RGB = colorRgb;
                            }
                      //  Growl.SuccessGlobal($"已成功应用颜色 #{color.R:X2}{color.G:X2}{color.B:X2} 到选中{(isCtrlPressed ? "文字边框" : "文字")}");
                        }
                    else
                        {
                        Growl.WarningGlobal("请先在PPT中选择要填充颜色的形状或文字");
                        }
                    }
                catch (Exception ex)
                    {
                    Growl.ErrorGlobal($"应用颜色时出错: {ex.Message}");
                    }
                }
            }

      

        private void Border_MouseRightButtonDown(object sender, MouseButtonEventArgs e)
            {
            // 获取点击的 Border 控件
            var border = sender as Border;
            if (border != null)
                {
                // 获取 Border 背景色的 Hex 值
                string colorHex = (border.Background as SolidColorBrush)?.Color.ToString();

                // 将颜色值复制到剪贴板
                if (!string.IsNullOrEmpty(colorHex))
                    {
                    Clipboard.SetText(colorHex);
                    Growl.SuccessGlobal($"颜色值 {colorHex} 已复制到剪贴板");
                    }
                }
            }
        }
    }
