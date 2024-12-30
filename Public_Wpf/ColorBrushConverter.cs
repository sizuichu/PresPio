using System;
using System.Windows.Media;
using System.Windows.Data;
using System.Globalization;

namespace PresPio.Public_Wpf
{
    public class ColorBrushConverter : IValueConverter
    {
        private static readonly SolidColorBrush DefaultBrush = new SolidColorBrush(Colors.Gray);

        private static readonly SolidColorBrush[] TypeBrushes = new[]
        {
            new SolidColorBrush(Color.FromRgb(0x42, 0xA5, 0xF5)), // 图片
            new SolidColorBrush(Color.FromRgb(0x66, 0xBB, 0x6A)), // 文本框
            new SolidColorBrush(Color.FromRgb(0xFF, 0xA7, 0x26)), // 自选图形
            new SolidColorBrush(Color.FromRgb(0xAB, 0x47, 0xBC)), // 组合
            new SolidColorBrush(Color.FromRgb(0xEC, 0x40, 0x7A)), // 表格
            new SolidColorBrush(Color.FromRgb(0x78, 0x90, 0x9C)), // OLE对象
            new SolidColorBrush(Color.FromRgb(0x26, 0xA6, 0x9A)), // 图表
            new SolidColorBrush(Color.FromRgb(0x5C, 0x6B, 0xC0)), // SmartArt
            new SolidColorBrush(Color.FromRgb(0xEF, 0x53, 0x50)), // 媒体
            new SolidColorBrush(Color.FromRgb(0x8D, 0x6E, 0x63)), // 艺术字
            new SolidColorBrush(Color.FromRgb(0x26, 0xC6, 0xDA)), // 任意多边形
            new SolidColorBrush(Color.FromRgb(0x9C, 0xCC, 0x65)), // 直线
            new SolidColorBrush(Color.FromRgb(0x7E, 0x57, 0xC2))  // 占位符
        };

        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value is string shapeType)
            {
                return shapeType switch
                {
                    "图片" => TypeBrushes[0],
                    "文本框" => TypeBrushes[1],
                    "自选图形" => TypeBrushes[2],
                    "组合" => TypeBrushes[3],
                    "表格" => TypeBrushes[4],
                    "OLE对象" => TypeBrushes[5],
                    "图表" => TypeBrushes[6],
                    "SmartArt" => TypeBrushes[7],
                    "媒体" => TypeBrushes[8],
                    "艺术字" => TypeBrushes[9],
                    "任意多边形" => TypeBrushes[10],
                    "直线" => TypeBrushes[11],
                    "占位符" => TypeBrushes[12],
                    _ => DefaultBrush
                };
            }

            return DefaultBrush;
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }
} 