using System;
using System.Globalization;
using System.Windows.Data;
using System.Windows.Media;

namespace PresPio.Public_Wpf
{
    public class ShapeTypeColorConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value is string shapeType)
            {
                return new SolidColorBrush(shapeType switch
                {
                    "图片" => Color.FromRgb(0x42, 0xA5, 0xF5),      // 蓝色
                    "文本框" => Color.FromRgb(0x66, 0xBB, 0x6A),    // 绿色
                    "自选图形" => Color.FromRgb(0xFF, 0xA7, 0x26),  // 橙色
                    "组合" => Color.FromRgb(0xAB, 0x47, 0xBC),      // 紫色
                    "表格" => Color.FromRgb(0xEC, 0x40, 0x7A),      // 粉色
                    "OLE对象" => Color.FromRgb(0x78, 0x90, 0x9C),   // 蓝灰色
                    "图表" => Color.FromRgb(0x26, 0xA6, 0x9A),      // 青色
                    "SmartArt" => Color.FromRgb(0x5C, 0x6B, 0xC0), // 靛蓝色
                    "媒体" => Color.FromRgb(0xEF, 0x53, 0x50),      // 红色
                    "艺术字" => Color.FromRgb(0x8D, 0x6E, 0x63),    // 棕色
                    "任意多边形" => Color.FromRgb(0x26, 0xC6, 0xDA), // 浅蓝色
                    "直线" => Color.FromRgb(0x9C, 0xCC, 0x65),      // 黄绿色
                    "占位符" => Color.FromRgb(0x7E, 0x57, 0xC2),    // 深紫色
                    _ => Color.FromRgb(0x9E, 0x9E, 0x9E)           // 灰色
                });
            }

            return new SolidColorBrush(Colors.Gray);
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }
} 