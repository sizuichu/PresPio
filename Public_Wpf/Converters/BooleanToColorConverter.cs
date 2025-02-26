using System;
using System.Globalization;
using System.Windows.Data;
using System.Windows.Media;

namespace PresPio.Public_Wpf.Converters
{
    public class BooleanToColorConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value is bool isSelected && parameter is Brush selectedBrush)
            {
                return isSelected ? selectedBrush : new SolidColorBrush(Colors.Gray);
            }
            return new SolidColorBrush(Colors.Gray);
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }
} 