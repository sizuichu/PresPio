using System;

namespace PresPio.Models
{
    public class ColorInfo
    {
        public string ColorHex { get; set; }
        public (double h, double s, double l) Hsl { get; set; }
        public double Percentage { get; set; }
    }
} 