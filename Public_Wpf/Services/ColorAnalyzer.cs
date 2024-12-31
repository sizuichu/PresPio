using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.Linq;
using PresPio.Public_Wpf.Models;

namespace PresPio.Public_Wpf.Services
{
    public class ColorAnalyzer
    {
        public static List<Models.ColorInfo> AnalyzeImage(string imagePath, int colorCount = 4)
        {
            using (var bitmap = new Bitmap(imagePath))
            {
                // 缩小图片以提高性能
                var resized = ResizeImage(bitmap, 100, 100);
                var colors = new Dictionary<System.Drawing.Color, int>();

                // 分析每个像素
                for (int x = 0; x < resized.Width; x++)
                {
                    for (int y = 0; y < resized.Height; y++)
                    {
                        var pixel = resized.GetPixel(x, y);
                        var quantizedColor = QuantizeColor(pixel);
                        if (colors.ContainsKey(quantizedColor))
                        {
                            colors[quantizedColor]++;
                        }
                        else
                        {
                            colors[quantizedColor] = 1;
                        }
                    }
                }

                // 获取主要颜色
                var totalPixels = resized.Width * resized.Height;
                var dominantColors = colors.OrderByDescending(x => x.Value)
                    .Take(colorCount)
                    .Select(x => new Models.ColorInfo
                    {
                        ColorHex = $"#{x.Key.R:X2}{x.Key.G:X2}{x.Key.B:X2}",
                        Percentage = (double)x.Value / totalPixels
                    })
                    .ToList();

                resized.Dispose();
                return dominantColors;
            }
        }

        private static System.Drawing.Color QuantizeColor(System.Drawing.Color color)
        {
            // 将颜色量化为32个级别以减少颜色数量
            int r = (color.R / 32) * 32;
            int g = (color.G / 32) * 32;
            int b = (color.B / 32) * 32;
            return System.Drawing.Color.FromArgb(r, g, b);
        }

        private static Bitmap ResizeImage(Image image, int width, int height)
        {
            var destRect = new Rectangle(0, 0, width, height);
            var destImage = new Bitmap(width, height);

            destImage.SetResolution(image.HorizontalResolution, image.VerticalResolution);

            using (var graphics = Graphics.FromImage(destImage))
            {
                graphics.CompositingMode = System.Drawing.Drawing2D.CompositingMode.SourceCopy;
                graphics.CompositingQuality = System.Drawing.Drawing2D.CompositingQuality.HighQuality;
                graphics.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.HighQualityBicubic;
                graphics.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.HighQuality;
                graphics.PixelOffsetMode = System.Drawing.Drawing2D.PixelOffsetMode.HighQuality;

                using (var wrapMode = new ImageAttributes())
                {
                    wrapMode.SetWrapMode(System.Drawing.Drawing2D.WrapMode.TileFlipXY);
                    graphics.DrawImage(image, destRect, 0, 0, image.Width, image.Height, GraphicsUnit.Pixel, wrapMode);
                }
            }

            return destImage;
        }
    }
} 