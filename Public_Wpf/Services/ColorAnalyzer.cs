using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using ColorThiefDotNet;
using PresPio.Public_Wpf.Models;

namespace PresPio.Public_Wpf.Services
{
    public class ColorAnalyzer
    {
        private static readonly ColorThief colorThief = new ColorThief();
        private const int DefaultColorCount = 12;
        private const int DefaultQuality = 5;

        public static List<Models.ColorInfo> AnalyzeImage(string imagePath)
        {
            try
            {
                using (var bitmap = new Bitmap(imagePath))
                {
                    // 获取主色调
                    var dominantColor = colorThief.GetColor(bitmap, DefaultQuality);
                    
                    // 获取调色板
                    var palette = colorThief.GetPalette(bitmap, DefaultColorCount, DefaultQuality);

                    var colorInfos = new List<Models.ColorInfo>();

                    // 添加主色调
                    if (dominantColor != null)
                    {
                        var hsl = RgbToHsl(dominantColor.Color.R, dominantColor.Color.G, dominantColor.Color.B);
                        var colorInfo = new Models.ColorInfo
                        {
                            ColorHex = $"#{dominantColor.Color.R:X2}{dominantColor.Color.G:X2}{dominantColor.Color.B:X2}",
                            Hsl = hsl,
                            Percentage = 30.0
                        };
                        colorInfos.Add(colorInfo);
                    }

                    // 添加调色板颜色
                    if (palette != null)
                    {
                        double remainingPercentage = 70.0;
                        var validColors = palette.ToList();
                        int validColorCount = validColors.Count;
                        
                        if (validColorCount > 0)
                        {
                            // 根据颜色的Population调整权重
                            double totalPopulation = validColors.Sum(c => c.Population);
                            
                            foreach (var paletteColor in validColors)
                            {
                                var color = paletteColor.Color;
                                // 检查颜色是否与主色调过于相似
                                bool isSimilarToDominant = dominantColor != null && 
                                    IsSimilarRgb(color, dominantColor.Color);

                                if (!isSimilarToDominant)
                                {
                                    var hsl = RgbToHsl(color.R, color.G, color.B);
                                    
                                    // 计算该颜色的权重
                                    double colorPercentage = (paletteColor.Population / totalPopulation) * remainingPercentage;
                                    
                                    var colorInfo = new Models.ColorInfo
                                    {
                                        ColorHex = $"#{color.R:X2}{color.G:X2}{color.B:X2}",
                                        Hsl = hsl,
                                        Percentage = colorPercentage
                                    };
                                    colorInfos.Add(colorInfo);
                                }
                            }
                        }
                    }

                    // 规范化百分比
                    NormalizePercentages(colorInfos);
                    
                    // 过滤掉占比过小的颜色
                    var significantColors = colorInfos.Where(c => c.Percentage >= 5.0).ToList();
                    
                    // 如果过滤后没有颜色，返回原始列表
                    return significantColors.Any() ? significantColors : colorInfos;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"颜色分析错误: {ex.Message}");
                return new List<Models.ColorInfo>();
            }
        }

        private static void NormalizePercentages(List<Models.ColorInfo> colors)
        {
            if (colors == null || !colors.Any()) return;

            double total = colors.Sum(c => c.Percentage);
            if (total > 0)
            {
                foreach (var color in colors)
                {
                    color.Percentage = Math.Round((color.Percentage / total) * 100, 1);
                }
            }
        }

        private static bool IsSimilarRgb(ColorThiefDotNet.Color c1, ColorThiefDotNet.Color c2)
        {
            // 使用欧几里得距离计算颜色相似度
            double distance = Math.Sqrt(
                Math.Pow(c1.R - c2.R, 2) +
                Math.Pow(c1.G - c2.G, 2) +
                Math.Pow(c1.B - c2.B, 2)
            );

            // 根据人眼对不同颜色的敏感度调整阈值
            return distance < 30;
        }

        public static (double H, double S, double L) RgbToHsl(int r, int g, int b)
        {
            double rd = r / 255.0;
            double gd = g / 255.0;
            double bd = b / 255.0;
            double max = Math.Max(rd, Math.Max(gd, bd));
            double min = Math.Min(rd, Math.Min(gd, bd));
            double h = 0, s, l = (max + min) / 2;

            if (max == min)
            {
                h = s = 0;
            }
            else
            {
                double d = max - min;
                s = l > 0.5 ? d / (2 - max - min) : d / (max + min);

                if (max == rd)
                    h = (gd - bd) / d + (gd < bd ? 6 : 0);
                else if (max == gd)
                    h = (bd - rd) / d + 2;
                else if (max == bd)
                    h = (rd - gd) / d + 4;

                h /= 6;
            }

            return (H: h * 360, S: s * 100, L: l * 100);
        }

        public static bool AreColorsSimilar(Models.ColorInfo c1, Models.ColorInfo c2)
        {
            // 对于灰度色（无色系）使用特殊的比较逻辑
            bool isC1Gray = IsGrayscale(c1);
            bool isC2Gray = IsGrayscale(c2);

            if (isC1Gray || isC2Gray)
            {
                // 如果两个都是灰度色，只比较亮度
                if (isC1Gray && isC2Gray)
                {
                    return Math.Abs(c1.Hsl.L - c2.Hsl.L) < 15;
                }
                // 如果只有一个是灰度色，需要考虑饱和度
                return Math.Abs(c1.Hsl.L - c2.Hsl.L) < 15 && 
                       Math.Min(c1.Hsl.S, c2.Hsl.S) < 15;
            }

            // 色相差异阈值（考虑360度循环）
            double hueDiff = Math.Min(
                Math.Abs(c1.Hsl.H - c2.Hsl.H),
                360 - Math.Abs(c1.Hsl.H - c2.Hsl.H)
            );

            // 根据饱和度和亮度动态调整色相容差
            double hueThreshold = GetDynamicHueThreshold(c1.Hsl.S, c1.Hsl.L, c2.Hsl.S, c2.Hsl.L);
            double satThreshold = GetDynamicSaturationThreshold(c1.Hsl.L, c2.Hsl.L);
            double lightThreshold = GetDynamicLightnessThreshold(c1.Hsl.S, c2.Hsl.S);

            return hueDiff <= hueThreshold &&
                   Math.Abs(c1.Hsl.S - c2.Hsl.S) <= satThreshold &&
                   Math.Abs(c1.Hsl.L - c2.Hsl.L) <= lightThreshold;
        }

        private static double GetDynamicHueThreshold(double s1, double l1, double s2, double l2)
        {
            // 当饱和度或亮度较低时，放宽色相阈值
            double avgSaturation = (s1 + s2) / 2;
            double avgLightness = (l1 + l2) / 2;
            
            if (avgSaturation < 20 || avgLightness < 20 || avgLightness > 80)
                return 60;
            if (avgSaturation < 40)
                return 45;
            return 30;
        }

        private static double GetDynamicSaturationThreshold(double l1, double l2)
        {
            // 当亮度接近极值时，放宽饱和度阈值
            double avgLightness = (l1 + l2) / 2;
            if (avgLightness < 20 || avgLightness > 80)
                return 45;
            return 35;
        }

        private static double GetDynamicLightnessThreshold(double s1, double s2)
        {
            // 当饱和度较低时，收紧亮度阈值
            double avgSaturation = (s1 + s2) / 2;
            if (avgSaturation < 20)
                return 15;
            return 25;
        }

        private static bool IsGrayscale(Models.ColorInfo color)
        {
            return color.Hsl.S < 10 || color.Hsl.L < 10 || color.Hsl.L > 90;
        }

        public static string GetClosestStandardColor(Models.ColorInfo color)
        {
            if (IsGrayscale(color))
            {
                if (color.Hsl.L < 20) return "黑色";
                if (color.Hsl.L > 80) return "白色";
                return "灰色";
            }

            var hue = color.Hsl.H;
            var sat = color.Hsl.S;
            var light = color.Hsl.L;

            if (sat < 20 || light < 10 || light > 90)
            {
                if (light < 20) return "黑色";
                if (light > 80) return "白色";
                return "灰色";
            }

            if (hue < 15 || hue >= 345) return "红色";
            if (hue < 45) return "橙色";
            if (hue < 75) return "黄色";
            if (hue < 165) return "绿色";
            if (hue < 195) return "青色";
            if (hue < 285) return "蓝色";
            if (hue < 345) return "紫色";

            return "红色";
        }
    }
} 