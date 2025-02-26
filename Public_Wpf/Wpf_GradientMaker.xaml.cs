using HandyControl.Controls;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using System;
using System.Windows;
using System.Windows.Media;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace PresPio
{
    public partial class Wpf_GradientMaker : Window
    {
        private PowerPoint.Application app;
        private Random random = new Random();

        public Wpf_GradientMaker()
        {
            InitializeComponent();
            InitializeGradient();
            SetupEventHandlers();
        }

        private void InitializeGradient()
        {
            UpdateGradient();
        }

        private void SetupEventHandlers()
        {
            // 参数变化事件
            StepInput.ValueChanged += (s, e) => UpdateGradient();
            ColorDiffInput.ValueChanged += (s, e) => UpdateGradient();
            AngleInput.ValueChanged += (s, e) => UpdateGradient();

            // 渐变类型切换
            LinearGradientRadio.Checked += (s, e) => UpdateGradient();
            RadialGradientRadio.Checked += (s, e) => UpdateGradient();
        }

        private void UpdateGradient()
        {
            if (PreviewRect == null) return;

            // 生成随机颜色
            Color startColor = GenerateRandomColor();
            Color endColor = GenerateColorWithDifference(startColor, (int)ColorDiffInput.Value);

            if (LinearGradientRadio.IsChecked == true)
            {
                var brush = new LinearGradientBrush();
                double angle = AngleInput.Value * Math.PI / 180;
                Point start = new Point(0.5 - Math.Cos(angle) / 2, 0.5 - Math.Sin(angle) / 2);
                Point end = new Point(0.5 + Math.Cos(angle) / 2, 0.5 + Math.Sin(angle) / 2);
                
                brush.StartPoint = start;
                brush.EndPoint = end;
                brush.GradientStops.Add(new GradientStop(startColor, 0));
                brush.GradientStops.Add(new GradientStop(endColor, 1));

                PreviewRect.Fill = brush;
            }
            else
            {
                var brush = new RadialGradientBrush();
                brush.Center = new Point(0.5, 0.5);
                brush.GradientOrigin = new Point(0.5, 0.5);
                brush.GradientStops.Add(new GradientStop(startColor, 0));
                brush.GradientStops.Add(new GradientStop(endColor, 1));

                PreviewRect.Fill = brush;
            }
        }

        private Color GenerateRandomColor()
        {
            return Color.FromRgb(
                (byte)random.Next(256),
                (byte)random.Next(256),
                (byte)random.Next(256)
            );
        }

        private Color GenerateColorWithDifference(Color baseColor, int difference)
        {
            int r = Math.Min(255, Math.Max(0, baseColor.R + random.Next(-difference, difference)));
            int g = Math.Min(255, Math.Max(0, baseColor.G + random.Next(-difference, difference)));
            int b = Math.Min(255, Math.Max(0, baseColor.B + random.Next(-difference, difference)));

            return Color.FromRgb((byte)r, (byte)g, (byte)b);
        }

        #region 按钮事件处理

        private void CopyBtn_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                app = Globals.ThisAddIn.Application;
                Selection sel = app.ActiveWindow.Selection;
                if (sel.Type == PpSelectionType.ppSelectionShapes)
                {
                    sel.ShapeRange.PickUp();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"复制时出错：{ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void PasteBtn_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                app = Globals.ThisAddIn.Application;
                Selection sel = app.ActiveWindow.Selection;
                if (sel.Type == PpSelectionType.ppSelectionShapes)
                {
                    sel.ShapeRange.Apply();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"粘贴时出错：{ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void ClearBtn_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                app = Globals.ThisAddIn.Application;
                Selection sel = app.ActiveWindow.Selection;
                if (sel.Type == PpSelectionType.ppSelectionShapes)
                {
                    foreach (Shape shape in sel.ShapeRange)
                    {
                        shape.Fill.Solid();
                        shape.Fill.ForeColor.RGB = 16777215; // 白色
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"清除时出错：{ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void GenerateBtn_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                app = Globals.ThisAddIn.Application;
                Selection sel = app.ActiveWindow.Selection;
                if (sel.Type == PpSelectionType.ppSelectionShapes)
                {
                    Shape shape = sel.ShapeRange[1];
                    LinearGradientBrush brush = PreviewRect.Fill as LinearGradientBrush;
                    if (brush != null && brush.GradientStops.Count >= 2)
                    {
                        shape.Fill.TwoColorGradient(MsoGradientStyle.msoGradientHorizontal, 1);
                        shape.Fill.GradientAngle = (float)AngleInput.Value;
                        shape.Fill.ForeColor.RGB = ColorToRGB(brush.GradientStops[0].Color);
                        shape.Fill.BackColor.RGB = ColorToRGB(brush.GradientStops[1].Color);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"生成渐变时出错：{ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void PresetList_SelectionChanged(object sender, System.Windows.Controls.SelectionChangedEventArgs e)
        {
            if (PresetList.SelectedItem == null) return;

            var selectedPreset = (System.Windows.Controls.ListBoxItem)PresetList.SelectedItem;
            switch (selectedPreset.Content.ToString())
            {
                case "蓝色渐变":
                    ApplyPreset(Color.FromRgb(75, 108, 183), Color.FromRgb(24, 40, 72));
                    break;
                case "绿色渐变":
                    ApplyPreset(Color.FromRgb(76, 209, 55), Color.FromRgb(35, 98, 26));
                    break;
                case "红色渐变":
                    ApplyPreset(Color.FromRgb(235, 87, 87), Color.FromRgb(150, 24, 24));
                    break;
                case "紫色渐变":
                    ApplyPreset(Color.FromRgb(187, 87, 235), Color.FromRgb(106, 24, 150));
                    break;
                case "橙色渐变":
                    ApplyPreset(Color.FromRgb(235, 151, 87), Color.FromRgb(150, 79, 24));
                    break;
            }
        }

        private void ApplyPreset(Color startColor, Color endColor)
        {
            var brush = new LinearGradientBrush();
            double angle = AngleInput.Value * Math.PI / 180;
            Point start = new Point(0.5 - Math.Cos(angle) / 2, 0.5 - Math.Sin(angle) / 2);
            Point end = new Point(0.5 + Math.Cos(angle) / 2, 0.5 + Math.Sin(angle) / 2);
            
            brush.StartPoint = start;
            brush.EndPoint = end;
            brush.GradientStops.Add(new GradientStop(startColor, 0));
            brush.GradientStops.Add(new GradientStop(endColor, 1));

            PreviewRect.Fill = brush;
        }

        #endregion

        private int ColorToRGB(Color color)
        {
            return color.R + (color.G << 8) + (color.B << 16);
        }
    }
} 