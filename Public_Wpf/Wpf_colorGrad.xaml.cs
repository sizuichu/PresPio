using HandyControl.Controls;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using System;
using System.Collections.ObjectModel;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Shapes;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Xml;
using Newtonsoft.Json;
using Point = System.Windows.Point;
using Shape = Microsoft.Office.Interop.PowerPoint.Shape;
using MessageBox = HandyControl.Controls.MessageBox;
using Window = HandyControl.Controls.Window;
using TextBox = HandyControl.Controls.TextBox;
using GradientStop = System.Windows.Media.GradientStop;

namespace PresPio
{
    public class ColorStop
    {
        public Color Color { get; set; }
        public double Position { get; set; }
    }

    public class GradientPreset
    {
        public string Name { get; set; }
        public List<ColorStop> ColorStops { get; set; }
        public string Type { get; set; }
        public double Angle { get; set; }
        public double CenterX { get; set; }
        public double CenterY { get; set; }
        public int RepeatCount { get; set; }
    }

    public partial class Wpf_GradientMaker : Window
    {
        private PowerPoint.Application app;
        private ObservableCollection<ColorStop> colorStops;
        private Point? dragStartPoint;
        private UIElement draggedElement;
        private bool isDragging;
        private const string CUSTOM_PRESETS_FILE = "custom_presets.json";

        public Wpf_GradientMaker()
        {
            InitializeComponent();
            InitializeGradient();
            SetupEventHandlers();
            LoadCustomPresets();
        }

        private void InitializeGradient()
        {
            colorStops = new ObservableCollection<ColorStop>
            {
                new ColorStop { Color = Color.FromRgb(75, 108, 183), Position = 0 },
                new ColorStop { Color = Color.FromRgb(24, 40, 72), Position = 100 }
            };
            ColorStopsList.ItemsSource = colorStops;
            UpdateGradient();
            UpdateControlPoints();
        }

        private void SetupEventHandlers()
        {
            // 参数变化事件
            AngleInput.ValueChanged += (s, e) => UpdateGradient();
            CenterXInput.ValueChanged += (s, e) => UpdateGradient();
            CenterYInput.ValueChanged += (s, e) => UpdateGradient();
            RepeatCountInput.ValueChanged += (s, e) => UpdateGradient();

            // 渐变类型切换
            LinearGradientRadio.Checked += (s, e) => OnGradientTypeChanged();
            RadialGradientRadio.Checked += (s, e) => OnGradientTypeChanged();
            ConicGradientRadio.Checked += (s, e) => OnGradientTypeChanged();
            RepeatingGradientRadio.Checked += (s, e) => OnGradientTypeChanged();
        }

        private void OnGradientTypeChanged()
        {
            // 显示/隐藏相关控件
            bool isRadial = RadialGradientRadio.IsChecked == true;
            bool isRepeating = RepeatingGradientRadio.IsChecked == true;
            
            CenterXPanel.Visibility = isRadial ? Visibility.Visible : Visibility.Collapsed;
            CenterYPanel.Visibility = isRadial ? Visibility.Visible : Visibility.Collapsed;
            RepeatCountPanel.Visibility = isRepeating ? Visibility.Visible : Visibility.Collapsed;
            
            UpdateGradient();
            UpdateControlPoints();
        }

        private void UpdateGradient()
        {
            if (PreviewRect == null || colorStops == null || colorStops.Count < 2) return;

            if (LinearGradientRadio.IsChecked == true)
            {
                UpdateLinearGradient();
            }
            else if (RadialGradientRadio.IsChecked == true)
            {
                UpdateRadialGradient();
            }
            else if (ConicGradientRadio.IsChecked == true)
            {
                UpdateConicGradient();
            }
            else if (RepeatingGradientRadio.IsChecked == true)
            {
                UpdateRepeatingGradient();
            }
        }

        private void UpdateLinearGradient()
        {
            var brush = new LinearGradientBrush();
            double angle = AngleInput.Value * Math.PI / 180;
            Point start = new Point(0.5 - Math.Cos(angle) / 2, 0.5 - Math.Sin(angle) / 2);
            Point end = new Point(0.5 + Math.Cos(angle) / 2, 0.5 + Math.Sin(angle) / 2);
            
            brush.StartPoint = start;
            brush.EndPoint = end;
            
            foreach (var stop in colorStops)
            {
                brush.GradientStops.Add(new GradientStop(stop.Color, stop.Position / 100));
            }

            PreviewRect.Fill = brush;
            UpdateControlPoints();
        }

        private void UpdateRadialGradient()
        {
            var brush = new RadialGradientBrush();
            brush.Center = new Point(CenterXInput.Value / 100, CenterYInput.Value / 100);
            brush.GradientOrigin = brush.Center;
            
            foreach (var stop in colorStops)
            {
                brush.GradientStops.Add(new GradientStop(stop.Color, stop.Position / 100));
            }

            PreviewRect.Fill = brush;
        }

        private void UpdateConicGradient()
        {
            // WPF不直接支持锥形渐变，这里使用径向渐变模拟
            var brush = new RadialGradientBrush();
            brush.Center = new Point(0.5, 0.5);
            brush.GradientOrigin = brush.Center;
            
            int steps = 360;
            double angle = AngleInput.Value;
            
            for (int i = 0; i <= steps; i++)
            {
                double position = i / (double)steps;
                double currentAngle = (position * 360 + angle) % 360;
                Color color = GetColorAtPosition(currentAngle / 360);
                brush.GradientStops.Add(new GradientStop(color, position));
            }

            PreviewRect.Fill = brush;
        }

        private void UpdateRepeatingGradient()
        {
            var brush = new LinearGradientBrush();
            double angle = AngleInput.Value * Math.PI / 180;
            Point start = new Point(0.5 - Math.Cos(angle) / 2, 0.5 - Math.Sin(angle) / 2);
            Point end = new Point(0.5 + Math.Cos(angle) / 2, 0.5 + Math.Sin(angle) / 2);
            
            brush.StartPoint = start;
            brush.EndPoint = end;
            brush.SpreadMethod = GradientSpreadMethod.Repeat;
            
            int repeatCount = (int)RepeatCountInput.Value;
            double segment = 1.0 / repeatCount;
            
            foreach (var stop in colorStops)
            {
                brush.GradientStops.Add(new GradientStop(stop.Color, (stop.Position / 100) * segment));
            }

            PreviewRect.Fill = brush;
            UpdateControlPoints();
        }

        private Color GetColorAtPosition(double position)
        {
            if (colorStops.Count < 2) return Colors.Black;

            for (int i = 0; i < colorStops.Count - 1; i++)
            {
                double currentPos = colorStops[i].Position / 100;
                double nextPos = colorStops[i + 1].Position / 100;
                
                if (position >= currentPos && position <= nextPos)
                {
                    double t = (position - currentPos) / (nextPos - currentPos);
                    return InterpolateColor(colorStops[i].Color, colorStops[i + 1].Color, t);
                }
            }

            return colorStops[colorStops.Count - 1].Color;
        }

        private Color InterpolateColor(Color c1, Color c2, double t)
        {
            return Color.FromArgb(
                (byte)(c1.A + (c2.A - c1.A) * t),
                (byte)(c1.R + (c2.R - c1.R) * t),
                (byte)(c1.G + (c2.G - c1.G) * t),
                (byte)(c1.B + (c2.B - c1.B) * t)
            );
        }

        private void UpdateControlPoints()
        {
            if (LinearGradientRadio.IsChecked == true || RepeatingGradientRadio.IsChecked == true)
            {
                double angle = AngleInput.Value * Math.PI / 180;
                double radius = Math.Min(ControlPointsCanvas.ActualWidth, ControlPointsCanvas.ActualHeight) / 2;
                
                Point center = new Point(ControlPointsCanvas.ActualWidth / 2, ControlPointsCanvas.ActualHeight / 2);
                Point start = new Point(
                    center.X - Math.Cos(angle) * radius,
                    center.Y - Math.Sin(angle) * radius
                );
                Point end = new Point(
                    center.X + Math.Cos(angle) * radius,
                    center.Y + Math.Sin(angle) * radius
                );

                Canvas.SetLeft(StartPoint, start.X - StartPoint.Width / 2);
                Canvas.SetTop(StartPoint, start.Y - StartPoint.Height / 2);
                Canvas.SetLeft(EndPoint, end.X - EndPoint.Width / 2);
                Canvas.SetTop(EndPoint, end.Y - EndPoint.Height / 2);

                GradientLine.X1 = start.X;
                GradientLine.Y1 = start.Y;
                GradientLine.X2 = end.X;
                GradientLine.Y2 = end.Y;

                ControlPointsCanvas.Visibility = Visibility.Visible;
            }
            else
            {
                ControlPointsCanvas.Visibility = Visibility.Collapsed;
            }
        }

        #region 拖拽控制点
        private void PreviewCanvas_MouseDown(object sender, MouseButtonEventArgs e)
        {
            dragStartPoint = e.GetPosition(ControlPointsCanvas);
            draggedElement = e.Source as UIElement;
            if (draggedElement != null && (draggedElement == StartPoint || draggedElement == EndPoint))
            {
                isDragging = true;
                draggedElement.CaptureMouse();
            }
        }

        private void PreviewCanvas_MouseMove(object sender, MouseEventArgs e)
        {
            if (isDragging && draggedElement != null && dragStartPoint.HasValue)
            {
                Point currentPoint = e.GetPosition(ControlPointsCanvas);
                double angle = Math.Atan2(
                    currentPoint.Y - ControlPointsCanvas.ActualHeight / 2,
                    currentPoint.X - ControlPointsCanvas.ActualWidth / 2
                );
                
                angle = angle * 180 / Math.PI;
                if (angle < 0) angle += 360;
                
                AngleInput.Value = angle;
            }
        }

        private void PreviewCanvas_MouseUp(object sender, MouseButtonEventArgs e)
        {
            if (isDragging && draggedElement != null)
            {
                isDragging = false;
                draggedElement.ReleaseMouseCapture();
                draggedElement = null;
                dragStartPoint = null;
            }
        }
        #endregion

        #region 颜色节点操作
        private void AddColorStop_Click(object sender, RoutedEventArgs e)
        {
            double position = 50;
            if (colorStops.Count >= 2)
            {
                position = (colorStops[0].Position + colorStops[1].Position) / 2;
            }

            Color color = GetColorAtPosition(position / 100);
            colorStops.Add(new ColorStop { Color = color, Position = position });
            UpdateGradient();
        }

        private void RemoveColorStop_Click(object sender, RoutedEventArgs e)
        {
            if (colorStops.Count <= 2) return;

            var button = sender as Button;
            var colorStop = button.DataContext as ColorStop;
            colorStops.Remove(colorStop);
            UpdateGradient();
        }
        #endregion

        #region 预设操作
        private void LoadCustomPresets()
        {
            try
            {
                if (File.Exists(CUSTOM_PRESETS_FILE))
                {
                    string json = File.ReadAllText(CUSTOM_PRESETS_FILE);
                    var presets = JsonConvert.DeserializeObject<List<GradientPreset>>(json);
                    foreach (var preset in presets)
                    {
                        CustomPresetList.Items.Add(new ListBoxItem { Content = preset.Name, Tag = preset });
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"加载预设时出错：{ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void SavePresetBtn_Click(object sender, RoutedEventArgs e)
        {
            var inputDialog = new HandyControl.Controls.InputDialog
            {
                Title = "保存预设",
                Content = "请输入预设名称：",
                DefaultValue = ""
            };

            if (inputDialog.ShowDialog().GetValueOrDefault())
            {
                string name = inputDialog.Text;
                if (string.IsNullOrWhiteSpace(name))
                {
                    MessageBox.Show("预设名称不能为空", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
                    return;
                }

                var preset = new GradientPreset
                {
                    Name = name,
                    ColorStops = new List<ColorStop>(colorStops),
                    Type = GetCurrentGradientType(),
                    Angle = AngleInput.Value,
                    CenterX = CenterXInput.Value,
                    CenterY = CenterYInput.Value,
                    RepeatCount = (int)RepeatCountInput.Value
                };

                SavePreset(preset);
                CustomPresetList.Items.Add(new ListBoxItem { Content = name, Tag = preset });
            }
        }

        private string GetCurrentGradientType()
        {
            if (LinearGradientRadio.IsChecked == true) return "Linear";
            if (RadialGradientRadio.IsChecked == true) return "Radial";
            if (ConicGradientRadio.IsChecked == true) return "Conic";
            if (RepeatingGradientRadio.IsChecked == true) return "Repeating";
            return "Linear";
        }

        private void SavePreset(GradientPreset preset)
        {
            try
            {
                List<GradientPreset> presets;
                if (File.Exists(CUSTOM_PRESETS_FILE))
                {
                    string json = File.ReadAllText(CUSTOM_PRESETS_FILE);
                    presets = JsonConvert.DeserializeObject<List<GradientPreset>>(json);
                }
                else
                {
                    presets = new List<GradientPreset>();
                }

                presets.Add(preset);
                string newJson = JsonConvert.SerializeObject(presets, Formatting.Indented);
                File.WriteAllText(CUSTOM_PRESETS_FILE, newJson);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"保存预设时出错：{ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void DeletePresetBtn_Click(object sender, RoutedEventArgs e)
        {
            var selectedItem = CustomPresetList.SelectedItem as ListBoxItem;
            if (selectedItem == null) return;

            try
            {
                string json = File.ReadAllText(CUSTOM_PRESETS_FILE);
                var presets = JsonConvert.DeserializeObject<List<GradientPreset>>(json);
                presets.RemoveAll(p => p.Name == selectedItem.Content.ToString());
                
                string newJson = JsonConvert.SerializeObject(presets, Formatting.Indented);
                File.WriteAllText(CUSTOM_PRESETS_FILE, newJson);
                
                CustomPresetList.Items.Remove(selectedItem);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"删除预设时出错：{ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void CustomPresetList_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            var selectedItem = CustomPresetList.SelectedItem as ListBoxItem;
            if (selectedItem == null) return;

            var preset = selectedItem.Tag as GradientPreset;
            if (preset == null) return;

            ApplyPreset(preset);
        }

        private void ApplyPreset(GradientPreset preset)
        {
            // 设置渐变类型
            switch (preset.Type)
            {
                case "Linear": LinearGradientRadio.IsChecked = true; break;
                case "Radial": RadialGradientRadio.IsChecked = true; break;
                case "Conic": ConicGradientRadio.IsChecked = true; break;
                case "Repeating": RepeatingGradientRadio.IsChecked = true; break;
            }

            // 设置参数
            AngleInput.Value = preset.Angle;
            CenterXInput.Value = preset.CenterX;
            CenterYInput.Value = preset.CenterY;
            RepeatCountInput.Value = preset.RepeatCount;

            // 设置颜色节点
            colorStops.Clear();
            foreach (var stop in preset.ColorStops)
            {
                colorStops.Add(new ColorStop { Color = stop.Color, Position = stop.Position });
            }

            UpdateGradient();
        }
        #endregion

        #region 导出功能
        private void ExportCssBtn_Click(object sender, RoutedEventArgs e)
        {
            StringBuilder css = new StringBuilder();
            css.AppendLine("background: ");

            if (LinearGradientRadio.IsChecked == true)
            {
                css.Append($"linear-gradient({AngleInput.Value}deg");
            }
            else if (RadialGradientRadio.IsChecked == true)
            {
                css.Append($"radial-gradient(circle at {CenterXInput.Value}% {CenterYInput.Value}%");
            }
            else if (ConicGradientRadio.IsChecked == true)
            {
                css.Append($"conic-gradient(from {AngleInput.Value}deg at 50% 50%");
            }
            else if (RepeatingGradientRadio.IsChecked == true)
            {
                css.Append($"repeating-linear-gradient({AngleInput.Value}deg");
            }

            foreach (var stop in colorStops)
            {
                css.Append($", {ColorToCssRgba(stop.Color)} {stop.Position}%");
            }
            css.Append(");");

            Clipboard.SetText(css.ToString());
            MessageBox.Show("CSS代码已复制到剪贴板", "导出成功", MessageBoxButton.OK, MessageBoxImage.Information);
        }

        private void ExportXamlBtn_Click(object sender, RoutedEventArgs e)
        {
            StringBuilder xaml = new StringBuilder();
            
            if (LinearGradientRadio.IsChecked == true)
            {
                double angle = AngleInput.Value * Math.PI / 180;
                Point start = new Point(0.5 - Math.Cos(angle) / 2, 0.5 - Math.Sin(angle) / 2);
                Point end = new Point(0.5 + Math.Cos(angle) / 2, 0.5 + Math.Sin(angle) / 2);
                
                xaml.AppendLine("<LinearGradientBrush StartPoint=\"" + start.X.ToString("F3") + "," + start.Y.ToString("F3") + 
                               "\" EndPoint=\"" + end.X.ToString("F3") + "," + end.Y.ToString("F3") + "\">");
            }
            else if (RadialGradientRadio.IsChecked == true)
            {
                xaml.AppendLine($"<RadialGradientBrush Center=\"{CenterXInput.Value/100},{CenterYInput.Value/100}\" " +
                               $"GradientOrigin=\"{CenterXInput.Value/100},{CenterYInput.Value/100}\">");
            }

            foreach (var stop in colorStops)
            {
                xaml.AppendLine($"    <GradientStop Color=\"{ColorToXamlString(stop.Color)}\" Offset=\"{stop.Position/100:F2}\" />");
            }

            xaml.AppendLine(LinearGradientRadio.IsChecked == true ? "</LinearGradientBrush>" : "</RadialGradientBrush>");

            Clipboard.SetText(xaml.ToString());
            MessageBox.Show("XAML代码已复制到剪贴板", "导出成功", MessageBoxButton.OK, MessageBoxImage.Information);
        }

        private string ColorToCssRgba(Color color)
        {
            return $"rgba({color.R},{color.G},{color.B},{color.A/255.0:F2})";
        }

        private string ColorToXamlString(Color color)
        {
            return $"#{color.A:X2}{color.R:X2}{color.G:X2}{color.B:X2}";
        }
        #endregion

        #region PowerPoint操作
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
                    
                    if (LinearGradientRadio.IsChecked == true)
                    {
                        shape.Fill.TwoColorGradient(MsoGradientStyle.msoGradientHorizontal, 1);
                        shape.Fill.GradientAngle = (float)AngleInput.Value;
                        
                        if (colorStops.Count >= 2)
                        {
                            shape.Fill.ForeColor.RGB = ColorToRGB(colorStops[0].Color);
                            shape.Fill.BackColor.RGB = ColorToRGB(colorStops[colorStops.Count - 1].Color);
                        }
                    }
                    else if (RadialGradientRadio.IsChecked == true)
                    {
                        shape.Fill.PresetGradient(MsoGradientStyle.msoGradientFromCenter, 1, MsoPresetGradientType.msoGradientBrass);
                        if (colorStops.Count >= 2)
                        {
                            shape.Fill.ForeColor.RGB = ColorToRGB(colorStops[0].Color);
                            shape.Fill.BackColor.RGB = ColorToRGB(colorStops[colorStops.Count - 1].Color);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"生成渐变时出错：{ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private int ColorToRGB(Color color)
        {
            return color.R + (color.G << 8) + (color.B << 16);
        }
        #endregion
    }
}