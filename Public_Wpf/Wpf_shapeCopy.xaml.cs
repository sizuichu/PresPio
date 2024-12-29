using HandyControl.Controls;
using Microsoft.Office.Interop.PowerPoint;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using System;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using System.Windows.Shapes;
using System.Linq;
using Shape = Microsoft.Office.Interop.PowerPoint.Shape;
using TabItem = System.Windows.Controls.TabItem;

namespace PresPio
{
    public partial class Wpf_shapeCopy 
    {
        private PowerPoint.Application app;
        private const string COPY_TAG = "ShapeArrayCopy";
        private double centerX;
        private double centerY;

        public Wpf_shapeCopy()
        {
            InitializeComponent();
            this.Loaded += Window_Loaded;
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            try
            {
                app = Globals.ThisAddIn.Application;
                if (copyModeTab != null && matrixSettings != null)
                {
                    copyModeTab.SelectedIndex = 0;
                    matrixSettings.Visibility = Visibility.Visible;
                    UpdatePreview();
                }
            }
            catch (Exception ex)
            {
                HandyControl.Controls.MessageBox.Show($"初始化失败\n{ex.Message}", "错误");
            }
        }

        private void OnValueChanged(object sender, HandyControl.Data.FunctionEventArgs<double> e)
        {
            UpdatePreview();
        }

        private void OnCheckBoxChanged(object sender, RoutedEventArgs e)
        {
            UpdatePreview();
        }

        private void TabControl_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (copyModeTab?.SelectedItem is TabItem selectedTab && settingsPanel != null)
            {
                try
                {
                    // 隐藏所有设置面板
                    foreach (var child in settingsPanel.Children)
                    {
                        if (child is GroupBox groupBox)
                        {
                            groupBox.Visibility = Visibility.Collapsed;
                        }
                    }

                    // 显示选中的设置面板
                    var tag = selectedTab.Tag?.ToString();
                    GroupBox targetSettings = null;
                    
                    switch (tag)
                    {
                        case "Matrix":
                            targetSettings = matrixSettings;
                            break;
                        case "Circle":
                            targetSettings = circleSettings;
                            break;
                        case "Diagonal":
                            targetSettings = diagonalSettings;
                            break;
                        case "Spiral":
                            targetSettings = spiralSettings;
                            break;
                        case "Wave":
                            targetSettings = waveSettings;
                            break;
                        case "Radial":
                            targetSettings = radialSettings;
                            break;
                        case "Grid":
                            targetSettings = gridSettings;
                            break;
                    }

                    if (targetSettings != null)
                    {
                        targetSettings.Visibility = Visibility.Visible;
                        // 确保ScrollViewer滚动到顶部
                        if (targetSettings.Parent is StackPanel panel && 
                            panel.Parent is HandyControl.Controls.ScrollViewer viewer)
                        {
                            viewer.ScrollToTop();
                        }
                    }

                    UpdatePreview();
                }
                catch (Exception ex)
                {
                    HandyControl.Controls.MessageBox.Show($"切换标签页失败\n{ex.Message}", "错误");
                }
            }
        }

        private void UpdatePreview()
        {
            if (previewCanvas == null || settingsPanel == null) return;

            try
            {
                previewCanvas.Children.Clear();

                // 计算中心点
                centerX = previewCanvas.ActualWidth / 2;
                centerY = previewCanvas.ActualHeight / 2;

                var visibleSettings = settingsPanel.Children.Cast<UIElement>()
                    .OfType<GroupBox>()
                    .FirstOrDefault(g => g.Visibility == Visibility.Visible);

                if (visibleSettings == null) return;

                switch (visibleSettings.Name)
                {
                    case "matrixSettings":
                        DrawMatrixPreview();
                        break;
                    case "circleSettings":
                        DrawCircularPreview();
                        break;
                    case "diagonalSettings":
                        DrawDiagonalPreview();
                        break;
                    case "spiralSettings":
                        DrawSpiralPreview();
                        break;
                    case "waveSettings":
                        DrawWavePreview();
                        break;
                    case "radialSettings":
                        DrawRadialPreview();
                        break;
                    case "gridSettings":
                        DrawGridPreview();
                        break;
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"预览更新失败: {ex.Message}");
            }
        }

        private void DrawShape(double x, double y, double size, double angle = 0)
        {
            Rectangle rect = new Rectangle
            {
                Width = size,
                Height = size,
                Fill = new SolidColorBrush(Colors.LightBlue),
                Stroke = new SolidColorBrush(Colors.Blue),
                StrokeThickness = 1
            };

            if (angle != 0)
            {
                rect.RenderTransform = new RotateTransform(angle, size / 2, size / 2);
            }

            Canvas.SetLeft(rect, x - size / 2);
            Canvas.SetTop(rect, y - size / 2);
            previewCanvas.Children.Add(rect);
        }

        // 预览方法实现...
        private void DrawMatrixPreview()
        {
            int rows = (int)matrixRows.Value;
            int columns = (int)matrixColumns.Value;
            double hSpacing = (double)matrixHSpacing.Value * 0.2;
            double vSpacing = (double)matrixVSpacing.Value * 0.2;

            double maxWidth = previewCanvas.ActualWidth / (columns + 1);
            double maxHeight = previewCanvas.ActualHeight / (rows + 1);
            double size = Math.Min(maxWidth, maxHeight) * 0.6;

            double totalWidth = columns * size + (columns - 1) * hSpacing;
            double totalHeight = rows * size + (rows - 1) * vSpacing;

            double startX = (previewCanvas.ActualWidth - totalWidth) / 2;
            double startY = (previewCanvas.ActualHeight - totalHeight) / 2;

            for (int i = 0; i < rows; i++)
            {
                for (int j = 0; j < columns; j++)
                {
                    double x = startX + j * (size + hSpacing) + size / 2;
                    double y = startY + i * (size + vSpacing) + size / 2;
                    DrawShape(x, y, size);
                }
            }
        }

        private void DrawCircularPreview()
        {
            int count = (int)circleCount.Value;
            double radius = (double)circleRadius.Value * 0.4;
            double startAng = (double)startAngle.Value;
            bool rotate = rotateAlign.IsChecked ?? false;

            double size = Math.Min(20, 180.0 / count);

            for (int i = 0; i < count; i++)
            {
                double angle = startAng + (360.0 / count * i);
                double radians = angle * Math.PI / 180;

                double x = centerX + radius * Math.Cos(radians);
                double y = centerY + radius * Math.Sin(radians);
                
                DrawShape(x, y, size, rotate ? angle : 0);
            }
        }

        private void DrawDiagonalPreview()
        {
            int count = (int)diagonalCount.Value;
            double spacing = (double)diagonalSpacing.Value * 0.2;
            double angle = (double)diagonalAngle.Value;
            double scale = (double)diagonalScale.Value / 100.0;

            double size = Math.Min(20, 180.0 / count);
            double radians = angle * Math.PI / 180;
            double scaleStep = (scale - 1.0) / (count - 1);

            double startX = centerX - (count - 1) * spacing * Math.Cos(radians) / 2;
            double startY = centerY - (count - 1) * spacing * Math.Sin(radians) / 2;

            for (int i = 0; i < count; i++)
            {
                double currentScale = 1 + (scaleStep * i);
                double x = startX + i * spacing * Math.Cos(radians);
                double y = startY + i * spacing * Math.Sin(radians);
                DrawShape(x, y, size * currentScale, angle);
            }
        }

        private void DrawSpiralPreview()
        {
            int turns = (int)spiralTurns.Value;
            int countPerTurn = (int)spiralCount.Value;
            double startRadius = (double)spiralRadius.Value * 0.2;
            double radiusInc = (double)spiralInc.Value * 0.2;
            bool rotate = spiralRotate.IsChecked ?? false;

            double size = Math.Min(15, 180.0 / (turns * countPerTurn));

            for (int turn = 0; turn < turns; turn++)
            {
                double currentRadius = startRadius + (turn * radiusInc);
                for (int i = 0; i < countPerTurn; i++)
                {
                    double angle = (360.0 / countPerTurn * i) + (turn * 360.0 / countPerTurn);
                    double radians = angle * Math.PI / 180;

                    double x = centerX + currentRadius * Math.Cos(radians);
                    double y = centerY + currentRadius * Math.Sin(radians);
                    
                    DrawShape(x, y, size, rotate ? angle : 0);
                }
            }
        }

        private void DrawWavePreview()
        {
            int count = (int)waveCount.Value;
            double wavelength = (double)waveLength.Value * 0.2;
            double amplitude = (double)waveAmplitude.Value * 0.2;
            double phase = (double)wavePhase.Value;

            double size = Math.Min(15, 180.0 / count);
            double stepX = previewCanvas.ActualWidth / (count + 1);

            for (int i = 0; i < count; i++)
            {
                double x = stepX * (i + 1);
                double angle = (x / wavelength * 360 + phase) * Math.PI / 180;
                double y = centerY + amplitude * Math.Sin(angle);
                
                DrawShape(x, y, size);
            }
        }

        private void DrawRadialPreview()
        {
            int count = (int)radialCount.Value;
            double startRadius = (double)radialStartRadius.Value * 0.2;
            double radiusInc = (double)radialRadiusInc.Value * 0.2;
            double angleInc = (double)radialAngleInc.Value;
            bool rotate = radialRotate.IsChecked ?? false;

            double size = Math.Min(15, 180.0 / count);
            double currentAngle = 0;
            double currentRadius = startRadius;

            for (int i = 0; i < count; i++)
            {
                double radians = currentAngle * Math.PI / 180;
                double x = centerX + currentRadius * Math.Cos(radians);
                double y = centerY + currentRadius * Math.Sin(radians);
                
                DrawShape(x, y, size, rotate ? currentAngle : 0);

                currentAngle += angleInc;
                currentRadius += radiusInc;
            }
        }

        private void DrawGridPreview()
        {
            int rows = (int)gridRows.Value;
            int columns = (int)gridColumns.Value;
            double cellSize = (double)gridSize.Value * 0.2;
            double angle = (double)gridAngle.Value;
            double offset = (double)gridOffset.Value * 0.2;

            double size = Math.Min(15, 180.0 / Math.Max(rows, columns));
            double radians = angle * Math.PI / 180;

            double totalWidth = columns * cellSize;
            double totalHeight = rows * cellSize;
            double startX = (previewCanvas.ActualWidth - totalWidth) / 2;
            double startY = (previewCanvas.ActualHeight - totalHeight) / 2;

            Random rand = new Random();

            for (int i = 0; i <= rows; i++)
            {
                for (int j = 0; j <= columns; j++)
                {
                    if (i == 0 && j == 0) continue;

                    double x = startX + j * cellSize + cellSize / 2;
                    double y = startY + i * cellSize + cellSize / 2;

                    if (offset > 0)
                    {
                        x += (rand.NextDouble() * 2 - 1) * offset;
                        y += (rand.NextDouble() * 2 - 1) * offset;
                    }

                    if (angle != 0)
                    {
                        double rotatedX = centerX + (x - centerX) * Math.Cos(radians) - (y - centerY) * Math.Sin(radians);
                        double rotatedY = centerY + (x - centerX) * Math.Sin(radians) + (y - centerY) * Math.Cos(radians);
                        x = rotatedX;
                        y = rotatedY;
                    }

                    DrawShape(x, y, size, angle);
                }
            }
        }

        private void ResetButton_Click(object sender, RoutedEventArgs e)
        {
            var visibleSettings = settingsPanel.Children.Cast<UIElement>()
                .OfType<GroupBox>()
                .FirstOrDefault(g => g.Visibility == Visibility.Visible);

            if (visibleSettings == null) return;

            switch (visibleSettings.Name)
            {
                case "matrixSettings":
                    matrixRows.Value = 2;
                    matrixColumns.Value = 2;
                    matrixHSpacing.Value = 50;
                    matrixVSpacing.Value = 50;
                    break;
                case "circleSettings":
                    circleCount.Value = 8;
                    circleRadius.Value = 100;
                    startAngle.Value = 0;
                    rotateAlign.IsChecked = true;
                    break;
                case "diagonalSettings":
                    diagonalCount.Value = 5;
                    diagonalSpacing.Value = 50;
                    diagonalAngle.Value = 45;
                    diagonalScale.Value = 100;
                    break;
                case "spiralSettings":
                    spiralTurns.Value = 2;
                    spiralCount.Value = 8;
                    spiralRadius.Value = 50;
                    spiralInc.Value = 20;
                    spiralRotate.IsChecked = true;
                    break;
                case "waveSettings":
                    waveCount.Value = 10;
                    waveLength.Value = 100;
                    waveAmplitude.Value = 50;
                    wavePhase.Value = 0;
                    break;
                case "radialSettings":
                    radialCount.Value = 8;
                    radialStartRadius.Value = 50;
                    radialRadiusInc.Value = 30;
                    radialAngleInc.Value = 15;
                    radialRotate.IsChecked = true;
                    break;
                case "gridSettings":
                    gridColumns.Value = 3;
                    gridRows.Value = 3;
                    gridSize.Value = 100;
                    gridAngle.Value = 0;
                    gridOffset.Value = 0;
                    break;
            }
            UpdatePreview();
        }

        private void ApplyButton_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (app?.ActiveWindow?.Selection == null)
                {
                    HandyControl.Controls.MessageBox.Show("PowerPoint未准备就绪", "错误");
                    return;
                }

                var selection = app.ActiveWindow.Selection;
                if (selection.Type != PowerPoint.PpSelectionType.ppSelectionShapes)
                {
                    HandyControl.Controls.MessageBox.Show("请先选择要复制的形状！", "提示");
                    return;
                }

                CleanupPreviousCopies();

                var visibleSettings = settingsPanel.Children.Cast<UIElement>()
                    .OfType<GroupBox>()
                    .FirstOrDefault(g => g.Visibility == Visibility.Visible);

                if (visibleSettings == null) return;

                switch (visibleSettings.Name)
                {
                    case "matrixSettings":
                        ApplyMatrixCopy(selection.ShapeRange);
                        break;
                    case "circleSettings":
                        ApplyCircularCopy(selection.ShapeRange);
                        break;
                    case "diagonalSettings":
                        ApplyDiagonalCopy(selection.ShapeRange);
                        break;
                    case "spiralSettings":
                        ApplySpiralCopy(selection.ShapeRange);
                        break;
                    case "waveSettings":
                        ApplyWaveCopy(selection.ShapeRange);
                        break;
                    case "radialSettings":
                        ApplyRadialCopy(selection.ShapeRange);
                        break;
                    case "gridSettings":
                        ApplyGridCopy(selection.ShapeRange);
                        break;
                }
            }
            catch (Exception ex)
            {
                HandyControl.Controls.MessageBox.Show($"复制操作失败\n{ex.Message}", "错误");
            }
        }

        private void CleanupPreviousCopies()
        {
            var slide = app.ActiveWindow.View.Slide;
            var selection = app.ActiveWindow.Selection;
            var selectedShapes = selection.ShapeRange;

            for (int i = slide.Shapes.Count; i >= 1; i--)
            {
                var shape = slide.Shapes[i];
                try
                {
                    bool isSelected = false;
                    foreach (Shape selectedShape in selectedShapes)
                    {
                        if (shape.Id == selectedShape.Id)
                        {
                            isSelected = true;
                            break;
                        }
                    }

                    if (!isSelected && shape.Tags["CopyType"] == COPY_TAG)
                    {
                        shape.Delete();
                    }
                }
                catch { }
            }
        }

        // 复制方法实现...
        private void ApplyMatrixCopy(ShapeRange shapes)
        {
            int rows = (int)matrixRows.Value;
            int columns = (int)matrixColumns.Value;
            float hSpacing = (float)matrixHSpacing.Value;
            float vSpacing = (float)matrixVSpacing.Value;

            float baseLeft = shapes.Left;
            float baseTop = shapes.Top;

            for (int i = 0; i < rows; i++)
            {
                for (int j = 0; j < columns; j++)
                {
                    if (i == 0 && j == 0) continue;

                    var newShape = shapes.Duplicate()[1];
                    newShape.Left = baseLeft + j * (shapes.Width + hSpacing);
                    newShape.Top = baseTop + i * (shapes.Height + vSpacing);
                    newShape.Tags.Add("CopyType", COPY_TAG);
                }
            }
        }

        private void ApplyCircularCopy(ShapeRange shapes)
        {
            int count = (int)circleCount.Value;
            float radius = (float)circleRadius.Value;
            float startAng = (float)startAngle.Value;
            bool rotate = rotateAlign.IsChecked ?? false;

            float centerX = shapes.Left + shapes.Width / 2;
            float centerY = shapes.Top + shapes.Height / 2;

            for (int i = 1; i <= count; i++)
            {
                double angle = startAng + (360.0 / count * i);
                double radians = angle * Math.PI / 180;

                var newShape = shapes.Duplicate()[1];
                float x = (float)(centerX + radius * Math.Cos(radians) - shapes.Width / 2);
                float y = (float)(centerY + radius * Math.Sin(radians) - shapes.Height / 2);

                newShape.Left = x;
                newShape.Top = y;
                if (rotate)
                {
                    newShape.Rotation = (float)angle;
                }
                newShape.Tags.Add("CopyType", COPY_TAG);
            }
        }

        private void ApplyDiagonalCopy(ShapeRange shapes)
        {
            int count = (int)diagonalCount.Value;
            float spacing = (float)diagonalSpacing.Value;
            float angle = (float)diagonalAngle.Value;
            float scale = (float)diagonalScale.Value / 100f;

            float radians = angle * (float)Math.PI / 180f;
            float scaleStep = (scale - 1.0f) / (count - 1);

            float centerX = shapes.Left + shapes.Width / 2;
            float centerY = shapes.Top + shapes.Height / 2;
            float startX = centerX - (count - 1) * spacing * (float)Math.Cos(radians) / 2;
            float startY = centerY - (count - 1) * spacing * (float)Math.Sin(radians) / 2;

            for (int i = 1; i < count; i++)
            {
                float currentScale = 1 + (scaleStep * i);
                var newShape = shapes.Duplicate()[1];

                float x = startX + i * spacing * (float)Math.Cos(radians);
                float y = startY + i * spacing * (float)Math.Sin(radians);

                newShape.Left = x - shapes.Width / 2;
                newShape.Top = y - shapes.Height / 2;
                newShape.Rotation = angle;

                try
                {
                    newShape.ScaleHeight(currentScale, Microsoft.Office.Core.MsoTriState.msoFalse);
                    newShape.ScaleWidth(currentScale, Microsoft.Office.Core.MsoTriState.msoFalse);
                }
                catch
                {
                    newShape.Height = shapes.Height * currentScale;
                    newShape.Width = shapes.Width * currentScale;
                }

                newShape.Tags.Add("CopyType", COPY_TAG);
            }
        }

        private void ApplySpiralCopy(ShapeRange shapes)
        {
            int turns = (int)spiralTurns.Value;
            int countPerTurn = (int)spiralCount.Value;
            float startRadius = (float)spiralRadius.Value;
            float radiusInc = (float)spiralInc.Value;
            bool rotate = spiralRotate.IsChecked ?? false;

            float centerX = shapes.Left + shapes.Width / 2;
            float centerY = shapes.Top + shapes.Height / 2;

            for (int turn = 0; turn < turns; turn++)
            {
                float currentRadius = startRadius + (turn * radiusInc);
                for (int i = 0; i < countPerTurn; i++)
                {
                    if (turn == 0 && i == 0) continue;

                    double angle = (360.0 / countPerTurn * i) + (turn * 360.0 / countPerTurn);
                    double radians = angle * Math.PI / 180;

                    var newShape = shapes.Duplicate()[1];
                    float x = (float)(centerX + currentRadius * Math.Cos(radians) - shapes.Width / 2);
                    float y = (float)(centerY + currentRadius * Math.Sin(radians) - shapes.Height / 2);

                    newShape.Left = x;
                    newShape.Top = y;
                    if (rotate)
                    {
                        newShape.Rotation = (float)angle;
                    }
                    newShape.Tags.Add("CopyType", COPY_TAG);
                }
            }
        }

        private void ApplyWaveCopy(ShapeRange shapes)
        {
            int count = (int)waveCount.Value;
            float wavelength = (float)waveLength.Value;
            float amplitude = (float)waveAmplitude.Value;
            float phase = (float)wavePhase.Value;

            float stepX = wavelength / count;
            float baseLeft = shapes.Left;
            float baseTop = shapes.Top + shapes.Height / 2;

            for (int i = 1; i < count; i++)
            {
                var newShape = shapes.Duplicate()[1];
                float x = baseLeft + i * stepX;
                double angle = (x / wavelength * 360 + phase) * Math.PI / 180;
                float y = baseTop + (float)(amplitude * Math.Sin(angle)) - shapes.Height / 2;

                newShape.Left = x;
                newShape.Top = y;
                newShape.Tags.Add("CopyType", COPY_TAG);
            }
        }

        private void ApplyRadialCopy(ShapeRange shapes)
        {
            int count = (int)radialCount.Value;
            float startRadius = (float)radialStartRadius.Value;
            float radiusInc = (float)radialRadiusInc.Value;
            float angleInc = (float)radialAngleInc.Value;
            bool rotate = radialRotate.IsChecked ?? false;

            float centerX = shapes.Left + shapes.Width / 2;
            float centerY = shapes.Top + shapes.Height / 2;

            float currentAngle = 0;
            float currentRadius = startRadius;

            for (int i = 1; i < count; i++)
            {
                double radians = currentAngle * Math.PI / 180;
                var newShape = shapes.Duplicate()[1];

                float x = (float)(centerX + currentRadius * Math.Cos(radians) - shapes.Width / 2);
                float y = (float)(centerY + currentRadius * Math.Sin(radians) - shapes.Height / 2);

                newShape.Left = x;
                newShape.Top = y;
                if (rotate)
                {
                    newShape.Rotation = currentAngle;
                }
                newShape.Tags.Add("CopyType", COPY_TAG);

                currentAngle += angleInc;
                currentRadius += radiusInc;
            }
        }

        private void ApplyGridCopy(ShapeRange shapes)
        {
            int rows = (int)gridRows.Value;
            int columns = (int)gridColumns.Value;
            float cellSize = (float)gridSize.Value;
            float angle = (float)gridAngle.Value;
            float offset = (float)gridOffset.Value;

            float radians = angle * (float)Math.PI / 180f;
            float centerX = shapes.Left + shapes.Width / 2;
            float centerY = shapes.Top + shapes.Height / 2;

            float totalWidth = columns * cellSize;
            float totalHeight = rows * cellSize;
            float startX = shapes.Left - totalWidth / 2 + shapes.Width / 2;
            float startY = shapes.Top - totalHeight / 2 + shapes.Height / 2;

            Random rand = new Random();

            for (int i = 0; i < rows; i++)
            {
                for (int j = 0; j < columns; j++)
                {
                    if (i == 0 && j == 0) continue;

                    float x = startX + j * cellSize;
                    float y = startY + i * cellSize;

                    if (offset > 0)
                    {
                        x += (float)(rand.NextDouble() * 2 - 1) * offset;
                        y += (float)(rand.NextDouble() * 2 - 1) * offset;
                    }

                    var newShape = shapes.Duplicate()[1];

                    if (angle != 0)
                    {
                        float rotatedX = centerX + (x - centerX) * (float)Math.Cos(radians) - (y - centerY) * (float)Math.Sin(radians);
                        float rotatedY = centerY + (x - centerX) * (float)Math.Sin(radians) + (y - centerY) * (float)Math.Cos(radians);
                        x = rotatedX;
                        y = rotatedY;
                    }

                    newShape.Left = x;
                    newShape.Top = y;
                    newShape.Rotation = angle;
                    newShape.Tags.Add("CopyType", COPY_TAG);
                }
            }
        }
    }
}
