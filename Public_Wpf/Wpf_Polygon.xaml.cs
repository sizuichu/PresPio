using System;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using MessageBox = HandyControl.Controls.MessageBox;
using Point = System.Windows.Point;
using Window = System.Windows.Window;

namespace PresPio.Public_Wpf
    {
    /// <summary>
    /// Wpf_Polygon.xaml 的交互逻辑
    /// </summary>
    public partial class Wpf_Polygon : Window
        {
        public Wpf_Polygon()
            {
            InitializeComponent();
            }

        private void Window_Loaded(object sender, RoutedEventArgs e)
            {
            GeneratePreview();
            }

        private void Parameter_ValueChanged(object sender, HandyControl.Data.FunctionEventArgs<double> e)
            {
            GeneratePreview();
            }

        private void ShapeType_SelectionChanged(object sender, SelectionChangedEventArgs e)
            {
            if (PolygonPanel == null) return; // 防止初始化时的空引用

            // 隐藏所有参数面板
            PolygonPanel.Visibility = Visibility.Collapsed;
            StarPanel.Visibility = Visibility.Collapsed;
            EllipsePanel.Visibility = Visibility.Collapsed;
            ArrowPanel.Visibility = Visibility.Collapsed;
            CrossPanel.Visibility = Visibility.Collapsed;

            // 根据选择显示相应的参数面板
            switch (((ComboBoxItem)ShapeTypeComboBox.SelectedItem).Content.ToString())
                {
                case "正多边形":
                    PolygonPanel.Visibility = Visibility.Visible;
                    break;

                case "星形":
                    StarPanel.Visibility = Visibility.Visible;
                    break;

                case "椭圆":
                    EllipsePanel.Visibility = Visibility.Visible;
                    break;

                case "箭头":
                    ArrowPanel.Visibility = Visibility.Visible;
                    break;

                case "十字形":
                    CrossPanel.Visibility = Visibility.Visible;
                    break;

                case "心形":
                    break;

                case "菱形":
                    break;
                }

            GeneratePreview();
            }

        private void GeneratePreview()
            {
            if (ShapeTypeComboBox == null || PreviewPolygon == null || RadiusInput == null) return;

            var points = new PointCollection();
            string shapeType = ((ComboBoxItem)ShapeTypeComboBox.SelectedItem).Content.ToString();

            try
                {
                switch (shapeType)
                    {
                    case "正多边形":
                        if (SidesInput != null)
                            points = GenerateRegularPolygon();
                        break;

                    case "星形":
                        if (StarPointsInput != null && StarInnerRadiusInput != null)
                            points = GenerateStarShape();
                        break;

                    case "椭圆":
                        if (MajorAxisInput != null && MinorAxisInput != null)
                            points = GenerateEllipse();
                        break;

                    case "箭头":
                        if (ArrowWidthInput != null && ArrowHeadLengthInput != null)
                            points = GenerateArrow();
                        break;

                    case "十字形":
                        if (CrossArmLengthInput != null && CrossArmWidthInput != null)
                            points = GenerateCross();
                        break;

                    case "心形":
                        points = GenerateHeart();
                        break;

                    case "菱形":
                        points = GenerateDiamond();
                        break;
                    }

                if (points.Count > 0)
                    {
                    PreviewPolygon.Points = points;
                    Canvas.SetLeft(PreviewPolygon, 0);
                    Canvas.SetTop(PreviewPolygon, 0);
                    }
                }
            catch (Exception)
                {
                // 忽略初始化过程中的异常
                }
            }

        private PointCollection GenerateRegularPolygon()
            {
            if (SidesInput == null || RadiusInput == null || RotationInput == null || PreviewCanvas == null)
                return new PointCollection();

            int sides = (int)SidesInput.Value;
            double radius = RadiusInput.Value;
            double rotation = RotationInput.Value;
            double centerX = PreviewCanvas.ActualWidth / 2;
            double centerY = PreviewCanvas.ActualHeight / 2;

            if (centerX == 0) centerX = 200;
            if (centerY == 0) centerY = 200;

            var points = new PointCollection();
            for (int i = 0 ; i < sides ; i++)
                {
                double angle = (i * (360.0 / sides) + rotation) * Math.PI / 180.0;
                double x = centerX + radius * Math.Cos(angle);
                double y = centerY + radius * Math.Sin(angle);
                points.Add(new Point(x, y));
                }

            return points;
            }

        private PointCollection GenerateStarShape()
            {
            if (StarPointsInput == null || RadiusInput == null || StarInnerRadiusInput == null ||
                RotationInput == null || PreviewCanvas == null)
                return new PointCollection();

            int points = (int)StarPointsInput.Value;
            double radius = RadiusInput.Value;
            double innerRadius = radius * StarInnerRadiusInput.Value;
            double rotation = RotationInput.Value;
            double centerX = PreviewCanvas.ActualWidth / 2;
            double centerY = PreviewCanvas.ActualHeight / 2;

            if (centerX == 0) centerX = 200;
            if (centerY == 0) centerY = 200;

            var starPoints = new PointCollection();
            for (int i = 0 ; i < points * 2 ; i++)
                {
                double angle = (i * (360.0 / (points * 2)) + rotation) * Math.PI / 180.0;
                double currentRadius = i % 2 == 0 ? radius : innerRadius;
                double x = centerX + currentRadius * Math.Cos(angle);
                double y = centerY + currentRadius * Math.Sin(angle);
                starPoints.Add(new Point(x, y));
                }

            return starPoints;
            }

        private PointCollection GenerateEllipse()
            {
            if (MajorAxisInput == null || MinorAxisInput == null || RotationInput == null || PreviewCanvas == null)
                return new PointCollection();

            double majorAxis = MajorAxisInput.Value;
            double minorAxis = MinorAxisInput.Value;
            double rotation = RotationInput.Value * Math.PI / 180.0;
            double centerX = PreviewCanvas.ActualWidth / 2;
            double centerY = PreviewCanvas.ActualHeight / 2;

            if (centerX == 0) centerX = 200;
            if (centerY == 0) centerY = 200;

            var points = new PointCollection();
            int segments = 36; // 使用36个点来近似椭圆
            for (int i = 0 ; i < segments ; i++)
                {
                double angle = i * (2 * Math.PI / segments);
                double x = majorAxis * Math.Cos(angle);
                double y = minorAxis * Math.Sin(angle);

                // 应用旋转
                double rotatedX = x * Math.Cos(rotation) - y * Math.Sin(rotation);
                double rotatedY = x * Math.Sin(rotation) + y * Math.Cos(rotation);

                points.Add(new Point(centerX + rotatedX, centerY + rotatedY));
                }

            return points;
            }

        private PointCollection GenerateArrow()
            {
            if (ArrowWidthInput == null || ArrowHeadLengthInput == null ||
                RadiusInput == null || RotationInput == null || PreviewCanvas == null)
                return new PointCollection();

            double radius = RadiusInput.Value;
            double width = ArrowWidthInput.Value;
            double headLength = ArrowHeadLengthInput.Value;
            double rotation = RotationInput.Value * Math.PI / 180.0;
            double centerX = PreviewCanvas.ActualWidth / 2;
            double centerY = PreviewCanvas.ActualHeight / 2;

            if (centerX == 0) centerX = 200;
            if (centerY == 0) centerY = 200;

            // 箭头的基本点（未旋转）
            var basePoints = new[]
            {
                new Point(-radius, -radius * width),        // 箭身左下
                new Point(radius * (1-headLength), -radius * width),  // 箭身右下
                new Point(radius * (1-headLength), -radius * width * 2), // 箭头左下
                new Point(radius, 0),                       // 箭头尖端
                new Point(radius * (1-headLength), radius * width * 2),  // 箭头右上
                new Point(radius * (1-headLength), radius * width),   // 箭身右上
                new Point(-radius, radius * width),         // 箭身左上
            };

            var points = new PointCollection();
            foreach (var point in basePoints)
                {
                // 应用旋转
                double rotatedX = point.X * Math.Cos(rotation) - point.Y * Math.Sin(rotation);
                double rotatedY = point.X * Math.Sin(rotation) + point.Y * Math.Cos(rotation);
                points.Add(new Point(centerX + rotatedX, centerY + rotatedY));
                }

            return points;
            }

        private PointCollection GenerateCross()
            {
            if (CrossArmLengthInput == null || CrossArmWidthInput == null ||
                RadiusInput == null || RotationInput == null || PreviewCanvas == null)
                return new PointCollection();

            double radius = RadiusInput.Value;
            double armLength = CrossArmLengthInput.Value;
            double armWidth = CrossArmWidthInput.Value;
            double rotation = RotationInput.Value * Math.PI / 180.0;
            double centerX = PreviewCanvas.ActualWidth / 2;
            double centerY = PreviewCanvas.ActualHeight / 2;

            if (centerX == 0) centerX = 200;
            if (centerY == 0) centerY = 200;

            // 十字架的基本点（未旋转）
            var basePoints = new[]
            {
                new Point(-radius * armWidth, -radius),        // 上臂左
                new Point(radius * armWidth, -radius),         // 上臂右
                new Point(radius * armWidth, -radius * armLength), // 上臂到右臂连接点
                new Point(radius, -radius * armLength),        // 右臂上
                new Point(radius, radius * armLength),         // 右臂下
                new Point(radius * armWidth, radius * armLength),  // 右臂到下臂连接点
                new Point(radius * armWidth, radius),          // 下臂右
                new Point(-radius * armWidth, radius),         // 下臂左
                new Point(-radius * armWidth, radius * armLength), // 下臂到左臂连接点
                new Point(-radius, radius * armLength),        // 左臂下
                new Point(-radius, -radius * armLength),       // 左臂上
                new Point(-radius * armWidth, -radius * armLength) // 左臂到上臂连接点
            };

            var points = new PointCollection();
            foreach (var point in basePoints)
                {
                // 应用旋转
                double rotatedX = point.X * Math.Cos(rotation) - point.Y * Math.Sin(rotation);
                double rotatedY = point.X * Math.Sin(rotation) + point.Y * Math.Cos(rotation);
                points.Add(new Point(centerX + rotatedX, centerY + rotatedY));
                }

            return points;
            }

        private PointCollection GenerateHeart()
            {
            if (RadiusInput == null || RotationInput == null || PreviewCanvas == null)
                return new PointCollection();

            double radius = RadiusInput.Value;
            double rotation = RotationInput.Value * Math.PI / 180.0;
            double centerX = PreviewCanvas.ActualWidth / 2;
            double centerY = PreviewCanvas.ActualHeight / 2;

            if (centerX == 0) centerX = 200;
            if (centerY == 0) centerY = 200;

            var points = new PointCollection();
            int segments = 36;
            for (int i = 0 ; i < segments ; i++)
                {
                double t = i * (2 * Math.PI / segments);
                // 心形曲线方程
                double x = 16 * Math.Pow(Math.Sin(t), 3);
                double y = 13 * Math.Cos(t) - 5 * Math.Cos(2 * t) - 2 * Math.Cos(3 * t) - Math.Cos(4 * t);

                // 缩放
                x = x * radius / 16;
                y = -y * radius / 16; // 反转Y轴使心形朝上

                // 应用旋转
                double rotatedX = x * Math.Cos(rotation) - y * Math.Sin(rotation);
                double rotatedY = x * Math.Sin(rotation) + y * Math.Cos(rotation);

                points.Add(new Point(centerX + rotatedX, centerY + rotatedY));
                }

            return points;
            }

        private PointCollection GenerateDiamond()
            {
            if (RadiusInput == null || RotationInput == null || PreviewCanvas == null)
                return new PointCollection();

            double radius = RadiusInput.Value;
            double rotation = RotationInput.Value * Math.PI / 180.0;
            double centerX = PreviewCanvas.ActualWidth / 2;
            double centerY = PreviewCanvas.ActualHeight / 2;

            if (centerX == 0) centerX = 200;
            if (centerY == 0) centerY = 200;

            // 菱形的四个顶点
            var basePoints = new[]
            {
                new Point(0, -radius),      // 上
                new Point(radius * 0.7, 0), // 右
                new Point(0, radius),       // 下
                new Point(-radius * 0.7, 0) // 左
            };

            var points = new PointCollection();
            foreach (var point in basePoints)
                {
                // 应用旋转
                double rotatedX = point.X * Math.Cos(rotation) - point.Y * Math.Sin(rotation);
                double rotatedY = point.X * Math.Sin(rotation) + point.Y * Math.Cos(rotation);
                points.Add(new Point(centerX + rotatedX, centerY + rotatedY));
                }

            return points;
            }

        private void InsertToPPT_Click(object sender, RoutedEventArgs e)
            {
            try
                {
                var app = Globals.ThisAddIn.Application;
                var slide = app.ActiveWindow.View.Slide;

                if (slide == null)
                    {
                    MessageBox.Show("请先选择一个幻灯片！", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
                    return;
                    }

                var points = PreviewPolygon.Points;
                float[,] pptPoints = new float[points.Count + 1, 2];

                // 转换点坐标
                for (int i = 0 ; i < points.Count ; i++)
                    {
                    pptPoints[i, 0] = (float)(XPositionInput.Value + (points[i].X - PreviewCanvas.ActualWidth / 2));
                    pptPoints[i, 1] = (float)(YPositionInput.Value + (points[i].Y - PreviewCanvas.ActualHeight / 2));
                    }

                // 闭合图形
                pptPoints[points.Count, 0] = pptPoints[0, 0];
                pptPoints[points.Count, 1] = pptPoints[0, 1];

                // 创建图形
                var shape = slide.Shapes.AddPolyline(pptPoints);

                this.Close();
                }
            catch (Exception ex)
                {
                MessageBox.Show($"插入图形时发生错误：{ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }

        protected override void OnRenderSizeChanged(SizeChangedInfo sizeInfo)
            {
            base.OnRenderSizeChanged(sizeInfo);
            GeneratePreview();
            }
        }
    }