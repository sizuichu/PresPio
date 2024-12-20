using HandyControl.Controls;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using System;
using System.ComponentModel;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.IO;
using System.Windows;
using System.Windows.Media.Imaging;
using Application = Microsoft.Office.Interop.PowerPoint.Application;

namespace PresPio
    {
    public partial class Wpf_ImageTra
        {
        private readonly Application _powerPoint;
        private Bitmap _originalImage;
        private bool _isProcessing = false;

        // 图像参数
        private double _overallOpacity = 100;
        private double _gradientOpacity = 0;
        private double _gradientAngle = 0;

        public event PropertyChangedEventHandler PropertyChanged;

        #region 属性定义
        public double OverallOpacity
            {
            get => _overallOpacity;
            set
                {
                if (_overallOpacity != value)
                    {
                    _overallOpacity = value;
                    OnPropertyChanged(nameof(OverallOpacity));
                    UpdatePreview();
                    }
                }
            }

        public double GradientOpacity
            {
            get => _gradientOpacity;
            set
                {
                if (_gradientOpacity != value)
                    {
                    _gradientOpacity = value;
                    OnPropertyChanged(nameof(GradientOpacity));
                    UpdatePreview();
                    }
                }
            }
        #endregion

        public Wpf_ImageTra()
            {
            InitializeComponent();
            DataContext = this;
            _powerPoint = Globals.ThisAddIn.Application;

            LoadSelectedImage();
            }

        private void LoadSelectedImage()
            {
            try
                {
                var selection = _powerPoint.ActiveWindow.Selection;
                if (selection.Type != PpSelectionType.ppSelectionShapes || selection.ShapeRange.Count != 1)
                    {
                    HandyControl.Controls.MessageBox.Show("请选择一个图片对象", "提示");
                    Close();
                    return;
                    }

                var shape = selection.ShapeRange[1];
                string tempPath = Path.GetTempFileName();
                string imagePath = tempPath + ".png";

                try
                    {
                    // 导出图片
                    shape.Export(imagePath, PpShapeFormat.ppShapeFormatPNG);

                    using (var fileStream = new FileStream(imagePath, FileMode.Open, FileAccess.Read))
                        {
                        var bitmap = new Bitmap(fileStream);
                        _originalImage = new Bitmap(bitmap);
                        bitmap.Dispose();
                        }

                    UpdatePreviewImage();
                    }
                finally
                    {
                    // 清理临时文件
                    if (File.Exists(tempPath)) File.Delete(tempPath);
                    if (File.Exists(imagePath)) File.Delete(imagePath);
                    }
                }
            catch (Exception ex)
                {
                HandyControl.Controls.MessageBox.Show($"加载图片失败: {ex.Message}", "错误");
                Close();
                }
            }

        private void UpdatePreviewImage()
            {
            if (_originalImage == null) return;

            try
                {
                using (var processedImage = ProcessImage(_originalImage))
                using (var ms = new MemoryStream())
                    {
                    processedImage.Save(ms, ImageFormat.Png);
                    ms.Position = 0;

                    var bitmapImage = new BitmapImage();
                    bitmapImage.BeginInit();
                    bitmapImage.CacheOption = BitmapCacheOption.OnLoad;
                    bitmapImage.StreamSource = ms;
                    bitmapImage.EndInit();
                    bitmapImage.Freeze();

                    PreviewImage.Source = bitmapImage;
                    }
                }
            catch (Exception ex)
                {
                HandyControl.Controls.MessageBox.Show($"更新预览图片失败: {ex.Message}", "错误");
                }
            }

        private void UpdatePreview()
            {
            if (_originalImage == null || _isProcessing) return;
            _isProcessing = true;

            try
                {
                UpdatePreviewImage();
                }
            finally
                {
                _isProcessing = false;
                }
            }

        private Bitmap ProcessImage(Bitmap source)
            {
            var result = new Bitmap(source.Width, source.Height);

            try
                {
                using (Graphics g = Graphics.FromImage(result))
                    {
                    g.InterpolationMode = InterpolationMode.HighQualityBicubic;
                    g.SmoothingMode = SmoothingMode.HighQuality;

                    var imageAttributes = new ImageAttributes();

                    // 创建颜色矩阵处理整体透明度
                    ColorMatrix colorMatrix = new ColorMatrix(new float[][]
                    {
                        new float[] {1, 0, 0, 0, 0},
                        new float[] {0, 1, 0, 0, 0},
                        new float[] {0, 0, 1, 0, 0},
                        new float[] {0, 0, 0, (float)(_overallOpacity/100), 0},
                        new float[] {0, 0, 0, 0, 1}
                    });

                    imageAttributes.SetColorMatrix(colorMatrix);

                    // 应用渐变透明度
                    if (_gradientOpacity > 0)
                        {
                        float gradientOpacityValue = (float)_gradientOpacity / 100;
                        double angleRad = _gradientAngle * Math.PI / 180;

                        float centerX = source.Width / 2f;
                        float centerY = source.Height / 2f;
                        float radius = (float)Math.Sqrt(source.Width * source.Width + source.Height * source.Height) / 2f;

                        PointF startPoint = new PointF(
                            centerX - (float)(Math.Cos(angleRad) * radius),
                            centerY - (float)(Math.Sin(angleRad) * radius)
                        );

                        PointF endPoint = new PointF(
                            centerX + (float)(Math.Cos(angleRad) * radius),
                            centerY + (float)(Math.Sin(angleRad) * radius)
                        );

                        using (var gradientBrush = new LinearGradientBrush(
                            startPoint,
                            endPoint,
                            Color.FromArgb((int)(255 * (1 - gradientOpacityValue)), Color.White),
                            Color.FromArgb(255, Color.White)))
                            {
                            g.FillRectangle(gradientBrush, 0, 0, source.Width, source.Height);
                            }
                        }

                    g.DrawImage(source,
                        new Rectangle(0, 0, result.Width, result.Height),
                        0, 0, source.Width, source.Height,
                        GraphicsUnit.Pixel,
                        imageAttributes);

                    imageAttributes.Dispose();
                    }

                return result;
                }
            catch
                {
                result.Dispose();
                throw;
                }
            }

        #region 透明度预设
        private void OnPresetFull(object sender, RoutedEventArgs e)
            {
            OverallOpacity = 100;
            GradientOpacity = 0;
            _gradientAngle = 0;
            UpdatePreview();
            }

        private void OnPresetLight(object sender, RoutedEventArgs e)
            {
            OverallOpacity = 80;
            GradientOpacity = 0;
            _gradientAngle = 0;
            UpdatePreview();
            }

        private void OnPresetMedium(object sender, RoutedEventArgs e)
            {
            OverallOpacity = 50;
            GradientOpacity = 0;
            _gradientAngle = 0;
            UpdatePreview();
            }

        private void OnPresetHigh(object sender, RoutedEventArgs e)
            {
            OverallOpacity = 20;
            GradientOpacity = 0;
            _gradientAngle = 0;
            UpdatePreview();
            }

        private void OnPresetFadeLeftRight(object sender, RoutedEventArgs e)
            {
            OverallOpacity = 100;
            GradientOpacity = 80;
            _gradientAngle = 0;
            UpdatePreview();
            }

        private void OnPresetFadeTopBottom(object sender, RoutedEventArgs e)
            {
            OverallOpacity = 100;
            GradientOpacity = 80;
            _gradientAngle = 90;
            UpdatePreview();
            }

        private void OnPresetFadeDiagonal(object sender, RoutedEventArgs e)
            {
            OverallOpacity = 100;
            GradientOpacity = 80;
            _gradientAngle = 45;
            UpdatePreview();
            }
        #endregion

        #region 渐变方向控制
        private void OnHorizontalGradient(object sender, RoutedEventArgs e)
            {
            _gradientAngle = 0;
            UpdatePreview();
            }

        private void OnVerticalGradient(object sender, RoutedEventArgs e)
            {
            _gradientAngle = 90;
            UpdatePreview();
            }

        private void OnDiagonalGradient(object sender, RoutedEventArgs e)
            {
            _gradientAngle = 45;
            UpdatePreview();
            }

        private void OnReverseGradient(object sender, RoutedEventArgs e)
            {
            _gradientAngle = (_gradientAngle + 180) % 360;
            UpdatePreview();
            }
        #endregion

        private void OnReset(object sender, RoutedEventArgs e)
            {
            OverallOpacity = 100;
            GradientOpacity = 0;
            _gradientAngle = 0;
            UpdatePreview();
            }

        private void OnCancel(object sender, RoutedEventArgs e)
            {
            Close();
            }

        private void OnApply(object sender, RoutedEventArgs e)
            {
            try
                {
                var selection = _powerPoint.ActiveWindow.Selection;
                if (selection.Type != PpSelectionType.ppSelectionShapes ||
                    selection.ShapeRange.Count != 1)
                    {
                    HandyControl.Controls.MessageBox.Show("请选择一个图片对象", "提示");
                    return;
                    }

                var shape = selection.ShapeRange[1];
                var left = shape.Left;
                var top = shape.Top;
                var width = shape.Width;
                var height = shape.Height;

                using (var processedImage = ProcessImage(_originalImage))
                    {
                    string tempPath = Path.GetTempFileName() + ".png";
                    processedImage.Save(tempPath, ImageFormat.Png);

                    shape.Delete();
                    var newShape = _powerPoint.ActiveWindow.View.Slide.Shapes.AddPicture(
                        tempPath,
                        MsoTriState.msoFalse,
                        MsoTriState.msoTrue,
                        left, top, width, height);

                    File.Delete(tempPath);
                    }

                Close();
                }
            catch (Exception ex)
                {
                HandyControl.Controls.MessageBox.Show($"应用更改失败: {ex.Message}", "错误");
                }
            }

        protected void OnPropertyChanged(string propertyName)
            {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
            }
        }
    }