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
    public partial class Wpf_ImageTra : INotifyPropertyChanged
        {
        private readonly Application _powerPoint;
        private Bitmap _originalImage;
        private bool _isProcessing = false;

        // 图像参数
        private double _overallOpacity = 100;

        private double _gradientOpacity = 0;
        private double _gradientAngle = 0;
        private double _brightness = 0;
        private double _contrast = 0;
        private double _saturation = 0;
        private double _hue = 0;
        private double _sharpness = 0;
        private bool _isInverted = false;

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
                    UpdatePreview(); // 直接调用预览更新
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
                    UpdatePreview(); // 直接调用预览更新
                    }
                }
            }

        public double Brightness
            {
            get => _brightness;
            set
                {
                if (_brightness != value)
                    {
                    _brightness = value;
                    OnPropertyChanged(nameof(Brightness));
                    UpdatePreview(); // 直接调用预览更新
                    }
                }
            }

        public double Contrast
            {
            get => _contrast;
            set
                {
                if (_contrast != value)
                    {
                    _contrast = value;
                    OnPropertyChanged(nameof(Contrast));
                    UpdatePreview(); // 直接调用预览更新
                    }
                }
            }

        public double Saturation
            {
            get => _saturation;
            set
                {
                if (_saturation != value)
                    {
                    _saturation = value;
                    OnPropertyChanged(nameof(Saturation));
                    UpdatePreview(); // 直接调用预览更新
                    }
                }
            }

        public double Hue
            {
            get => _hue;
            set
                {
                if (_hue != value)
                    {
                    _hue = value;
                    OnPropertyChanged(nameof(Hue));
                    UpdatePreview(); // 直接调用预览更新
                    }
                }
            }

        public double Sharpness
            {
            get => _sharpness;
            set
                {
                if (_sharpness != value)
                    {
                    _sharpness = value;
                    OnPropertyChanged(nameof(Sharpness));
                    UpdatePreview(); // 直接调用预览更新
                    }
                }
            }

        #endregion 属性定义

        private bool _isInPreviewUpdate = false;

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
                string tempPath = Path.GetTempFileName() + ".png";
                shape.Export(tempPath, PpShapeFormat.ppShapeFormatPNG);

                _originalImage = new Bitmap(tempPath);
                PreviewImage.ImageSource = (BitmapFrame)ConvertBitmapToBitmapSource(_originalImage);

                File.Delete(tempPath);
                }
            catch (Exception ex)
                {
                HandyControl.Controls.MessageBox.Show($"加载图片失败: {ex.Message}", "错误");
                Close();
                }
            }

        // 执行图像处理并更新预览
        private void UpdatePreview()
            {
            if (_originalImage == null || _isProcessing) return;
            _isProcessing = true;

            try
                {
                // 创建原图的副本进行处理，避免原图占用
                using (var originalImageCopy = new Bitmap(_originalImage))
                    {
                    using (var processedImage = ProcessImage(originalImageCopy))
                        {
                        PreviewImage.ImageSource = (BitmapFrame)ConvertBitmapToBitmapSource(processedImage);
                        }
                    }
                }
            finally
                {
                _isProcessing = false;
                }
            }

        // 图像处理方法（如亮度、对比度等）
        private Bitmap ProcessImage(Bitmap source)
            {
            var result = new Bitmap(source.Width, source.Height);
            using (Graphics g = Graphics.FromImage(result))
                {
                g.InterpolationMode = InterpolationMode.HighQualityBicubic;
                g.SmoothingMode = SmoothingMode.HighQuality;

                // 创建颜色矩阵
                float brightness = (float)(_brightness + 100) / 100;
                float contrast = (float)(_contrast + 100) / 100;
                float saturation = (float)(_saturation + 100) / 100;

                float[][] matrixItems = {
                    new float[] {contrast * saturation, 0, 0, 0, 0},
                    new float[] {0, contrast * saturation, 0, 0, 0},
                    new float[] {0, 0, contrast * saturation, 0, 0},
                    new float[] {0, 0, 0, (float)(_overallOpacity/100), 0},
                    new float[] {brightness - 1, brightness - 1, brightness - 1, 0, 1}
                };

                var colorMatrix = new ColorMatrix(matrixItems);
                var imageAttributes = new ImageAttributes();
                imageAttributes.SetColorMatrix(colorMatrix);

                // 应用渐变透明度
                if (Math.Abs(_gradientOpacity) > 0)
                    {
                    float gradientOpacityValue = (float)Math.Abs(_gradientOpacity) / 100;
                    using (var gradientBrush = new LinearGradientBrush(
                        new System.Drawing.Point(0, 0),
                        new System.Drawing.Point(
                            (int)(source.Width * Math.Cos(_gradientAngle * Math.PI / 180)),
                            (int)(source.Height * Math.Sin(_gradientAngle * Math.PI / 180))),
                        Color.FromArgb((int)(255 * (1 - gradientOpacityValue)), Color.White),
                        Color.FromArgb(255, Color.White)))
                        {
                        g.FillRectangle(gradientBrush, 0, 0, source.Width, source.Height);
                        }
                    }

                // 应用锐化
                if (_sharpness != 0)
                    {
                    float sharpness = (float)(_sharpness / 100);
                    float[][] sharpenMatrix = {
                        new float[] {-sharpness, -sharpness, -sharpness},
                        new float[] {-sharpness, 1 + (8 * sharpness), -sharpness},
                        new float[] {-sharpness, -sharpness, -sharpness}
                    };
                    imageAttributes.SetColorMatrix(new ColorMatrix(sharpenMatrix));
                    }

                g.DrawImage(source,
                    new System.Drawing.Rectangle(0, 0, result.Width, result.Height),
                    0, 0, source.Width, source.Height,
                    GraphicsUnit.Pixel,
                    imageAttributes);
                }

            return result;
            }

        // 将Bitmap转换为BitmapSource
        private BitmapSource ConvertBitmapToBitmapSource(Bitmap bitmap)
            {
            using (MemoryStream memory = new MemoryStream())
                {
                bitmap.Save(memory, ImageFormat.Png);
                memory.Position = 0;

                var decoder = BitmapDecoder.Create(
                    memory,
                    BitmapCreateOptions.PreservePixelFormat,
                    BitmapCacheOption.OnLoad);

                return decoder.Frames[0];
                }
            }

        #region 滤镜预设

        private void OnFilterSoft(object sender, RoutedEventArgs e)
            {
            Brightness = 10;
            Contrast = -10;
            Saturation = -20;
            Sharpness = -30;
            UpdatePreview();
            }

        private void OnFilterVintage(object sender, RoutedEventArgs e)
            {
            Brightness = -10;
            Contrast = 20;
            Saturation = -30;
            Hue = 15;
            UpdatePreview();
            }

        private void OnFilterBW(object sender, RoutedEventArgs e)
            {
            Saturation = -100;
            Contrast = 20;
            UpdatePreview();
            }

        private void OnFilterInvert(object sender, RoutedEventArgs e)
            {
            _isInverted = !_isInverted;
            UpdatePreview();
            }

        private void OnFilterSharpen(object sender, RoutedEventArgs e)
            {
            Sharpness = 50;
            Contrast = 10;
            UpdatePreview();
            }

        #endregion 滤镜预设

        private void OnDirectionChanged(object sender, RoutedEventArgs e)
            {
            if (sender is System.Windows.Controls.RadioButton radioButton &&
                radioButton.Tag is string angleStr)
                {
                if (double.TryParse(angleStr, out double angle))
                    {
                    _gradientAngle = angle;
                    UpdatePreview();
                    }
                }
            }

        private void OnReset(object sender, RoutedEventArgs e)
            {
            OverallOpacity = 100;
            GradientOpacity = 0;
            Brightness = 0;
            Contrast = 0;
            Saturation = 0;
            Hue = 0;
            Sharpness = 0;
            _gradientAngle = 0;
            _isInverted = false;
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