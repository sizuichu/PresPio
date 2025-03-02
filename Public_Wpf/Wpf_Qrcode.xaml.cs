using System;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using HandyControl.Controls;
using QRCoder;
using OpenFileDialog = System.Windows.Forms.OpenFileDialog;
using SaveFileDialog = System.Windows.Forms.SaveFileDialog;

namespace PresPio
    {
    /// <summary>
    /// Wpf_Qrcode.xaml 的交互逻辑
    /// </summary>
    public partial class Wpf_Qrcode
        {
        private System.Windows.Media.Color foregroundColor = Colors.Black;
        private System.Windows.Media.Color backgroundColor = Colors.White;
        private string logoPath = null;

        public Wpf_Qrcode()
            {
            InitializeComponent();

            // 初始化默认值
            SizeComboBox.SelectedIndex = 0;
            ErrorCorrectionComboBox.SelectedIndex = 0;
            }

        private void CreatQr_Click(object sender, RoutedEventArgs e)
            {
            string qrText = InputTextBox.Text.Trim();
            if (string.IsNullOrEmpty(qrText))
                {
                Growl.Error("请输入二维码内容。");
                return;
                }

            // 获取二维码大小
            string sizeStr = ((ComboBoxItem)SizeComboBox.SelectedItem).Content.ToString();
            int size = 200; // 默认大小
            if (sizeStr.Contains("300"))
                size = 300;
            else if (sizeStr.Contains("400"))
                size = 400;

            // 获取错误纠正级别
            string eccStr = ((ComboBoxItem)ErrorCorrectionComboBox.SelectedItem).Content.ToString();
            QRCodeGenerator.ECCLevel eccLevel = QRCodeGenerator.ECCLevel.H;
            switch (eccStr)
                {
                case "中 (M)":
                    eccLevel = QRCodeGenerator.ECCLevel.M;
                    break;

                case "较高 (Q)":
                    eccLevel = QRCodeGenerator.ECCLevel.Q;
                    break;

                case "高 (H)":
                    eccLevel = QRCodeGenerator.ECCLevel.H;
                    break;

                default:
                    eccLevel = QRCodeGenerator.ECCLevel.L;
                    break;
                }

            try
                {
                using (QRCodeGenerator qrGenerator = new QRCodeGenerator())
                    {
                    QRCodeData qrCodeData = qrGenerator.CreateQrCode(qrText, eccLevel);
                    QRCode qrCode = new QRCode(qrCodeData);
                    Bitmap qrCodeImage = qrCode.GetGraphic(20, System.Drawing.Color.FromArgb(foregroundColor.R, foregroundColor.G, foregroundColor.B),
                                                            System.Drawing.Color.FromArgb(backgroundColor.R, backgroundColor.G, backgroundColor.B),
                                                            logoPath != null ? new Bitmap(logoPath) : null, 25, 2, true, null);

                    QrImageBox.Source = BitmapToImageSource(qrCodeImage);
                    }
                }
            catch (Exception ex)
                {
                Growl.Error($"生成二维码时出错：{ex.Message}");
                }
            }

        private void SaveQr2png_Click(object sender, RoutedEventArgs e)
            {
            if (QrImageBox.Source == null)
                {
                Growl.Warning("请先生成二维码。");
                return;
                }

            SaveFileDialog saveFileDialog = new SaveFileDialog
                {
                Filter = "PNG 图片 (*.png)|*.png",
                FileName = "qrcode.png"
                };

            if (saveFileDialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                using (FileStream fs = new FileStream(saveFileDialog.FileName, FileMode.Create))
                    {
                    PngBitmapEncoder encoder = new PngBitmapEncoder();
                    encoder.Frames.Add(BitmapFrame.Create((BitmapSource)QrImageBox.Source));
                    encoder.Save(fs);
                    }
                Growl.SuccessGlobal("二维码已保存。");
                }
            }

        private void CopyQr_Click(object sender, RoutedEventArgs e)
            {
            if (QrImageBox.Source == null)
                {
                Growl.WarningGlobal("请先生成二维码。");
                return;
                }

            BitmapSource bitmapSource = (BitmapSource)QrImageBox.Source;
            Bitmap bitmap = BitmapSourceToBitmap(bitmapSource);

            System.Windows.Forms.Clipboard.SetImage(bitmap);
            Growl.SuccessGlobal("二维码已复制到剪贴板。");
            }

        private void InsetQr_Click(object sender, RoutedEventArgs e)
            {
            if (QrImageBox.Source == null)
                {
                Growl.WarningGlobal("请先生成二维码。");
                return;
                }

            try
                {
                BitmapSource bitmapSource = (BitmapSource)QrImageBox.Source;
                Bitmap bitmap = BitmapSourceToBitmap(bitmapSource);
                System.Windows.Forms.Clipboard.SetImage(bitmap);

                // 使用PowerPoint的COM对象插入图片
                var pptApp = Globals.ThisAddIn.Application;
                var activeSlide = pptApp.ActiveWindow.View.Slide;
                activeSlide.Shapes.Paste();

                Growl.SuccessGlobal("二维码已复制到当前PPT页面。");
                }
            catch (Exception ex)
                {
                Growl.ErrorGlobal($"将二维码复制到PPT时出错：{ex.Message}");
                }
            }

        private Bitmap BitmapSourceToBitmap(BitmapSource bitmapSource)
            {
            Bitmap bitmap;
            using (MemoryStream outStream = new MemoryStream())
                {
                BitmapEncoder enc = new BmpBitmapEncoder();
                enc.Frames.Add(BitmapFrame.Create(bitmapSource));
                enc.Save(outStream);
                bitmap = new Bitmap(outStream);
                }
            return bitmap;
            }

        private BitmapSource BitmapToImageSource(Bitmap bitmap)
            {
            using (MemoryStream memoryStream = new MemoryStream())
                {
                bitmap.Save(memoryStream, ImageFormat.Png);
                memoryStream.Position = 0;
                BitmapImage bitmapImage = new BitmapImage();
                bitmapImage.BeginInit();
                bitmapImage.StreamSource = memoryStream;
                bitmapImage.CacheOption = BitmapCacheOption.OnLoad;
                bitmapImage.EndInit();
                bitmapImage.Freeze();
                return bitmapImage;
                }
            }

        private void ForegroundColorButton_Click(object sender, RoutedEventArgs e)
            {
            System.Windows.Forms.ColorDialog colorDialog = new System.Windows.Forms.ColorDialog
                {
                Color = System.Drawing.Color.FromArgb(foregroundColor.A, foregroundColor.R, foregroundColor.G, foregroundColor.B)
                };

            if (colorDialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                foregroundColor = System.Windows.Media.Color.FromArgb(colorDialog.Color.A, colorDialog.Color.R, colorDialog.Color.G, colorDialog.Color.B);
                ForegroundColorButton.Background = new SolidColorBrush(foregroundColor);
                }
            }

        private void BackgroundColorButton_Click(object sender, RoutedEventArgs e)
            {
            System.Windows.Forms.ColorDialog colorDialog = new System.Windows.Forms.ColorDialog
                {
                Color = System.Drawing.Color.FromArgb(backgroundColor.A, backgroundColor.R, backgroundColor.G, backgroundColor.B)
                };

            if (colorDialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                backgroundColor = System.Windows.Media.Color.FromArgb(colorDialog.Color.A, colorDialog.Color.R, colorDialog.Color.G, colorDialog.Color.B);
                BackgroundColorButton.Background = new SolidColorBrush(backgroundColor);
                }
            }

        private void AddLogoButton_Click(object sender, RoutedEventArgs e)
            {
            OpenFileDialog openFileDialog = new OpenFileDialog
                {
                Title = "选择Logo图片",
                Filter = "图像文件|*.png;*.jpg;*.jpeg;*.bmp;*.gif"
                };

            if (openFileDialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                logoPath = openFileDialog.FileName;
                Growl.SuccessGlobal("Logo已添加。");
                }
            }
        }
    }