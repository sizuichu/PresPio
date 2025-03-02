using System;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Forms;
using System.Windows.Media.Imaging;
using HandyControl.Controls;
using Microsoft.Office.Interop.PowerPoint;

using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace PresPio
    {
    public class ExportSettings : INotifyPropertyChanged
        {
        private int _quality;
        private string _namePattern;
        private bool _openAfterExport;
        private bool _createSubFolder;

        public int Quality
            {
            get => _quality;
            set
                {
                if (_quality != value)
                    {
                    _quality = value;
                    Properties.Settings.Default.JPEGQuality = value;
                    Properties.Settings.Default.Save();
                    OnPropertyChanged(nameof(Quality));
                    }
                }
            }

        public string NamePattern
            {
            get => _namePattern;
            set
                {
                if (_namePattern != value)
                    {
                    _namePattern = value;
                    Properties.Settings.Default.NamePattern = value;
                    Properties.Settings.Default.Save();
                    OnPropertyChanged(nameof(NamePattern));
                    }
                }
            }

        public bool OpenAfterExport
            {
            get => _openAfterExport;
            set
                {
                if (_openAfterExport != value)
                    {
                    _openAfterExport = value;
                    Properties.Settings.Default.OpenAfterExport = value;
                    Properties.Settings.Default.Save();
                    OnPropertyChanged(nameof(OpenAfterExport));
                    }
                }
            }

        public bool CreateSubFolder
            {
            get => _createSubFolder;
            set
                {
                if (_createSubFolder != value)
                    {
                    _createSubFolder = value;
                    Properties.Settings.Default.CreateSubFolder = value;
                    Properties.Settings.Default.Save();
                    OnPropertyChanged(nameof(CreateSubFolder));
                    }
                }
            }

        public event PropertyChangedEventHandler PropertyChanged;

        protected virtual void OnPropertyChanged(string propertyName)
            {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
            }

        public void LoadFromSettings()
            {
            _quality = Properties.Settings.Default.JPEGQuality;
            _namePattern = Properties.Settings.Default.NamePattern;
            _openAfterExport = Properties.Settings.Default.OpenAfterExport;
            _createSubFolder = Properties.Settings.Default.CreateSubFolder;
            }
        }

    public class ExportHistory
        {
        public DateTime ExportTime { get; set; }
        public string Format { get; set; }
        public int DPI { get; set; }
        public string Path { get; set; }
        public int SlideCount { get; set; }
        }

    public partial class Wpf_ImageExport
        {
        public PowerPoint.Application app { get; set; }
        public ExportSettings Settings { get; set; }
        public ObservableCollection<ExportHistory> ExportHistories { get; set; }
        private Progress<int> _exportProgress;
        private int currentSlideIndex = 0;
        private SlideRange selectedSlides;

        public Wpf_ImageExport()
            {
            app = Globals.ThisAddIn.Application;
            Settings = new ExportSettings();
            Settings.LoadFromSettings();
            ExportHistories = new ObservableCollection<ExportHistory>();
            _exportProgress = new Progress<int>(value =>
            {
                Dispatcher.Invoke(() =>
                {
                    ProgressBar.Value = value;
                    ProgressText.Text = $"导出进度: {value}%";
                });
            });

            InitializeComponent();
            LoadDpi();
            LoadFilePath();
            SizeTextBox.Text = GetSize();

            DataContext = this;
            InitializePreview();
            }

        private void InitializePreview()
            {
            if (app.ActiveWindow?.Selection?.Type == PpSelectionType.ppSelectionSlides)
                {
                selectedSlides = app.ActiveWindow.Selection.SlideRange;
                currentSlideIndex = 0;
                UpdatePreview();
                UpdateSlideCountText();
                }
            }

        private void UpdateSlideCountText()
            {
            if (selectedSlides != null && selectedSlides.Count > 0)
                {
                SlideCountText.Text = $"第 {currentSlideIndex + 1} 张，共 {selectedSlides.Count} 张";
                }
            else
                {
                SlideCountText.Text = "未选择幻灯片";
                }
            }

        private void PreviousSlide_Click(object sender, RoutedEventArgs e)
            {
            if (selectedSlides != null && selectedSlides.Count > 0)
                {
                currentSlideIndex--;
                if (currentSlideIndex < 0)
                    {
                    currentSlideIndex = selectedSlides.Count - 1;
                    }
                UpdatePreview();
                UpdateSlideCountText();
                }
            }

        private void NextSlide_Click(object sender, RoutedEventArgs e)
            {
            if (selectedSlides != null && selectedSlides.Count > 0)
                {
                currentSlideIndex++;
                if (currentSlideIndex >= selectedSlides.Count)
                    {
                    currentSlideIndex = 0;
                    }
                UpdatePreview();
                UpdateSlideCountText();
                }
            }

        private void UpdatePreview()
            {
            try
                {
                if (selectedSlides != null && selectedSlides.Count > 0)
                    {
                    var slide = selectedSlides[currentSlideIndex + 1]; // PowerPoint 的索引从1开始
                    string tempPath = Path.Combine(Path.GetTempPath(), $"PresPio_Preview_{Guid.NewGuid()}.png");

                    // 计算预览尺寸，保持原始比例
                    float originalWidth = slide.Master.Width;
                    float originalHeight = slide.Master.Height;
                    float ratio = originalWidth / originalHeight;

                    // 设置预览最大尺寸
                    const int maxPreviewSize = 800;
                    int previewWidth, previewHeight;

                    if (ratio > 1) // 宽度大于高度
                        {
                        previewWidth = maxPreviewSize;
                        previewHeight = (int)(maxPreviewSize / ratio);
                        }
                    else // 高度大于或等于宽度
                        {
                        previewHeight = maxPreviewSize;
                        previewWidth = (int)(maxPreviewSize * ratio);
                        }

                    // 导出临时预览图片
                    slide.Export(tempPath, "PNG", previewWidth, previewHeight);

                    // 加载预览图片
                    using (var stream = new FileStream(tempPath, FileMode.Open, FileAccess.Read))
                        {
                        ImageViewer.ImageSource = BitmapFrame.Create(
                            stream,
                            BitmapCreateOptions.None,
                            BitmapCacheOption.OnLoad
                        );
                        }

                    // 删除临时文件
                    try { File.Delete(tempPath); } catch { }
                    }
                }
            catch (Exception ex)
                {
                Growl.Warning($"预览更新失败: {ex.Message}");
                }
            }

        private async void ExportSlides(Selection sel, string pptFileName, string exportFolderPath, float dpi, string imageFormat)
            {
            try
                {
                // 重置进度条
                Dispatcher.Invoke(() =>
                {
                    ProgressBar.Value = 0;
                    ProgressText.Text = "准备导出...";
                });

                int totalSlides = sel.SlideRange.Count;
                int processedSlides = 0;

                // 如果需要创建子文件夹
                if (Settings.CreateSubFolder)
                    {
                    exportFolderPath = Path.Combine(exportFolderPath, DateTime.Now.ToString("yyyy-MM-dd_HHmmss"));
                    Directory.CreateDirectory(exportFolderPath);
                    }

                foreach (Slide slide in sel.SlideRange.Cast<Slide>())
                    {
                    int width = (int)(slide.Master.Width / 72 * dpi);
                    int height = (int)(slide.Master.Height / 72 * dpi);

                    string fileName = Settings.NamePattern
                        .Replace("{filename}", pptFileName)
                        .Replace("{number}", slide.SlideNumber.ToString("D3"));

                    string fullPath = Path.Combine(exportFolderPath, $"{fileName}.{imageFormat.ToLower()}");

                    await Task.Run(() =>
                    {
                        if (imageFormat == "JPG")
                            {
                            ExportJPEG(slide, fullPath, width, height);
                            }
                        else
                            {
                            slide.Export(fullPath, imageFormat, width, height);
                            }
                    });

                    processedSlides++;
                    var progressPercentage = (int)((double)processedSlides / totalSlides * 100);
                    ((IProgress<int>)_exportProgress).Report(progressPercentage);
                    }

                // 记录导出历史
                Dispatcher.Invoke(() =>
                {
                    ExportHistories.Add(new ExportHistory
                        {
                        ExportTime = DateTime.Now,
                        Format = imageFormat,
                        DPI = (int)dpi,
                        Path = exportFolderPath,
                        SlideCount = totalSlides
                        });
                });

                if (Settings.OpenAfterExport)
                    {
                    Process.Start(exportFolderPath);
                    }

                Dispatcher.Invoke(() =>
                {
                    Growl.Success($"成功导出 {totalSlides} 张幻灯片！");
                    ProgressText.Text = "导出完成";
                });
                }
            catch (Exception ex)
                {
                Dispatcher.Invoke(() =>
                {
                    Growl.Error($"导出失败: {ex.Message}");
                    ProgressText.Text = "导出失败";
                });
                }
            }

        private void ExportJPEG(Slide slide, string path, int width, int height)
            {
            try
                {
                // 先导出为临时文件
                string tempPath = Path.Combine(Path.GetTempPath(), $"PresPio_Temp_{Guid.NewGuid()}.jpg");
                slide.Export(tempPath, "JPG", width, height);

                // 使用新的质量设置重新保存
                using (var sourceImage = System.Drawing.Image.FromFile(tempPath))
                    {
                    var jpegEncoder = System.Drawing.Imaging.ImageCodecInfo.GetImageEncoders()
                        .First(c => c.FormatID == System.Drawing.Imaging.ImageFormat.Jpeg.Guid);

                    var encoderParams = new System.Drawing.Imaging.EncoderParameters(1);
                    encoderParams.Param[0] = new System.Drawing.Imaging.EncoderParameter(
                        System.Drawing.Imaging.Encoder.Quality, Settings.Quality);

                    // 保存到最终目标
                    sourceImage.Save(path, jpegEncoder, encoderParams);
                    }

                // 清理临时文件
                try
                    {
                    if (File.Exists(tempPath))
                        {
                        File.Delete(tempPath);
                        }
                    }
                catch { /* 忽略临时文件删除失败 */ }
                }
            catch (Exception ex)
                {
                Dispatcher.Invoke(() =>
                {
                    Growl.Error($"JPEG导出失败: {ex.Message}");
                });
                throw; // 重新抛出异常以便上层处理
                }
            }

        private void RadioButton_Checked(object sender, RoutedEventArgs e)
            {
            // 移除不必要的提示
            }

        public void LoadFilePath()
            {
            string pptFilePath = app.ActivePresentation.FullName;
            string pptFileName = Path.GetFileNameWithoutExtension(pptFilePath);
            string exportFolderPath = Path.Combine(Path.GetDirectoryName(pptFilePath), $"{pptFileName}_HD");

            Properties.Settings.Default.exportFolderPath = exportFolderPath;
            Properties.Settings.Default.Save();
            filePath.Text = exportFolderPath;
            }

        public void LoadDpi()
            {
            DpiComBox.Items.Clear();
            int defaultDpi = GetDPI();
            double[] dpiValues = { defaultDpi, 75, 100, 150, 300, 350, 400, 450, 500, 800, 1000 };

            foreach (var dpi in dpiValues)
                {
                DpiComBox.Items.Add(dpi);
                }

            if (DpiComBox.Items.Count > 0)
                {
                DpiComBox.SelectedItem = DpiComBox.Items[0];
                }
            }

        public int GetDPI()
            {
            string regPath = @"HKEY_CURRENT_USER\Software\Microsoft\Office";
            double version = double.Parse(app.Version, CultureInfo.InvariantCulture);
            int[] versions = { 16, 15, 14 };
            string[] versionPaths = { @"\16.0\PowerPoint\Options", @"\15.0\PowerPoint\Options", @"\14.0\PowerPoint\Options" };

            for (int i = 0 ; i < versions.Length ; i++)
                {
                if (version == versions[i])
                    {
                    regPath += versionPaths[i];
                    break;
                    }
                }

            object dpiValue = Microsoft.Win32.Registry.GetValue(regPath, "ExportBitmapResolution", null);
            return dpiValue is int dpi ? dpi : 0;
            }

        public void OutImg(string format)
            {
            Selection sel = app.ActiveWindow.Selection;
            if (sel.Type != PpSelectionType.ppSelectionSlides)
                {
                Growl.Warning("请选择幻灯片页面后再试");
                return;
                }

            string pptFilePath = app.ActivePresentation.FullName;
            string pptFileName = Path.GetFileNameWithoutExtension(pptFilePath);
            string exportFolderPath = Properties.Settings.Default.exportFolderPath;

            // 确保导出文件夹存在
            EnsureExportFolderExists(exportFolderPath);

            float dpi = GetDPI();
            string imageFormat = GetImageFormat(format);

            // 使用并行处理导出每个幻灯片
            ExportSlides(sel, pptFileName, exportFolderPath, dpi, imageFormat);

            // 打开导出文件夹
            Process.Start(exportFolderPath);
            }

        private void EnsureExportFolderExists(string exportFolderPath)
            {
            if (!Directory.Exists(exportFolderPath))
                {
                Directory.CreateDirectory(exportFolderPath);
                }
            }

        private string GetImageFormat(string format)
            {
            if (format == "JPEG")
                return "JPG";
            else if (format == "PNG")
                return "PNG";
            else if (format == "GIF")
                return "GIF";
            else if (format == "TIFF")
                return "TIF";
            else if (format == "BMP")
                return "BMP";
            else
                return "PNG"; // 默认格式
            }

        private void Button_Click(object sender, RoutedEventArgs e)
            {
            if (RadioButtonJpg.IsChecked == true)
                {
                OutImg("JPEG");
                }
            else if (RadioButtonPng.IsChecked == true)
                {
                OutImg("PNG");
                }
            else if (RadioButtonGif.IsChecked == true)
                {
                OutImg("GIF");
                }
            else if (RadioButtonTif.IsChecked == true)
                {
                OutImg("TIFF");
                }
            else if (RadioButtonBmp.IsChecked == true)
                {
                OutImg("BMP");
                }
            }

        private void DpiComBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
            {
            if (DpiComBox.SelectedItem is double selectedDpi)
                {
                Properties.Settings.Default.PPtDPI = (int)selectedDpi;
                Properties.Settings.Default.Save();
                DpiSetting();
                SizeTextBox.Text = GetSize();
                UpdatePreview(); // 更新预览
                }
            }

        public void DpiSetting()
            {
            int dpi = Properties.Settings.Default.PPtDPI;
            string regPath = @"HKEY_CURRENT_USER\Software\Microsoft\Office";
            double version = double.Parse(app.Version, CultureInfo.InvariantCulture);

            // 根据 PowerPoint 版本构建注册表路径
            if (version == 16)
                {
                regPath += @"\16.0\PowerPoint\Options";
                }
            else if (version == 15)
                {
                regPath += @"\15.0\PowerPoint\Options";
                }
            else if (version == 14)
                {
                regPath += @"\14.0\PowerPoint\Options";
                }
            else
                {
                throw new InvalidOperationException($"无���处理 PowerPoint {version}。");
                }

            // 设置注册表值
            Microsoft.Win32.Registry.SetValue(regPath, "ExportBitmapResolution", dpi, Microsoft.Win32.RegistryValueKind.DWord);
            }

        public string GetSize()
            {
            Slide slide = app.ActiveWindow.View.Slide;
            float dpi = GetDPI();
            int width = (int)(slide.Master.Width / 72 * dpi);
            int height = (int)(slide.Master.Height / 72 * dpi);
            return $"{width}*{height}";
            }

        private void FolderSelect_Click(object sender, RoutedEventArgs e)
            {
            string exportFolder = Properties.Settings.Default.exportFolderPath;
            using (var folderBrowserDialog = new FolderBrowserDialog())
                {
                folderBrowserDialog.Description = "请选择导出文件夹";
                folderBrowserDialog.ShowNewFolderButton = true;
                folderBrowserDialog.SelectedPath = exportFolder;

                if (folderBrowserDialog.ShowDialog() == System.Windows.Forms.DialogResult.OK && !string.IsNullOrWhiteSpace(folderBrowserDialog.SelectedPath))
                    {
                    exportFolder = folderBrowserDialog.SelectedPath;
                    Properties.Settings.Default.exportFolderPath = exportFolder;
                    filePath.Text = exportFolder;
                    }
                }
            }
        }
    }