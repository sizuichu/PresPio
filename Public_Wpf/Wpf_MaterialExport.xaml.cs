using System;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.IO;
using System.Windows;
using System.Windows.Forms;
using MessageBox = System.Windows.MessageBox;
using System.Linq;
using System.Windows.Controls;
using System.Runtime.InteropServices;
using System.Threading.Tasks;
using System.Windows.Media.Imaging;
using System.Windows.Input;
using Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using System.IO.Packaging;
using A = DocumentFormat.OpenXml.Drawing;
using System.Windows.Media;
using System.Windows.Threading;
using ProgressBar = System.Windows.Controls.ProgressBar;

namespace PresPio.Public_Wpf
{
    public partial class Wpf_MaterialExport : Window
    {
        private ObservableCollection<MaterialItem> Materials { get; set; }
        private ObservableCollection<MaterialItem> FilteredMaterials { get; set; }
        private PowerPoint.Application pptApplication;
        private PowerPoint.Presentation activePresentation;
        private string tempFolder;
        private bool isExternalPptApp = true;  // 标记是否是外部传入的PowerPoint应用程序

        public Wpf_MaterialExport(PowerPoint.Application app)
        {
            InitializeComponent();
            pptApplication = app;
            isExternalPptApp = true;

            // 创建临时文件夹
            tempFolder = Path.Combine(Path.GetTempPath(), "PresPio_Export_" + Guid.NewGuid().ToString());
            Directory.CreateDirectory(tempFolder);
            Directory.CreateDirectory(Path.Combine(tempFolder, "thumbnails")); // 创建缩略图子文件夹

            InitializeControls();
        }

        private void InitializeControls()
        {
            Materials = new ObservableCollection<MaterialItem>();
            FilteredMaterials = new ObservableCollection<MaterialItem>();
            lvMaterials.ItemsSource = FilteredMaterials;

            // 初始化进度条
            mainProgressBar.Visibility = Visibility.Collapsed;
            mainProgressText.Visibility = Visibility.Collapsed;

            // 绑定事件处理
            btnSelectAll.Click += BtnSelectAll_Click;
            btnUnselectAll.Click += BtnUnselectAll_Click;
            btnBrowse.Click += BtnBrowse_Click;
            btnExport.Click += BtnExport_Click;
            btnCancel.Click += BtnCancel_Click;
            cmbExportType.SelectionChanged += CmbExportType_SelectionChanged;

            // 设置默认选项
            cmbExportType.SelectedIndex = 0;

            // 加载素材列表
            LoadMaterials();
        }

        private void ShowProgress(bool show, string message = "", double value = 0)
        {
            Dispatcher.Invoke(() =>
            {
                mainProgressBar.Visibility = show ? Visibility.Visible : Visibility.Collapsed;
                mainProgressText.Visibility = show ? Visibility.Visible : Visibility.Collapsed;
                if (show)
                {
                    mainProgressBar.Value = value;
                    mainProgressText.Text = message;
                }
            });
        }

        private async void LoadMaterials()
        {
            try
            {
                Mouse.OverrideCursor = System.Windows.Input.Cursors.Wait;
                ShowProgress(true, "正在加载素材...", 0);

                if (pptApplication == null)
                {
                    try
                    {
                        // 如果没有传入PowerPoint应用程序，尝试获取当前运行的实例
                        pptApplication = Marshal.GetActiveObject("PowerPoint.Application") as PowerPoint.Application;
                        isExternalPptApp = false;
                    }
                    catch
                    {
                        MessageBox.Show("无法获取PowerPoint应用程序", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
                        return;
                    }
                }

                try
                {
                    activePresentation = pptApplication.ActivePresentation;
                    if (activePresentation == null)
                    {
                        MessageBox.Show("请先打开一个PowerPoint演示文稿", "提示", MessageBoxButton.OK, MessageBoxImage.Warning);
                        return;
                    }

                    string pptPath = activePresentation.FullName;
                    if (string.IsNullOrEmpty(pptPath))
                    {
                        MessageBox.Show("请先保存PowerPoint演示文稿", "提示", MessageBoxButton.OK, MessageBoxImage.Warning);
                        return;
                    }

                    // 设置默认导出路径为PPT所在文件夹
                    string defaultExportPath = Path.Combine(Path.GetDirectoryName(pptPath), "导出的素材");
                    txtExportPath.Text = defaultExportPath;

                    Materials.Clear();

                    int totalSlides = activePresentation.Slides.Count;
                    int currentSlide = 0;

                    foreach (PowerPoint.Slide slide in activePresentation.Slides)
                    {
                        currentSlide++;
                        ShowProgress(true, $"正在处理幻灯片 ({currentSlide}/{totalSlides})", (currentSlide / (double)totalSlides) * 100);

                        foreach (PowerPoint.Shape shape in slide.Shapes)
                        {
                            try
                            {
                                MaterialItem item = null;

                                // 处理图片
                                if (shape.Type == MsoShapeType.msoPicture)
                                {
                                    try
                                    {
                                        string fileName = $"image_{Guid.NewGuid()}.png";
                                        string tempFilePath = Path.Combine(tempFolder, fileName);
                                        
                                        // 导出图片
                                        try
                                        {
                                            shape.Export(tempFilePath, PowerPoint.PpShapeFormat.ppShapeFormatPNG);
                                            
                                            // 确保文件已经被创建
                                            if (File.Exists(tempFilePath))
                                            {
                                                item = new MaterialItem
                                                {
                                                    Name = string.IsNullOrEmpty(shape.Name) ? fileName : shape.Name,
                                                    Type = "图片",
                                                    FilePath = tempFilePath,
                                                    CreateTime = DateTime.Now,
                                                    SlideNumber = slide.SlideNumber
                                                };

                                                // 生成缩略图
                                                await GenerateThumbnail(item);
                                            }
                                        }
                                        catch (Exception ex)
                                        {
                                            System.Diagnostics.Debug.WriteLine($"导出图片失败: {ex.Message}");
                                            continue;
                                        }
                                    }
                                    catch
                                    {
                                        continue;
                                    }
                                }
                                // 处理媒体（视频和音频）
                                else if (shape.Type == MsoShapeType.msoMedia)
                                {
                                    string mediaType = "未知";
                                    string extension = ".mp4"; // 默认扩展名

                                    try
                                    {
                                        PowerPoint.MediaFormat mediaFormat = null;
                                        try
                                        {
                                            mediaFormat = shape.MediaFormat;
                                        }
                                        catch
                                        {
                                            continue; // 如果无法获取MediaFormat，跳过此项
                                        }

                                        if (mediaFormat != null)
                                        {
                                            // 尝试从OLEFormat获取信息
                                            if (shape.OLEFormat != null)
                                            {
                                                string progID = shape.OLEFormat.ProgID ?? "";
                                                
                                                // 根据ProgID判断媒体类型
                                                if (progID.Contains("Video"))
                                                {
                                                    mediaType = "视频";
                                                    extension = ".mp4";
                                                }
                                                else if (progID.Contains("Audio") || progID.Contains("Sound"))
                                                {
                                                    mediaType = "音频";
                                                    extension = ".mp3";
                                                }
                                            }

                                            // 如果还未确定类型，尝试从名称判断
                                            if (mediaType == "未知")
                                            {
                                                string name = shape.Name.ToLower();
                                                if (name.Contains("video") || name.EndsWith(".mp4") || name.EndsWith(".avi") || 
                                                    name.EndsWith(".wmv") || name.EndsWith(".mov"))
                                                {
                                                    mediaType = "视频";
                                                    extension = GetExtensionFromName(name);
                                                }
                                                else if (name.Contains("audio") || name.EndsWith(".mp3") || name.EndsWith(".wav") || 
                                                         name.EndsWith(".wma"))
                                                {
                                                    mediaType = "音频";
                                                    extension = GetExtensionFromName(name);
                                                }
                                            }

                                            string fileName = $"{mediaType}_{Guid.NewGuid()}{extension}";
                                            string tempFilePath = Path.Combine(tempFolder, fileName);

                                            try
                                            {
                                                // 尝试导出媒体文件
                                                shape.Export(tempFilePath, PowerPoint.PpShapeFormat.ppShapeFormatPNG);

                                                item = new MaterialItem
                                                {
                                                    Name = shape.Name,
                                                    Type = mediaType,
                                                    FilePath = tempFilePath,
                                                    CreateTime = DateTime.Now
                                                };

                                                // 设置默认缩略图
                                                if (mediaType == "视频")
                                                {
                                                    item.ThumbnailPath = "pack://application:,,,/PresPio;component/Resources/video_thumbnail.png";
                                                }
                                                else if (mediaType == "音频")
                                                {
                                                    item.ThumbnailPath = "pack://application:,,,/PresPio;component/Resources/audio_thumbnail.png";
                                                }
                                            }
                                            catch
                                            {
                                                continue; // 如果导出失败，跳过此项
                                            }
                                        }
                                    }
                                    catch
                                    {
                                        continue; // 如果无法获取媒体信息，跳过此项
                                    }
                                }

                                if (item != null)
                                {
                                    Materials.Add(item);
                                    await UpdateFileInfo(item);
                                }
                            }
                            catch
                            {
                                // 如果处理单个形状时出错，继续处理下一个
                                continue;
                            }
                        }
                    }

                    ApplyFilter();

                    if (Materials.Count == 0)
                    {
                        MessageBox.Show("未找到任何可导出的素材", "提示", MessageBoxButton.OK, MessageBoxImage.Information);
                    }
                }
                catch (COMException comEx)
                {
                    MessageBox.Show("无法访问PowerPoint演示文稿，请确保文件未被锁定且PowerPoint运行正常。", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
                    return;
                }
                catch (Exception ex) when (ex is InvalidCastException || ex is System.Runtime.InteropServices.InvalidComObjectException)
                {
                    MessageBox.Show("PowerPoint应用程序状态异常，请重新打开PowerPoint。", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
                    return;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"加载素材时出错：{ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                ShowProgress(false);
                Mouse.OverrideCursor = null;
            }
        }

        private async Task GenerateThumbnail(MaterialItem item)
        {
            if (string.IsNullOrEmpty(item.FilePath) || !File.Exists(item.FilePath))
            {
                item.ThumbnailPath = "pack://application:,,,/PresPio;component/Resources/image_thumbnail.png";
                return;
            }

            try
            {
                // 使用专门的缩略图文件夹
                string thumbnailsFolder = Path.Combine(tempFolder, "thumbnails");
                string thumbnailPath = Path.Combine(thumbnailsFolder, $"thumb_{Path.GetFileName(item.FilePath)}");
                
                using (var fileStream = new FileStream(item.FilePath, FileMode.Open, FileAccess.Read))
                {
                    BitmapImage image = new BitmapImage();
                    image.BeginInit();
                    image.CacheOption = BitmapCacheOption.OnLoad;
                    image.StreamSource = fileStream;
                    image.EndInit();
                    image.Freeze(); // 重要：防止内存泄漏

                    var encoder = new JpegBitmapEncoder();
                    var bitmap = new TransformedBitmap(image, new ScaleTransform(
                        100.0 / image.PixelWidth,
                        100.0 / image.PixelHeight));
                    
                    encoder.Frames.Add(BitmapFrame.Create(bitmap));
                    
                    using (var outputStream = new FileStream(thumbnailPath, FileMode.Create))
                    {
                        encoder.Save(outputStream);
                    }
                }
                
                item.ThumbnailPath = thumbnailPath;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"生成缩略图失败: {ex.Message}");
                // 如果缩略图生成失败，使用默认图标
                item.ThumbnailPath = "pack://application:,,,/PresPio;component/Resources/image_thumbnail.png";
            }
        }

        private string GetExtensionFromContentType(string contentType)
        {
            switch (contentType.ToLower())
            {
                case "image/png": return ".png";
                case "image/jpeg": return ".jpg";
                case "image/gif": return ".gif";
                case "image/bmp": return ".bmp";
                case "image/tiff": return ".tiff";
                case "video/mp4": return ".mp4";
                case "video/avi": return ".avi";
                case "video/wmv": return ".wmv";
                case "audio/mp3": return ".mp3";
                case "audio/wav": return ".wav";
                case "audio/wma": return ".wma";
                default: return ".bin";
            }
        }

        private bool IsImageExtension(string extension)
        {
            string[] validExtensions = { ".png", ".jpg", ".jpeg", ".gif", ".bmp", ".tiff" };
            return validExtensions.Contains(extension.ToLower());
        }

        private async Task UpdateFileInfo(MaterialItem item)
        {
            if (File.Exists(item.FilePath))
            {
                var fileInfo = new FileInfo(item.FilePath);
                item.Size = FormatFileSize(fileInfo.Length);
                item.CreateTime = fileInfo.CreationTime;
            }
        }

        private string FormatFileSize(long bytes)
        {
            string[] sizes = { "B", "KB", "MB", "GB" };
            int order = 0;
            double size = bytes;
            while (size >= 1024 && order < sizes.Length - 1)
            {
                order++;
                size = size / 1024;
            }
            return $"{size:0.##} {sizes[order]}";
        }

        private async void ExportMaterials()
        {
            try
            {
                Mouse.OverrideCursor = System.Windows.Input.Cursors.Wait;

                var selectedItems = FilteredMaterials.Where(x => x.IsSelected).ToList();
                if (!selectedItems.Any())
                {
                    MessageBox.Show("请至少选择一个要导出的素材", "提示", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                int successCount = 0;
                int failCount = 0;

                foreach (var item in selectedItems)
                {
                    try
                    {
                        if (File.Exists(item.FilePath))
                        {
                            string targetPath = Path.Combine(txtExportPath.Text, Path.GetFileName(item.FilePath));
                            // 如果目标文件已存在，添加数字后缀
                            int counter = 1;
                            string fileNameWithoutExt = Path.GetFileNameWithoutExtension(targetPath);
                            string extension = Path.GetExtension(targetPath);
                            while (File.Exists(targetPath))
                            {
                                targetPath = Path.Combine(txtExportPath.Text, $"{fileNameWithoutExt}_{counter}{extension}");
                                counter++;
                            }

                            File.Copy(item.FilePath, targetPath);
                            successCount++;
                        }
                        else
                        {
                            failCount++;
                        }
                    }
                    catch
                    {
                        failCount++;
                    }
                }

                MessageBox.Show($"导出完成\n成功：{successCount}个\n失败：{failCount}个", "导出结果", MessageBoxButton.OK, 
                    failCount == 0 ? MessageBoxImage.Information : MessageBoxImage.Warning);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"导出过程中出错：{ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                Mouse.OverrideCursor = null;
            }
        }

        private void CmbExportType_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            ApplyFilter();
        }

        private void ApplyFilter()
        {
            if (Materials == null || cmbExportType == null) return;

            FilteredMaterials.Clear();
            var selectedItem = cmbExportType.SelectedItem as ComboBoxItem;
            if (selectedItem == null) return;

            var materials = Materials.ToList();
            switch (selectedItem.Content.ToString())
            {
                case "所有素材":
                    materials.ForEach(m => FilteredMaterials.Add(m));
                    break;
                case "仅图片":
                    materials.Where(m => m.Type == "图片").ToList().ForEach(m => FilteredMaterials.Add(m));
                    break;
                case "仅视频":
                    materials.Where(m => m.Type == "视频").ToList().ForEach(m => FilteredMaterials.Add(m));
                    break;
                case "仅音频":
                    materials.Where(m => m.Type == "音频").ToList().ForEach(m => FilteredMaterials.Add(m));
                    break;
            }
        }

        private void BtnSelectAll_Click(object sender, RoutedEventArgs e)
        {
            foreach (var item in Materials)
            {
                item.IsSelected = true;
            }
        }

        private void BtnUnselectAll_Click(object sender, RoutedEventArgs e)
        {
            foreach (var item in Materials)
            {
                item.IsSelected = false;
            }
        }

        private void BtnBrowse_Click(object sender, RoutedEventArgs e)
        {
            using (var dialog = new FolderBrowserDialog())
            {
                if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    txtExportPath.Text = dialog.SelectedPath;
                }
            }
        }

        private void BtnExport_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrEmpty(txtExportPath.Text))
            {
                MessageBox.Show("请选择导出位置", "提示", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            if (!Directory.Exists(txtExportPath.Text))
            {
                try
                {
                    Directory.CreateDirectory(txtExportPath.Text);
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"创建导出目录失败：{ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
                    return;
                }
            }

            ExportMaterials();
        }

        private void BtnCancel_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }

        protected override void OnClosed(EventArgs e)
        {
            base.OnClosed(e);
            
            // 清理COM对象
            if (activePresentation != null)
            {
                try
                {
                    Marshal.ReleaseComObject(activePresentation);
                }
                catch { }
                activePresentation = null;
            }

            // 只有在不是外部传入的情况下才释放PowerPoint应用程序对象
            if (!isExternalPptApp && pptApplication != null)
            {
                try
                {
                    Marshal.ReleaseComObject(pptApplication);
                }
                catch { }
                pptApplication = null;
            }

            // 清理临时文件夹
            try
            {
                if (Directory.Exists(tempFolder))
                {
                    Directory.Delete(tempFolder, true);
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"清理临时文件夹失败: {ex.Message}");
            }
        }

        private string GetExtensionFromName(string name)
        {
            if (name.EndsWith(".mp4")) return ".mp4";
            if (name.EndsWith(".avi")) return ".avi";
            if (name.EndsWith(".wmv")) return ".wmv";
            if (name.EndsWith(".mov")) return ".mov";
            if (name.EndsWith(".mp3")) return ".mp3";
            if (name.EndsWith(".wav")) return ".wav";
            if (name.EndsWith(".wma")) return ".wma";
            return ".mp4"; // 默认扩展名
        }
    }

    public class MaterialItem : INotifyPropertyChanged
    {
        private bool _isSelected;
        public bool IsSelected
        {
            get => _isSelected;
            set
            {
                _isSelected = value;
                OnPropertyChanged(nameof(IsSelected));
            }
        }

        private string _size;
        public string Size
        {
            get => _size;
            set
            {
                _size = value;
                OnPropertyChanged(nameof(Size));
            }
        }

        private string _thumbnailPath;
        public string ThumbnailPath
        {
            get => _thumbnailPath;
            set
            {
                _thumbnailPath = value;
                OnPropertyChanged(nameof(ThumbnailPath));
            }
        }

        private string _name;
        public string Name
        {
            get => _name;
            set
            {
                _name = value;
                OnPropertyChanged(nameof(Name));
            }
        }

        public string Type { get; set; }
        public DateTime CreateTime { get; set; }
        public string FilePath { get; set; }
        public int SlideNumber { get; set; }

        public event PropertyChangedEventHandler PropertyChanged;

        protected virtual void OnPropertyChanged(string propertyName)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }
}
