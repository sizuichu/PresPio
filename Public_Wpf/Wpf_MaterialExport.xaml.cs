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
using System.Collections.Generic;

namespace PresPio.Public_Wpf
{
    public partial class Wpf_MaterialExport : Window
    {
        private ObservableCollection<MaterialItem> Materials { get; set; }
        private ObservableCollection<MaterialItem> FilteredMaterials { get; set; }
        private readonly PowerPoint.Application pptApplication;
        private PowerPoint.Presentation activePresentation;
        private string tempFolder;
        private bool isExternalPptApp = true;
        private bool isClosing = false;

        public Wpf_MaterialExport(PowerPoint.Application app)
        {
            if (app == null)
            {
                throw new ArgumentNullException(nameof(app), "PowerPoint应用程序对象不能为空");
            }

            try
            {
                // 验证PowerPoint应用程序是否可用
                var test = app.ActivePresentation;
                pptApplication = app;
                isExternalPptApp = true;
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException("PowerPoint应用程序对象无效或已关闭", ex);
            }

            InitializeComponent();

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
            cmbPageRange.SelectionChanged += CmbPageRange_SelectionChanged;
            txtCustomPages.TextChanged += TxtCustomPages_TextChanged;
            lvMaterials.SelectionChanged += LvMaterials_SelectionChanged;

            // 媒体控件事件
            btnPlay.Click += BtnPlay_Click;
            btnPause.Click += BtnPause_Click;
            btnStop.Click += BtnStop_Click;
            mediaPreview.MediaEnded += MediaPreview_MediaEnded;

            // 设置默认选项
            cmbExportType.SelectedIndex = 0;
            cmbPageRange.SelectedIndex = 0;

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
            if (isClosing) return;

            try
            {
                Mouse.OverrideCursor = System.Windows.Input.Cursors.Wait;
                ShowProgress(true, "正在加载素材...", 0);

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

                        // 检查是否应该包含当前幻灯片
                        if (!ShouldIncludeSlide(slide.SlideNumber))
                        {
                            continue;
                        }

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
                                        
                                        // 尝试获取原始文件路径
                                        string originalPath = "";
                                        try
                                        {
                                            if (shape.LinkFormat != null)
                                            {
                                                originalPath = shape.LinkFormat.SourceFullName;
                                            }
                                        }
                                        catch { }

                                        // 如果有原始文件，直接复制
                                        if (!string.IsNullOrEmpty(originalPath) && File.Exists(originalPath))
                                        {
                                            File.Copy(originalPath, tempFilePath, true);
                                        }
                                        else // 否则导出图片
                                        {
                                            shape.Export(tempFilePath, PowerPoint.PpShapeFormat.ppShapeFormatPNG);
                                        }
                                        
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
                                // 处理媒体（视频和音频）
                                else if (shape.Type == MsoShapeType.msoMedia || shape.Type == MsoShapeType.msoLinkedOLEObject || 
                                        shape.Type == MsoShapeType.msoEmbeddedOLEObject)
                                {
                                    string mediaType = "未知";
                                    string extension = ".mp4"; // 默认扩展名
                                    string originalPath = "";

                                    try
                                    {
                                        // 尝试从不同属性获取媒体信息
                                        if (shape.OLEFormat != null)
                                        {
                                            string progID = shape.OLEFormat.ProgID ?? "";
                                            if (progID.Contains("Video") || progID.Contains("Media"))
                                            {
                                                mediaType = "视频";
                                            }
                                            else if (progID.Contains("Audio") || progID.Contains("Sound"))
                                            {
                                                mediaType = "音频";
                                            }
                                        }

                                        // 尝试获取原始文件路径
                                        try
                                        {
                                            if (shape.LinkFormat != null)
                                            {
                                                originalPath = shape.LinkFormat.SourceFullName;
                                            }
                                        }
                                        catch { }

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

                                        if (mediaType != "未知")
                                        {
                                            string fileName = $"{mediaType}_{Guid.NewGuid()}{extension}";
                                            string tempFilePath = Path.Combine(tempFolder, fileName);

                                            // 如果有原始文件，直接复制
                                            if (!string.IsNullOrEmpty(originalPath) && File.Exists(originalPath))
                                            {
                                                File.Copy(originalPath, tempFilePath, true);
                                            }
                                            else // 尝试导出
                                            {
                                                try
                                                {
                                                    shape.Export(tempFilePath, PowerPoint.PpShapeFormat.ppShapeFormatPNG);
                                                }
                                                catch
                                                {
                                                    continue;
                                                }
                                            }

                                            if (File.Exists(tempFilePath))
                                            {
                                                item = new MaterialItem
                                                {
                                                    Name = string.IsNullOrEmpty(shape.Name) ? fileName : shape.Name,
                                                    Type = mediaType,
                                                    FilePath = tempFilePath,
                                                    CreateTime = DateTime.Now,
                                                    SlideNumber = slide.SlideNumber
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
                                        }
                                    }
                                    catch (Exception ex)
                                    {
                                        System.Diagnostics.Debug.WriteLine($"处理媒体失败: {ex.Message}");
                                        continue;
                                    }
                                }

                                if (item != null)
                                {
                                    Materials.Add(item);
                                    await UpdateFileInfo(item);
                                }
                            }
                            catch (Exception ex)
                            {
                                System.Diagnostics.Debug.WriteLine($"处理形状失败: {ex.Message}");
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
                    if (!isClosing)
                    {
                        MessageBox.Show("无法访问PowerPoint演示文稿，请确保文件未被锁定且PowerPoint运行正常。", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
                    }
                    return;
                }
                catch (Exception ex) when (ex is InvalidCastException || ex is System.Runtime.InteropServices.InvalidComObjectException)
                {
                    if (!isClosing)
                    {
                        MessageBox.Show("PowerPoint应用程序状态异常，请重新打开PowerPoint。", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
                    }
                    return;
                }
            }
            catch (Exception ex)
            {
                if (!isClosing)
                {
                    MessageBox.Show($"加载素材时出错：{ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
            finally
            {
                if (!isClosing)
                {
                    ShowProgress(false);
                    Mouse.OverrideCursor = null;
                }
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
            isClosing = true;
            this.Close();
        }

        protected override void OnClosing(CancelEventArgs e)
        {
            isClosing = true;
            base.OnClosing(e);
        }

        protected override void OnClosed(EventArgs e)
        {
            base.OnClosed(e);
            
            // 清理COM对象
            if (activePresentation != null)
            {
                try
                {
                    Marshal.FinalReleaseComObject(activePresentation);
                }
                catch { }
                activePresentation = null;
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

        private void CmbPageRange_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (cmbPageRange.SelectedItem is ComboBoxItem selectedItem)
            {
                txtCustomPages.Visibility = selectedItem.Content.ToString() == "自定义页面" ? 
                    Visibility.Visible : Visibility.Collapsed;
                LoadMaterials();
            }
        }

        private void TxtCustomPages_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (cmbPageRange.SelectedItem is ComboBoxItem selectedItem && 
                selectedItem.Content.ToString() == "自定义页面")
            {
                LoadMaterials();
            }
        }

        private HashSet<int> ParseCustomPages(string input)
        {
            var result = new HashSet<int>();
            if (string.IsNullOrWhiteSpace(input)) return result;

            try
            {
                var parts = input.Split(',');
                foreach (var part in parts)
                {
                    if (part.Contains("-"))
                    {
                        var range = part.Split('-');
                        if (range.Length == 2 && int.TryParse(range[0], out int start) && 
                            int.TryParse(range[1], out int end))
                        {
                            for (int i = start; i <= end; i++)
                            {
                                result.Add(i);
                            }
                        }
                    }
                    else if (int.TryParse(part, out int pageNumber))
                    {
                        result.Add(pageNumber);
                    }
                }
            }
            catch
            {
                // 忽略解析错误
            }

            return result;
        }

        private void LvMaterials_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (lvMaterials.SelectedItem is MaterialItem selectedItem)
            {
                UpdatePreview(selectedItem);
            }
        }

        private void UpdatePreview(MaterialItem item)
        {
            if (item == null) return;

            // 更新预览标题
            txtPreviewTitle.Text = item.Name;

            // 更新详细信息
            txtPreviewName.Text = item.Name;
            txtPreviewType.Text = item.Type;
            txtPreviewSize.Text = item.Size;
            txtPreviewPath.Text = item.FilePath;
            txtPreviewCreateTime.Text = item.CreateTime.ToString("yyyy-MM-dd HH:mm:ss");
            txtPreviewSlide.Text = $"第 {item.SlideNumber} 页";

            // 停止当前媒体播放
            mediaPreview.Stop();
            mediaPreview.Source = null;

            // 根据类型显示不同的预览
            switch (item.Type)
            {
                case "图片":
                    imgPreview.Source = new BitmapImage(new Uri(item.FilePath));
                    imgPreview.Visibility = Visibility.Visible;
                    mediaPreview.Visibility = Visibility.Collapsed;
                    mediaControls.Visibility = Visibility.Collapsed;
                    break;

                case "视频":
                    mediaPreview.Source = new Uri(item.FilePath);
                    mediaPreview.Visibility = Visibility.Visible;
                    mediaControls.Visibility = Visibility.Visible;
                    imgPreview.Visibility = Visibility.Collapsed;
                    break;

                case "音频":
                    mediaPreview.Source = new Uri(item.FilePath);
                    mediaPreview.Visibility = Visibility.Collapsed;
                    mediaControls.Visibility = Visibility.Visible;
                    imgPreview.Visibility = Visibility.Visible;
                    imgPreview.Source = new BitmapImage(new Uri(item.ThumbnailPath));
                    break;
            }
        }

        private void BtnPlay_Click(object sender, RoutedEventArgs e)
        {
            mediaPreview.Play();
        }

        private void BtnPause_Click(object sender, RoutedEventArgs e)
        {
            mediaPreview.Pause();
        }

        private void BtnStop_Click(object sender, RoutedEventArgs e)
        {
            mediaPreview.Stop();
        }

        private void MediaPreview_MediaEnded(object sender, RoutedEventArgs e)
        {
            mediaPreview.Position = TimeSpan.Zero;
            mediaPreview.Play();
        }

        private bool ShouldIncludeSlide(int slideNumber)
        {
            if (cmbPageRange.SelectedItem is ComboBoxItem selectedItem)
            {
                switch (selectedItem.Content.ToString())
                {
                    case "当前页面":
                        return slideNumber == pptApplication.ActiveWindow.Selection.SlideRange.SlideNumber;
                    case "所有页面":
                        return true;
                    case "自定义页面":
                        var customPages = ParseCustomPages(txtCustomPages.Text);
                        return customPages.Contains(slideNumber);
                    default:
                        return true;
                }
            }
            return true;
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
