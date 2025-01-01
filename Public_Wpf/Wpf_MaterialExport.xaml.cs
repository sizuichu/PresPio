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
        private string tempPptxPath;

        public Wpf_MaterialExport(PowerPoint.Application app)
        {
            if (app == null)
            {
                throw new ArgumentNullException(nameof(app), "PowerPoint应用程序对象不能为空");
            }

            try
            {
                // 检查PowerPoint应用程序状态
                var presentations = app.Presentations;
                if (presentations == null)
                {
                    throw new InvalidOperationException("PowerPoint应用程序对象无效");
                }

                // 检查是否有活动的演示文稿
                var presentation = app.ActivePresentation;
                if (presentation == null)
                {
                    throw new InvalidOperationException("请先打开一个PowerPoint演示文稿");
                }

                // 检查演示文稿是否已保存
                string pptPath = presentation.FullName;
                if (string.IsNullOrEmpty(pptPath))
                {
                    throw new InvalidOperationException("请先保存PowerPoint演示文稿");
                }

                // 创建临时文件夹
                tempFolder = Path.Combine(Path.GetTempPath(), "PresPio_Export_" + Guid.NewGuid().ToString());
                Directory.CreateDirectory(tempFolder);
                Directory.CreateDirectory(Path.Combine(tempFolder, "thumbnails")); // 创建缩略图子文件夹

                // 将当前PPT保存为临时文件
                tempPptxPath = Path.Combine(tempFolder, "temp.pptx");
                presentation.SaveCopyAs(tempPptxPath);

                // 所有检查通过后，再赋值
                pptApplication = app;
                isExternalPptApp = true;

                // 初始化UI组件
                InitializeComponent();

                // 设置默认导出路径
                txtExportPath.Text = Path.Combine(Path.GetDirectoryName(pptPath), "导出的素材");

                // 初始化其他控件
                InitializeControls();
            }
            catch (COMException comEx)
            {
                CleanupTempFiles();
                throw new InvalidOperationException("PowerPoint应用程序通信错误", comEx);
            }
            catch (Exception ex) when (!(ex is InvalidOperationException))
            {
                CleanupTempFiles();
                throw new InvalidOperationException("PowerPoint应用程序状态异常", ex);
            }
        }

        private void CleanupTempFiles()
        {
            try
            {
                if (Directory.Exists(tempFolder))
                {
                    Directory.Delete(tempFolder, true);
                }
            }
            catch { }
        }

        private void InitializeControls()
        {
            try
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
            catch (Exception ex)
            {
                throw new InvalidOperationException("初始化控件失败", ex);
            }
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
                    // 使用临时保存的PPT文件
                    if (!File.Exists(tempPptxPath))
                    {
                        MessageBox.Show("临时文件不存在，请重试", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
                        return;
                    }

                    Materials.Clear();

                    using (PresentationDocument presentationDocument = PresentationDocument.Open(tempPptxPath, false))
                    {
                        var presentationPart = presentationDocument.PresentationPart;
                        if (presentationPart != null)
                        {
                            // 获取当前幻灯片编号（如果需要）
                            int currentSlideNumber = 1;
                            try
                            {
                                currentSlideNumber = pptApplication.ActiveWindow.Selection.SlideRange.SlideNumber;
                            }
                            catch { }

                            // 获取所有幻灯片
                            var slideIds = presentationPart.Presentation.SlideIdList?.ChildElements
                                .OfType<SlideId>().ToList() ?? new List<SlideId>();

                            int totalSlides = slideIds.Count;
                            int currentSlide = 0;

                            foreach (var slideId in slideIds)
                            {
                                currentSlide++;
                                ShowProgress(true, $"正在处理幻灯片 ({currentSlide}/{totalSlides})", 
                                    (currentSlide / (double)totalSlides) * 100);

                                var slidePart = presentationPart.GetPartById(slideId.RelationshipId) as SlidePart;
                                if (slidePart == null) continue;

                                int slideNumber = slideIds.IndexOf(slideId) + 1;

                                // 检查是否应该包含当前幻灯片
                                if (!ShouldIncludeSlide(slideNumber, currentSlideNumber))
                                {
                                    continue;
                                }

                                // 处理所有图片部件
                                foreach (var part in slidePart.Parts)
                                {
                                    if (part.OpenXmlPart is ImagePart imagePart)
                                    {
                                        try
                                        {
                                            string contentType = imagePart.ContentType;
                                            string extension = GetExtensionFromContentType(contentType);
                                            
                                            if (IsImageExtension(extension))
                                            {
                                                string fileName = $"image_{Guid.NewGuid()}{extension}";
                                                string tempFilePath = Path.Combine(tempFolder, fileName);

                                                using (Stream stream = imagePart.GetStream())
                                                using (FileStream fileStream = new FileStream(tempFilePath, FileMode.Create))
                                                {
                                                    stream.CopyTo(fileStream);
                                                }

                                                // 检查是否为链接图片
                                                bool isLinked = false;
                                                string pictureName = fileName;

                                                // 查找对应的Picture元素
                                                var pictures = slidePart.Slide.Descendants<DocumentFormat.OpenXml.Presentation.Picture>();
                                                foreach (var picture in pictures)
                                                {
                                                    var blipFill = picture.BlipFill;
                                                    if (blipFill?.Blip?.Embed?.Value == part.RelationshipId)
                                                    {
                                                        isLinked = true;
                                                        // 尝试获取图片名称
                                                        var nvPicPr = picture.NonVisualPictureProperties;
                                                        if (nvPicPr?.NonVisualDrawingProperties?.Name != null)
                                                        {
                                                            pictureName = nvPicPr.NonVisualDrawingProperties.Name;
                                                        }
                                                        break;
                                                    }
                                                }

                                                var item = new MaterialItem
                                                {
                                                    Name = pictureName,
                                                    Type = isLinked ? "图片(链接)" : "图片(嵌入)",
                                                    FilePath = tempFilePath,
                                                    CreateTime = DateTime.Now,
                                                    SlideNumber = slideNumber
                                                };

                                                await GenerateThumbnail(item);
                                                Materials.Add(item);
                                                await UpdateFileInfo(item);
                                            }
                                        }
                                        catch (Exception ex)
                                        {
                                            System.Diagnostics.Debug.WriteLine($"处理图片失败: {ex.Message}");
                                        }
                                    }
                                }

                                // 处理媒体文件
                                foreach (var part in slidePart.Parts)
                                {
                                    var mediaContentType = part.OpenXmlPart.ContentType.ToLower();
                                    if (mediaContentType.Contains("video/") || mediaContentType.Contains("audio/"))
                                    {
                                        try
                                        {
                                            string mediaType = mediaContentType.Contains("video/") ? "视频" : "音频";
                                            string extension = GetExtensionFromContentType(mediaContentType);
                                            string fileName = $"{mediaType}_{Guid.NewGuid()}{extension}";
                                            string tempFilePath = Path.Combine(tempFolder, fileName);

                                            using (Stream stream = part.OpenXmlPart.GetStream())
                                            using (FileStream fileStream = new FileStream(tempFilePath, FileMode.Create))
                                            {
                                                stream.CopyTo(fileStream);
                                            }

                                            var item = new MaterialItem
                                            {
                                                Name = fileName,
                                                Type = mediaType,
                                                FilePath = tempFilePath,
                                                CreateTime = DateTime.Now,
                                                SlideNumber = slideNumber
                                            };

                                            // 设置默认缩略图
                                            item.ThumbnailPath = mediaType == "视频" ?
                                                "pack://application:,,,/PresPio;component/Resources/video_thumbnail.png" :
                                                "pack://application:,,,/PresPio;component/Resources/audio_thumbnail.png";

                                            Materials.Add(item);
                                            await UpdateFileInfo(item);
                                        }
                                        catch (Exception ex)
                                        {
                                            System.Diagnostics.Debug.WriteLine($"处理媒体文件失败: {ex.Message}");
                                        }
                                    }
                                }
                            }
                        }
                    }

                    ApplyFilter();

                    if (Materials.Count == 0)
                    {
                        MessageBox.Show("未找到任何可导出的素材", "提示", MessageBoxButton.OK, MessageBoxImage.Information);
                    }
                }
                catch (Exception ex)
                {
                    if (!isClosing)
                    {
                        MessageBox.Show($"处理演示文稿时出错：{ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
                    }
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
                Directory.CreateDirectory(thumbnailsFolder); // 确保文件夹存在
                string thumbnailPath = Path.Combine(thumbnailsFolder, $"thumb_{Path.GetFileName(item.FilePath)}");
                
                if (item.Type.StartsWith("图片"))
                {
                    using (var fileStream = new FileStream(item.FilePath, FileMode.Open, FileAccess.Read))
                    {
                        var decoder = BitmapDecoder.Create(
                            fileStream,
                            BitmapCreateOptions.PreservePixelFormat,
                            BitmapCacheOption.OnLoad);

                        if (decoder.Frames[0] != null)
                        {
                            // 计算缩略图尺寸
                            double scale = Math.Min(120.0 / decoder.Frames[0].PixelWidth,
                                                  120.0 / decoder.Frames[0].PixelHeight);
                            int width = (int)(decoder.Frames[0].PixelWidth * scale);
                            int height = (int)(decoder.Frames[0].PixelHeight * scale);

                            var thumbnail = new TransformedBitmap(
                                decoder.Frames[0],
                                new ScaleTransform(scale, scale));

                            using (var thumbnailStream = new FileStream(thumbnailPath, FileMode.Create))
                            {
                                var encoder = new PngBitmapEncoder();
                                encoder.Frames.Add(BitmapFrame.Create(thumbnail));
                                encoder.Save(thumbnailStream);
                            }

                            item.ThumbnailPath = thumbnailPath;
                        }
                    }
                }
                else if (item.Type == "视频")
                {
                    item.ThumbnailPath = "pack://application:,,,/PresPio;component/Resources/video_thumbnail.png";
                }
                else if (item.Type == "音频")
                {
                    item.ThumbnailPath = "pack://application:,,,/PresPio;component/Resources/audio_thumbnail.png";
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"生成缩略图失败: {ex.Message}");
                // 设置默认缩略图
                if (item.Type.StartsWith("图片"))
                {
                    item.ThumbnailPath = "pack://application:,,,/PresPio;component/Resources/image_thumbnail.png";
                }
                else if (item.Type == "视频")
                {
                    item.ThumbnailPath = "pack://application:,,,/PresPio;component/Resources/video_thumbnail.png";
                }
                else if (item.Type == "音频")
                {
                    item.ThumbnailPath = "pack://application:,,,/PresPio;component/Resources/audio_thumbnail.png";
                }
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
            try
            {
                if (File.Exists(item.FilePath))
                {
                    var fileInfo = new FileInfo(item.FilePath);
                    double sizeInKB = fileInfo.Length / 1024.0;
                    item.Size = sizeInKB >= 1024 ? 
                        $"{sizeInKB / 1024:F2} MB" : 
                        $"{sizeInKB:F2} KB";
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"更新文件信息失败: {ex.Message}");
                item.Size = "未知";
            }
        }

        private async void ExportMaterials()
        {
            try
            {
                Mouse.OverrideCursor = System.Windows.Input.Cursors.Wait;
                ShowProgress(true, "正在导出素材...", 0);

                var selectedItems = FilteredMaterials.Where(x => x.IsSelected).ToList();
                if (!selectedItems.Any())
                {
                    MessageBox.Show("请至少选择一个要导出的素材", "提示", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                int totalCount = selectedItems.Count;
                int currentCount = 0;
                int successCount = 0;
                int failCount = 0;

                foreach (var item in selectedItems)
                {
                    currentCount++;
                    ShowProgress(true, $"正在导出 ({currentCount}/{totalCount})", 
                        (currentCount / (double)totalCount) * 100);

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
                ShowProgress(false);
                Mouse.OverrideCursor = null;
            }
        }

        private void CmbExportType_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            ApplyFilter();
        }

        private void ApplyFilter()
        {
            if (FilteredMaterials == null) return;

            FilteredMaterials.Clear();
            var selectedType = (cmbExportType.SelectedItem as ComboBoxItem)?.Content.ToString();

            var filteredItems = Materials.Where(item =>
            {
                if (selectedType == "全部")
                    return true;
                else if (selectedType == "图片")
                    return item.Type.StartsWith("图片");
                else
                    return item.Type == selectedType;
            });

            // 应用排序
            filteredItems = ApplySorting(filteredItems);

            foreach (var item in filteredItems)
            {
                FilteredMaterials.Add(item);
            }
        }

        private IEnumerable<MaterialItem> ApplySorting(IEnumerable<MaterialItem> items)
        {
            // 默认按页码和类型排序
            return items.OrderBy(item => item.SlideNumber)
                       .ThenBy(item => GetTypeOrder(item.Type))
                       .ThenBy(item => item.Name);
        }

        private int GetTypeOrder(string type)
        {
            switch (type)
            {
                case "图片(链接)": return 1;
                case "图片(嵌入)": return 2;
                case "视频": return 3;
                case "音频": return 4;
                default: return 99;
            }
        }

        private void BtnSelectAll_Click(object sender, RoutedEventArgs e)
        {
            foreach (var item in FilteredMaterials)
            {
                item.IsSelected = true;
            }
        }

        private void BtnUnselectAll_Click(object sender, RoutedEventArgs e)
        {
            foreach (var item in FilteredMaterials)
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
            if (string.IsNullOrWhiteSpace(txtExportPath.Text))
            {
                MessageBox.Show("请选择导出位置", "提示", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            try
            {
                // 确保导出目录存在
                if (!Directory.Exists(txtExportPath.Text))
                {
                    Directory.CreateDirectory(txtExportPath.Text);
                }

                ExportMaterials();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"创建导出目录失败：{ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        protected override void OnClosing(CancelEventArgs e)
        {
            isClosing = true;
            base.OnClosing(e);
        }

        protected override void OnClosed(EventArgs e)
        {
            try
            {
                isClosing = true;
                mediaPreview.Stop();
                mediaPreview.Close();

                // 清理临时文件
                CleanupTempFiles();

                // 释放COM对象
                if (pptApplication != null && !isExternalPptApp)
                {
                    Marshal.ReleaseComObject(pptApplication);
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"关闭窗口时出错: {ex.Message}");
            }

            base.OnClosed(e);
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

            try
            {
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
                    case "图片(链接)":
                    case "图片(嵌入)":
                        try
                        {
                            if (File.Exists(item.FilePath))
                            {
                                var bitmap = new BitmapImage();
                                bitmap.BeginInit();
                                bitmap.CacheOption = BitmapCacheOption.OnLoad;
                                bitmap.CreateOptions = BitmapCreateOptions.IgnoreImageCache;
                                bitmap.UriSource = new Uri(item.FilePath);
                                bitmap.EndInit();
                                bitmap.Freeze();
                                imgPreview.Source = bitmap;
                            }
                            else
                            {
                                imgPreview.Source = new BitmapImage(new Uri("pack://application:,,,/PresPio;component/Resources/image_thumbnail.png", UriKind.Absolute));
                            }
                        }
                        catch (Exception ex)
                        {
                            System.Diagnostics.Debug.WriteLine($"加载图片预览失败: {ex.Message}");
                            imgPreview.Source = new BitmapImage(new Uri("pack://application:,,,/PresPio;component/Resources/image_thumbnail.png", UriKind.Absolute));
                        }
                        imgPreview.Visibility = Visibility.Visible;
                        mediaPreview.Visibility = Visibility.Collapsed;
                        mediaControls.Visibility = Visibility.Collapsed;
                        break;

                    case "视频":
                        try
                        {
                            if (File.Exists(item.FilePath))
                            {
                                mediaPreview.Source = new Uri(item.FilePath);
                                mediaPreview.Visibility = Visibility.Visible;
                                mediaControls.Visibility = Visibility.Visible;
                                imgPreview.Visibility = Visibility.Collapsed;
                            }
                            else
                            {
                                imgPreview.Source = new BitmapImage(new Uri("pack://application:,,,/PresPio;component/Resources/video_thumbnail.png", UriKind.Absolute));
                                imgPreview.Visibility = Visibility.Visible;
                                mediaPreview.Visibility = Visibility.Collapsed;
                                mediaControls.Visibility = Visibility.Collapsed;
                            }
                        }
                        catch (Exception ex)
                        {
                            System.Diagnostics.Debug.WriteLine($"加载视频预览失败: {ex.Message}");
                            imgPreview.Source = new BitmapImage(new Uri("pack://application:,,,/PresPio;component/Resources/video_thumbnail.png", UriKind.Absolute));
                            imgPreview.Visibility = Visibility.Visible;
                            mediaPreview.Visibility = Visibility.Collapsed;
                            mediaControls.Visibility = Visibility.Collapsed;
                        }
                        break;

                    case "音频":
                        try
                        {
                            if (File.Exists(item.FilePath))
                            {
                                mediaPreview.Source = new Uri(item.FilePath);
                                mediaPreview.Visibility = Visibility.Collapsed;
                                mediaControls.Visibility = Visibility.Visible;
                            }
                            else
                            {
                                mediaControls.Visibility = Visibility.Collapsed;
                            }
                            imgPreview.Source = new BitmapImage(new Uri("pack://application:,,,/PresPio;component/Resources/audio_thumbnail.png", UriKind.Absolute));
                            imgPreview.Visibility = Visibility.Visible;
                        }
                        catch (Exception ex)
                        {
                            System.Diagnostics.Debug.WriteLine($"加载音频预览失败: {ex.Message}");
                            imgPreview.Source = new BitmapImage(new Uri("pack://application:,,,/PresPio;component/Resources/audio_thumbnail.png", UriKind.Absolute));
                            imgPreview.Visibility = Visibility.Visible;
                            mediaPreview.Visibility = Visibility.Collapsed;
                            mediaControls.Visibility = Visibility.Collapsed;
                        }
                        break;
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"更新预览失败: {ex.Message}");
                // 显示默认图标
                string defaultIcon = "pack://application:,,,/PresPio;component/Resources/image_thumbnail.png";
                if (item.Type == "视频")
                    defaultIcon = "pack://application:,,,/PresPio;component/Resources/video_thumbnail.png";
                else if (item.Type == "音频")
                    defaultIcon = "pack://application:,,,/PresPio;component/Resources/audio_thumbnail.png";

                imgPreview.Source = new BitmapImage(new Uri(defaultIcon, UriKind.Absolute));
                imgPreview.Visibility = Visibility.Visible;
                mediaPreview.Visibility = Visibility.Collapsed;
                mediaControls.Visibility = Visibility.Collapsed;
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

        private bool ShouldIncludeSlide(int slideNumber, int currentSlideNumber)
        {
            if (cmbPageRange.SelectedItem is ComboBoxItem selectedItem)
            {
                switch (selectedItem.Content.ToString())
                {
                    case "当前页面":
                        return slideNumber == currentSlideNumber;
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
