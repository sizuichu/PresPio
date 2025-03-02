using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using LiteDB;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using Newtonsoft.Json;
using Application = System.Windows.Application;

namespace PresPio
    {
    // 添加顶级静态类用于扩展方法
    public static class StringExtensions
        {
        public static byte[] ToBytes(this string str)
            {
            return System.Text.Encoding.UTF8.GetBytes(str);
            }
        }

    public partial class Page_textPods : UserControl
        {
        private readonly string JsonFilePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "ClipboardData.json");
        private ObservableCollection<ClipboardItem> ClipboardItems { get; set; }
        private ObservableCollection<ClipboardItem> FilteredClipboardItems { get; set; }
        private string currentFilter = "All";

        public Page_textPods()
            {
            // 初始化集合
            ClipboardItems = new ObservableCollection<ClipboardItem>();
            FilteredClipboardItems = new ObservableCollection<ClipboardItem>();

            InitializeComponent();

            // 设置数据上下文
            this.DataContext = this;

            // 确保控件已初始化后再设置数据源
            if (ClipboardItemsControl != null)
                {
                ClipboardItemsControl.ItemsSource = FilteredClipboardItems;
                }

            // 加载数据
            Loaded += Page_textPods_Loaded;
            }

        private void Page_textPods_Loaded(object sender, RoutedEventArgs e)
            {
            LoadClipboardItems();

            // 如果数据为空，添加测试数据
            if (!ClipboardItems.Any())
                {
                AddTestData();
                }

            // 初始化筛选为"全部"
            FilterItems("All");
            }

        private void LoadClipboardItems()
            {
            try
                {
                if (File.Exists(JsonFilePath))
                    {
                    var json = File.ReadAllText(JsonFilePath);
                    var items = JsonConvert.DeserializeObject<List<ClipboardItem>>(json);

                    Application.Current.Dispatcher.Invoke(() =>
                    {
                        ClipboardItems.Clear();
                        FilteredClipboardItems.Clear();

                        foreach (var item in items.OrderByDescending(x => x.CreatedTime))
                            {
                            ClipboardItems.Add(item);
                            FilteredClipboardItems.Add(item);
                            }
                    });
                    }
                }
            catch (Exception ex)
                {
                MessageBox.Show($"加载剪贴板项目失败: {ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }

        private void SaveClipboardItems()
            {
            try
                {
                var json = JsonConvert.SerializeObject(ClipboardItems.ToList(), Formatting.Indented);
                File.WriteAllText(JsonFilePath, json);
                }
            catch (Exception ex)
                {
                MessageBox.Show($"保存剪贴板项目失败: {ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }

        public void AddClipboardItem(ClipboardItem item)
            {
            try
                {
                Application.Current.Dispatcher.Invoke(() =>
                {
                    ClipboardItems.Insert(0, item);
                    if (ShouldShowItem(item))
                        {
                        FilteredClipboardItems.Insert(0, item);
                        }
                });

                SaveClipboardItems();
                }
            catch (Exception ex)
                {
                MessageBox.Show($"添加剪贴板项目失败: {ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }

        private bool ShouldShowItem(ClipboardItem item)
            {
            return currentFilter == "All" || item.Type == currentFilter;
            }

        public void DeleteClipboardItem(ClipboardItem item)
            {
            try
                {
                Application.Current.Dispatcher.Invoke(() =>
                {
                    ClipboardItems.Remove(item);
                    FilteredClipboardItems.Remove(item);
                });

                SaveClipboardItems();
                }
            catch (Exception ex)
                {
                MessageBox.Show($"删除剪贴板项目失败: {ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }

        public void ClearAllItems()
            {
            try
                {
                Application.Current.Dispatcher.Invoke(() =>
                {
                    ClipboardItems.Clear();
                    FilteredClipboardItems.Clear();
                });

                SaveClipboardItems();
                }
            catch (Exception ex)
                {
                MessageBox.Show($"清空剪贴板项目失败: {ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }

        private void Button_Clear_Click(object sender, RoutedEventArgs e)
            {
            if (MessageBox.Show("确定要清空所有内容吗？", "确认", MessageBoxButton.YesNo, MessageBoxImage.Question) == MessageBoxResult.Yes)
                {
                ClearAllItems();
                }
            }

        private void Button_Delete_Click(object sender, RoutedEventArgs e)
            {
            if (sender is Button button && button.DataContext is ClipboardItem item)
                {
                DeleteClipboardItem(item);
                }
            }

        private void Button_Insert_Click(object sender, RoutedEventArgs e)
            {
            if (sender is Button button && button.DataContext is ClipboardItem item)
                {
                try
                    {
                    var app = Globals.ThisAddIn.Application;
                    if (app == null)
                        {
                        MessageBox.Show("无法访问PowerPoint应用程序", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
                        return;
                        }

                    var presentation = app.ActivePresentation;
                    if (presentation == null)
                        {
                        MessageBox.Show("请先打开一个PowerPoint演示文稿", "提示", MessageBoxButton.OK, MessageBoxImage.Information);
                        return;
                        }

                    // 获取当前选中的幻灯片
                    var selection = app.ActiveWindow.Selection;
                    Slide currentSlide = null;

                    // 尝试从选中内容获取幻灯片
                    if (selection.Type == PpSelectionType.ppSelectionSlides)
                        {
                        currentSlide = selection.SlideRange[1];
                        }
                    else
                        {
                        // 如果没有选中幻灯片，则使用当前显示的幻灯片
                        currentSlide = app.ActiveWindow.View.Slide;
                        }

                    if (currentSlide == null)
                        {
                        MessageBox.Show("请先选择一个幻灯片", "提示", MessageBoxButton.OK, MessageBoxImage.Information);
                        return;
                        }

                    switch (item.Type)
                        {
                        case "Text":
                            InsertText(currentSlide, item);
                            break;

                        case "Image":
                            InsertImage(currentSlide, item);
                            break;

                        case "PPT":
                            InsertPPTContent(currentSlide, item);
                            break;

                        default:
                            MessageBox.Show($"不支持的内容类型：{item.Type}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
                            break;
                        }
                    }
                catch (Exception ex)
                    {
                    MessageBox.Show($"插入内容失败: {ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
                    }
                }
            }

        private void InsertText(Slide slide, ClipboardItem item)
            {
            try
                {
                float left = 100;
                float top = 100;
                float width = 400;
                float height = 100;

                var textBox = slide.Shapes.AddTextbox(
                    MsoTextOrientation.msoTextOrientationHorizontal,
                    left, top, width, height);

                string text = System.Text.Encoding.UTF8.GetString(item.ContentData);
                textBox.TextFrame.TextRange.Text = text;

                // 设置文本框样式
                textBox.TextFrame.AutoSize = PpAutoSize.ppAutoSizeShapeToFitText;
                textBox.TextFrame.WordWrap = MsoTriState.msoTrue;
                }
            catch (Exception ex)
                {
                MessageBox.Show($"插入文本失败: {ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }

        private void InsertImage(Slide slide, ClipboardItem item)
            {
            try
                {
                // 创建临时文件
                string tempImagePath = Path.Combine(Path.GetTempPath(), $"temp_image_{Guid.NewGuid()}.png");
                File.WriteAllBytes(tempImagePath, item.ContentData);

                try
                    {
                    float left = 100;
                    float top = 100;

                    var shape = slide.Shapes.AddPicture(
                        tempImagePath,
                        MsoTriState.msoFalse,
                        MsoTriState.msoTrue,
                        left, top);

                    // 调整图片大小，保持纵横比
                    float maxWidth = 400;
                    float maxHeight = 300;
                    float ratio = Math.Min(maxWidth / shape.Width, maxHeight / shape.Height);
                    shape.Width = shape.Width * ratio;
                    shape.Height = shape.Height * ratio;
                    }
                finally
                    {
                    // 清理临时文件
                    try { File.Delete(tempImagePath); } catch { }
                    }
                }
            catch (Exception ex)
                {
                MessageBox.Show($"插入图片失败: {ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }

        private void InsertPPTContent(Slide slide, ClipboardItem item)
            {
            try
                {
                // 从ContentData还原剪贴板数据
                var formatData = JsonConvert.DeserializeObject<Dictionary<string, byte[]>>(
                    System.Text.Encoding.UTF8.GetString(item.ContentData));

                // 创建新的DataObject
                var dataObject = new System.Windows.Forms.DataObject();
                foreach (var format in formatData)
                    {
                    dataObject.SetData(format.Key, format.Value);
                    }

                // 设置剪贴板
                System.Windows.Forms.Clipboard.SetDataObject(dataObject, true);

                // 粘贴内容
                slide.Shapes.Paste();

                // 恢复原始剪贴板内容
                var originalClipboard = System.Windows.Clipboard.GetDataObject();
                if (originalClipboard != null)
                    {
                    System.Windows.Forms.Clipboard.SetDataObject(originalClipboard, true);
                    }
                }
            catch (Exception ex)
                {
                MessageBox.Show($"插入PPT内容失败: {ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }

        private void ComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
            {
            if (sender is ComboBox comboBox && comboBox.SelectedItem is ComboBoxItem selectedItem)
                {
                currentFilter = selectedItem.Tag?.ToString() ?? "All";
                FilterItems(currentFilter);
                }
            }

        private void FilterItems(string filter)
            {
            if (FilteredClipboardItems == null || ClipboardItems == null) return;

            currentFilter = filter;
            Application.Current.Dispatcher.Invoke(() =>
            {
                FilteredClipboardItems.Clear();
                var filteredItems = filter == "All"
                    ? ClipboardItems
                    : ClipboardItems.Where(item => item.Type == filter);

                foreach (var item in filteredItems)
                    {
                    FilteredClipboardItems.Add(item);
                    }
            });
            }

        private void AddTestData()
            {
            try
                {
                // 添加文本测试数据
                AddClipboardItem(new ClipboardItem
                    {
                    Title = "重要会议记录",
                    Type = "Text",
                    Description = "2024年第一季度销售目标讨论会议记录...",
                    ContentData = System.Text.Encoding.UTF8.GetBytes("这是一段重要的会议记录内容，包含了销售目标和具体执行计划。")
                    });

                AddClipboardItem(new ClipboardItem
                    {
                    Title = "产品说明文档",
                    Type = "Text",
                    Description = "新产品功能说明和使用指南...",
                    ContentData = System.Text.Encoding.UTF8.GetBytes("产品主要功能包括...具体使用方法如下...")
                    });

                // 添加PPT内容测试数据
                AddClipboardItem(new ClipboardItem
                    {
                    Title = "项目展示PPT",
                    Type = "PPT",
                    Description = "项目进度报告幻灯片...",
                    ContentData = System.Text.Encoding.UTF8.GetBytes("PPT内容数据")
                    });

                AddClipboardItem(new ClipboardItem
                    {
                    Title = "营销方案",
                    Type = "PPT",
                    Description = "2024年营销策略方案...",
                    ContentData = System.Text.Encoding.UTF8.GetBytes("营销方案细内容")
                    });

                // 添加图片测试数据
                var defaultImagePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "default_image.png");
                if (File.Exists(defaultImagePath))
                    {
                    var imageBytes = File.ReadAllBytes(defaultImagePath);
                    AddClipboardItem(new ClipboardItem
                        {
                        Title = "产品展示图",
                        Type = "Image",
                        Description = "新产品外观展示图片...",
                        ContentData = imageBytes,
                        PreviewImageData = imageBytes
                        });

                    AddClipboardItem(new ClipboardItem
                        {
                        Title = "数据图表",
                        Type = "Image",
                        Description = "销售数据统计图表...",
                        ContentData = imageBytes,
                        PreviewImageData = imageBytes
                        });
                    }
                else
                    {
                    // 如果没有默认图片，创建一个简单的纯色图片
                    var bitmap = new WriteableBitmap(100, 100, 96, 96, PixelFormats.Bgr32, null);
                    var pixels = new byte[100 * 100 * 4];
                    for (int i = 0 ; i < pixels.Length ; i += 4)
                        {
                        pixels[i] = 200;     // Blue
                        pixels[i + 1] = 200; // Green
                        pixels[i + 2] = 200; // Red
                        pixels[i + 3] = 255; // Alpha
                        }
                    bitmap.WritePixels(new Int32Rect(0, 0, 100, 100), pixels, 400, 0);

                    using (var stream = new MemoryStream())
                        {
                        var encoder = new PngBitmapEncoder();
                        encoder.Frames.Add(BitmapFrame.Create(bitmap));
                        encoder.Save(stream);
                        var imageBytes = stream.ToArray();

                        AddClipboardItem(new ClipboardItem
                            {
                            Title = "示例图片1",
                            Type = "Image",
                            Description = "示例图片描述...",
                            ContentData = imageBytes,
                            PreviewImageData = imageBytes
                            });

                        AddClipboardItem(new ClipboardItem
                            {
                            Title = "示例图片2",
                            Type = "Image",
                            Description = "另一个示例图片...",
                            ContentData = imageBytes,
                            PreviewImageData = imageBytes
                            });
                        }
                    }

                MessageBox.Show("测试数据添加成功！", "提示", MessageBoxButton.OK, MessageBoxImage.Information);
                }
            catch (Exception ex)
                {
                MessageBox.Show($"添加测试数据失败: {ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }

        private void Button_Import_Click(object sender, RoutedEventArgs e)
            {
            try
                {
                var clipboardContent = System.Windows.Clipboard.GetDataObject();
                if (clipboardContent == null)
                    {
                    MessageBox.Show("剪贴板中没有内容", "提示", MessageBoxButton.OK, MessageBoxImage.Information);
                    return;
                    }

                // 获取所有可用的格式
                var formats = clipboardContent.GetFormats();

                // 检查是否包含PPT元素
                if (formats.Any(f => f.Contains("PowerPoint") || f.Contains("Slide") || f.Contains("Shape")))
                    {
                    ImportPPTContent(clipboardContent);
                    }
                else if (clipboardContent.GetDataPresent(DataFormats.Text))
                    {
                    ImportText(clipboardContent);
                    }
                else if (clipboardContent.GetDataPresent(DataFormats.Bitmap))
                    {
                    ImportImage(clipboardContent);
                    }
                else
                    {
                    MessageBox.Show("不支持的剪贴板内容格式", "提示", MessageBoxButton.OK, MessageBoxImage.Information);
                    }
                }
            catch (Exception ex)
                {
                MessageBox.Show($"导入内容失败: {ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }

        private void ImportText(IDataObject clipboardContent)
            {
            string text = clipboardContent.GetData(DataFormats.Text) as string;
            if (string.IsNullOrEmpty(text)) return;

            var item = new ClipboardItem
                {
                Title = text.Length > 20 ? text.Substring(0, 17) + "..." : text,
                Type = "Text",
                Description = text.Length > 50 ? text.Substring(0, 47) + "..." : text,
                ContentData = System.Text.Encoding.UTF8.GetBytes(text)
                };

            AddClipboardItem(item);
            }

        private void ImportImage(IDataObject clipboardContent)
            {
            var bitmap = clipboardContent.GetData(DataFormats.Bitmap) as System.Drawing.Bitmap;
            if (bitmap == null) return;

            using (var stream = new MemoryStream())
                {
                bitmap.Save(stream, System.Drawing.Imaging.ImageFormat.Png);
                var imageBytes = stream.ToArray();

                var item = new ClipboardItem
                    {
                    Title = $"图片 {DateTime.Now:yyyy-MM-dd HH:mm:ss}",
                    Type = "Image",
                    Description = $"从剪贴板导入的图片 ({bitmap.Width}x{bitmap.Height})",
                    ContentData = imageBytes,
                    PreviewImageData = imageBytes
                    };

                AddClipboardItem(item);
                }
            }

        private void ImportPPTContent(IDataObject clipboardContent)
            {
            try
                {
                // 获取PowerPoint应用程序实例
                var app = Globals.ThisAddIn.Application;
                if (app == null) return;

                // 保存原始剪贴板数据
                var formats = clipboardContent.GetFormats();
                var formatData = new Dictionary<string, byte[]>();
                foreach (var format in formats)
                    {
                    try
                        {
                        var data = clipboardContent.GetData(format);
                        if (data is MemoryStream ms)
                            {
                            formatData[format] = ms.ToArray();
                            }
                        else if (data is byte[] bytes)
                            {
                            formatData[format] = bytes;
                            }
                        else if (data != null)
                            {
                            using (var memStream = new MemoryStream())
                                {
                                var formatter = new System.Runtime.Serialization.Formatters.Binary.BinaryFormatter();
                                formatter.Serialize(memStream, data);
                                formatData[format] = memStream.ToArray();
                                }
                            }
                        }
                    catch (Exception)
                        {
                        // 忽略无法序列化的格式
                        continue;
                        }
                    }

                // 创建一个临时演示文稿来获取预览图
                var tempPresentation = app.Presentations.Add(MsoTriState.msoFalse);
                var tempSlide = tempPresentation.Slides.Add(1, PpSlideLayout.ppLayoutBlank);

                try
                    {
                    // 获取当前选中的内容
                    var selection = app.ActiveWindow.Selection;
                    if (selection != null)
                        {
                        // 复制选中的内容
                        selection.Copy();

                        // 粘贴到临时幻灯片
                        tempSlide.Shapes.Paste();

                        // 导出临时幻灯片为图片作为预览
                        string tempImagePath = Path.Combine(Path.GetTempPath(), $"preview_{Guid.NewGuid()}.png");
                        tempSlide.Export(tempImagePath, "PNG", 800, 600);

                        // 读取预览图片
                        byte[] previewImageData = File.ReadAllBytes(tempImagePath);
                        File.Delete(tempImagePath);

                        // 创建剪贴板项
                        var item = new ClipboardItem
                            {
                            Title = $"PPT元素 {DateTime.Now:yyyy-MM-dd HH:mm:ss}",
                            Type = "PPT",
                            Description = $"包含 {tempSlide.Shapes.Count} 个形状的PPT内容",
                            ContentData = System.Text.Encoding.UTF8.GetBytes(JsonConvert.SerializeObject(formatData)),
                            PreviewImageData = previewImageData
                            };

                        AddClipboardItem(item);
                        }
                    else
                        {
                        MessageBox.Show("请先选择要导入的PPT内容", "提示", MessageBoxButton.OK, MessageBoxImage.Information);
                        }
                    }
                finally
                    {
                    try
                        {
                        tempPresentation.Close();
                        }
                    catch { }
                    }
                }
            catch (Exception ex)
                {
                MessageBox.Show($"导入PPT内容失败: {ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
        }

    public class ClipboardItem : INotifyPropertyChanged
        {
        private string _id;
        private string _title;
        private string _type;
        private string _description;
        private byte[] _previewImageData;
        private byte[] _contentData;
        private DateTime _createdTime;
        private ImageSource _previewImage;

        public string Id
            {
            get => _id;
            set
                {
                if (_id != value)
                    {
                    _id = value;
                    OnPropertyChanged(nameof(Id));
                    }
                }
            }

        public string Title
            {
            get => _title;
            set
                {
                if (_title != value)
                    {
                    _title = value;
                    OnPropertyChanged(nameof(Title));
                    }
                }
            }

        public string Type
            {
            get => _type;
            set
                {
                if (_type != value)
                    {
                    _type = value;
                    OnPropertyChanged(nameof(Type));
                    }
                }
            }

        public string Description
            {
            get => _description;
            set
                {
                if (_description != value)
                    {
                    _description = value;
                    OnPropertyChanged(nameof(Description));
                    }
                }
            }

        public byte[] PreviewImageData
            {
            get => _previewImageData;
            set
                {
                if (_previewImageData != value)
                    {
                    _previewImageData = value;
                    OnPropertyChanged(nameof(PreviewImageData));
                    UpdatePreviewImage();
                    }
                }
            }

        public byte[] ContentData
            {
            get => _contentData;
            set
                {
                if (_contentData != value)
                    {
                    _contentData = value;
                    OnPropertyChanged(nameof(ContentData));
                    }
                }
            }

        public DateTime CreatedTime
            {
            get => _createdTime;
            set
                {
                if (_createdTime != value)
                    {
                    _createdTime = value;
                    OnPropertyChanged(nameof(CreatedTime));
                    }
                }
            }

        [BsonIgnore]
        public ImageSource PreviewImage
            {
            get => _previewImage;
            private set
                {
                if (_previewImage != value)
                    {
                    _previewImage = value;
                    OnPropertyChanged(nameof(PreviewImage));
                    }
                }
            }

        private void UpdatePreviewImage()
            {
            if (PreviewImageData == null)
                {
                PreviewImage = null;
                return;
                }

            try
                {
                Application.Current.Dispatcher.Invoke(() =>
                {
                    var image = new BitmapImage();
                    image.BeginInit();
                    image.StreamSource = new MemoryStream(PreviewImageData);
                    image.CacheOption = BitmapCacheOption.OnLoad;
                    image.EndInit();
                    image.Freeze(); // 提高性能
                    PreviewImage = image;
                });
                }
            catch
                {
                PreviewImage = null;
                }
            }

        public ClipboardItem()
            {
            Id = Guid.NewGuid().ToString();
            CreatedTime = DateTime.Now;
            }

        public event PropertyChangedEventHandler PropertyChanged;

        protected virtual void OnPropertyChanged(string propertyName)
            {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
            }
        }
    }