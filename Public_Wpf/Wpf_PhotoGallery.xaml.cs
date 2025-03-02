using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using HandyControl.Controls;
using MahApps.Metro.IconPacks;
using PresPio.Public_Wpf.Models;
using PresPio.Public_Wpf.Services;
using Application = System.Windows.Application;
using Button = System.Windows.Controls.Button;
using CancelEventArgs = System.ComponentModel.CancelEventArgs;
using Clipboard = System.Windows.Clipboard;
using ContextMenu = System.Windows.Controls.ContextMenu;
using DataFormats = System.Windows.DataFormats;
using DataObject = System.Windows.DataObject;
using FolderBrowserDialog = System.Windows.Forms.FolderBrowserDialog;
using MenuItem = System.Windows.Controls.MenuItem;

namespace PresPio.Public_Wpf
    {
    public partial class Wpf_PhotoGallery : HandyControl.Controls.Window, IDisposable
        {
        private DatabaseService _dbService;
        private ObservableCollection<TagItem> Tags { get; set; }
        private ObservableCollection<FilterTagItem> FilterTags { get; set; }
        private ObservableCollection<ImageItem> Images { get; set; }
        private string currentFolderPath;
        private readonly string[] supportedExtensions = { ".jpg", ".jpeg", ".png", ".gif", ".bmp" };
        private string ImageGalleryUrlKey => Properties.Settings.Default.ImageGalleryUrl;
        private const string DbFileName = "ImageGallery.db";
        private bool isInitialLoad = true;

        private readonly Brush[] TagColors = new[]
        {
            new SolidColorBrush((Color)ColorConverter.ConvertFromString("#3498db")), // 蓝色
            new SolidColorBrush((Color)ColorConverter.ConvertFromString("#2ecc71")), // 绿色
            new SolidColorBrush((Color)ColorConverter.ConvertFromString("#e74c3c")), // 红色
            new SolidColorBrush((Color)ColorConverter.ConvertFromString("#f1c40f")), // 黄色
            new SolidColorBrush((Color)ColorConverter.ConvertFromString("#9b59b6")), // 紫色
        };

        private int currentColorIndex = 0;

        private Point? lastDragPosition;

        //private double currentScale = 1.0;
        private const double MinScale = 0.1;

        private const double MaxScale = 5.0;
        private const double ScaleIncrement = 0.1;

        private ObservableCollection<CategoryItem> Categories { get; set; }
        private ImageItem selectedImage;

        private ObservableCollection<PathItem> ImagePaths { get; set; }
        private const string PathsKey = "ImageGalleryPaths";

        private ObservableCollection<TagItem> CommonTags { get; set; }

        private bool isLoading = false;  // 添加加载状态标志

        public Wpf_PhotoGallery()
            {
            InitializeComponent();

            // 初始化集合
            InitializeCategories();
            InitializeFilterTags();
            InitializeTags();
            InitializeImages();
            InitializeImagePaths();

            // 绑定选择文件夹按钮的点击事件
            SelectFolderButton.Click += SelectFolder_Click;
            CategoryTreeView.SelectedItemChanged += CategoryTreeView_SelectedItemChanged;

            // 加载已保存的数据
            LoadSavedData();

            // 修改窗口关闭行为
            this.Closing += Window_Closing;
            }

        private void Window_Closing(object sender, CancelEventArgs e)
            {
            try
                {
                // 保存所有未保存的更改
                SaveAllChanges();

                // 释放数据库连接
                _dbService?.Dispose();
                _dbService = null;

                e.Cancel = true;  // 取消关闭操作
                this.Hide();      // 隐藏窗口
                }
            catch (Exception ex)
                {
                HandyControl.Controls.MessageBox.Error($"关闭窗口时出错：{ex.Message}", "错误");
                }
            }

        private void SaveAllChanges()
            {
            try
                {
                // 保存图片路径
                SaveImagePaths();

                // 保存其他设置
                Properties.Settings.Default.Save();
                }
            catch (Exception ex)
                {
                HandyControl.Controls.MessageBox.Error($"保存设置时出错：{ex.Message}", "错误");
                }
            }

        public void Dispose()
            {
            try
                {
                // 保存所有更改
                SaveAllChanges();

                // 释放数据库连接
                _dbService?.Dispose();
                _dbService = null;

                // 关闭应用程序
                Application.Current.Dispatcher.InvokeShutdown();
                }
            catch (Exception ex)
                {
                HandyControl.Controls.MessageBox.Error($"释放资源时出错：{ex.Message}", "错误");
                }
            }

        protected override void OnClosed(EventArgs e)
            {
            try
                {
                // 释放数据库连接
                _dbService?.Dispose();
                _dbService = null;
                }
            catch (Exception ex)
                {
                HandyControl.Controls.MessageBox.Error($"关闭窗口时出错：{ex.Message}", "错误");
                }
            finally
                {
                base.OnClosed(e);
                }
            }

        private async void LoadSavedData()
            {
            // 加载图库路径
            var savedPath = Properties.Settings.Default.ImageGalleryUrl;
            if (!string.IsNullOrEmpty(savedPath) && savedPath != "ImageGalleryUrl")
                {
                currentFolderPath = savedPath;
                FolderPathText.Text = currentFolderPath;
                RootFolderText.Text = Path.GetFileName(currentFolderPath);
                InitializeDatabase(currentFolderPath);
                await LoadImagesFromFolder(currentFolderPath);
                }
            else
                {
                UpdatePathStatus(false, "未设置路径");
                RootFolderText.Text = "未设置";
                }
            }

        private void InitializeDatabase(string folderPath)
            {
            try
                {
                string dbPath = Path.Combine(folderPath, DbFileName);
                _dbService?.Dispose();

                // 如果数据库文件不存在，则创建新的数据库
                bool isNewDb = !File.Exists(dbPath);
                _dbService = new DatabaseService(dbPath);

                if (isNewDb)
                    {
                    HandyControl.Controls.Growl.Info("正在创建新的图库数据库...");

                    // 扫描并添加子文件夹作为分类
                    var directories = Directory.GetDirectories(folderPath, "*", SearchOption.TopDirectoryOnly);
                    int totalImages = 0;

                    // 处理主文件夹中的图片
                    var rootImages = Directory.GetFiles(folderPath)
                        .Where(file => supportedExtensions.Contains(Path.GetExtension(file).ToLower()))
                        .ToList();

                    foreach (var file in rootImages)
                        {
                        try
                            {
                            var fileInfo = new FileInfo(file);
                            // 获取图片尺寸
                            int width = 0, height = 0;
                            using (var stream = new FileStream(file, FileMode.Open, FileAccess.Read))
                                {
                                var decoder = BitmapDecoder.Create(stream, BitmapCreateOptions.None, BitmapCacheOption.None);
                                width = decoder.Frames[0].PixelWidth;
                                height = decoder.Frames[0].PixelHeight;
                                }

                            // 分析图片颜色
                            var dominantColors = ColorAnalyzer.AnalyzeImage(file);

                            var imageInfo = new ImageInfo
                                {
                                FilePath = file,
                                FileName = Path.GetFileName(file),
                                FileSize = GetFileSizeString(fileInfo.Length),
                                Width = width,
                                Height = height,
                                CreationTime = fileInfo.CreationTime,
                                ModificationTime = fileInfo.LastWriteTime,
                                LastAccessTime = DateTime.Now,
                                ImportTime = DateTime.Now,
                                DominantColors = dominantColors,
                                Tags = new List<string>()
                                };

                            _dbService.UpsertImage(imageInfo);
                            totalImages++;
                            }
                        catch (Exception ex)
                            {
                            HandyControl.Controls.Growl.Warning($"处理图片失败: {Path.GetFileName(file)} - {ex.Message}");
                            }
                        }

                    // 处理子文件夹
                    foreach (var dir in directories)
                        {
                        try
                            {
                            var dirName = Path.GetFileName(dir);
                            var imageFiles = Directory.GetFiles(dir)
                                .Where(file => supportedExtensions.Contains(Path.GetExtension(file).ToLower()))
                                .ToList();

                            if (imageFiles.Any())
                                {
                                var categoryInfo = new CategoryInfo
                                    {
                                    Name = dirName,
                                    Path = dir,
                                    ImageCount = imageFiles.Count,
                                    CreationTime = Directory.GetCreationTime(dir)
                                    };

                                _dbService.UpsertCategory(categoryInfo);

                                // 处理分类下的图片
                                foreach (var file in imageFiles)
                                    {
                                    try
                                        {
                                        var fileInfo = new FileInfo(file);
                                        int width = 0, height = 0;
                                        using (var stream = new FileStream(file, FileMode.Open, FileAccess.Read))
                                            {
                                            var decoder = BitmapDecoder.Create(stream, BitmapCreateOptions.None, BitmapCacheOption.None);
                                            width = decoder.Frames[0].PixelWidth;
                                            height = decoder.Frames[0].PixelHeight;
                                            }

                                        var dominantColors = ColorAnalyzer.AnalyzeImage(file);

                                        var imageInfo = new ImageInfo
                                            {
                                            FilePath = file,
                                            FileName = Path.GetFileName(file),
                                            FileSize = GetFileSizeString(fileInfo.Length),
                                            Width = width,
                                            Height = height,
                                            CreationTime = fileInfo.CreationTime,
                                            ModificationTime = fileInfo.LastWriteTime,
                                            LastAccessTime = DateTime.Now,
                                            ImportTime = DateTime.Now,
                                            Category = dirName,
                                            DominantColors = dominantColors,
                                            Tags = new List<string>()
                                            };

                                        _dbService.UpsertImage(imageInfo);
                                        totalImages++;
                                        }
                                    catch (Exception ex)
                                        {
                                        HandyControl.Controls.Growl.Warning($"处理图片失败: {Path.GetFileName(file)} - {ex.Message}");
                                        }
                                    }
                                }
                            }
                        catch (Exception ex)
                            {
                            HandyControl.Controls.Growl.Warning($"处理文件夹失败: {Path.GetFileName(dir)} - {ex.Message}");
                            }
                        }

                    var categoryCount = directories.Count(d => Directory.GetFiles(d).Any(f => supportedExtensions.Contains(Path.GetExtension(f).ToLower())));
                    HandyControl.Controls.Growl.Success($"已导入 {categoryCount} 个分类，共 {totalImages} 张图片");
                    }
                else if (isInitialLoad)
                    {
                    HandyControl.Controls.Growl.Info("正在加载现有图库数据库...");
                    LoadExistingData();
                    }

                // 数据库初始化完成后，重新加载标签
                LoadTagsFromDatabase();
                }
            catch (Exception ex)
                {
                HandyControl.Controls.MessageBox.Error($"初始化数据库时出错：{ex.Message}", "错误");
                }
            }

        private void LoadExistingData()
            {
            try
                {
                // 加载分类
                Categories.Clear();
                var allCategories = _dbService.GetAllCategories()
                    .OrderBy(c => c.Name)
                    .ToList();

                foreach (var category in allCategories)
                    {
                    // 验证分类文件夹是否存在
                    if (Directory.Exists(category.Path))
                        {
                        // 更新图片数量
                        var imageCount = Directory.GetFiles(category.Path)
                            .Count(file => supportedExtensions.Contains(Path.GetExtension(file).ToLower()));

                        if (category.ImageCount != imageCount)
                            {
                            category.ImageCount = imageCount;
                            _dbService.UpsertCategory(category);
                            }

                        // 只添加包含图片的分类
                        if (imageCount > 0)
                            {
                            Categories.Add(new CategoryItem
                                {
                                Name = category.Name,
                                Path = category.Path,
                                ImageCount = imageCount
                                });
                            }
                        }
                    else
                        {
                        // 如果文件夹不存在，从数据库中删除该分类
                        _dbService.DeleteCategory(category.Name);
                        }
                    }

                // 加载标签
                Tags.Clear();
                FilterTags.Clear();
                foreach (var tag in _dbService.GetAllTags())
                    {
                    var color = new SolidColorBrush((Color)ColorConverter.ConvertFromString(tag.ColorHex));
                    var tagItem = new TagItem { Name = tag.Name, Color = color };
                    Tags.Add(tagItem);

                    if (tag.ImageCount > 0)
                        {
                        FilterTags.Add(new FilterTagItem
                            {
                            Name = tag.Name,
                            Color = color,
                            IsSelected = false
                            });
                        }
                    }

                // 更新UI显示
                if (Categories.Any())
                    {
                    HandyControl.Controls.Growl.Success($"已加载 {Categories.Count} 个分类");
                    }
                else
                    {
                    HandyControl.Controls.Growl.Info("暂无分类");
                    }
                }
            catch (Exception ex)
                {
                HandyControl.Controls.MessageBox.Error($"加载现有数据时出错：{ex.Message}", "错误");
                }
            }

        private async Task LoadImagesFromFolder(string folderPath)
            {
            // 如果正在加载，直接返回
            if (isLoading) return;

            try
                {
                isLoading = true;
                if (Application.Current?.Dispatcher == null) return;

                await Application.Current.Dispatcher.InvokeAsync(() =>
                {
                    LoadingMask.Visibility = Visibility.Visible;
                    LoadingProgress.Value = 0;
                    Images.Clear();  // 清空当前图片列表
                });

                var files = Directory.GetFiles(folderPath)
                    .Where(file => supportedExtensions.Contains(Path.GetExtension(file).ToLower()))
                    .ToList();

                var totalFiles = files.Count;
                var loadedFiles = 0;

                foreach (var file in files)
                    {
                    try
                        {
                        var fileInfo = new FileInfo(file);
                        var existingImage = _dbService.GetImageByPath(file);

                        if (existingImage != null)
                            {
                            // 检查文件是否被修改
                            if (existingImage.ModificationTime != fileInfo.LastWriteTime)
                                {
                                // 获取新的图片尺寸
                                using (var stream = new FileStream(file, FileMode.Open, FileAccess.Read))
                                    {
                                    var decoder = BitmapDecoder.Create(stream, BitmapCreateOptions.None, BitmapCacheOption.None);
                                    existingImage.Width = decoder.Frames[0].PixelWidth;
                                    existingImage.Height = decoder.Frames[0].PixelHeight;
                                    }
                                existingImage.ModificationTime = fileInfo.LastWriteTime;
                                existingImage.FileSize = GetFileSizeString(fileInfo.Length);
                                existingImage.DominantColors = ColorAnalyzer.AnalyzeImage(file);
                                _dbService.UpsertImage(existingImage);
                                }
                            }
                        else
                            {
                            // 获取新图片的尺寸
                            int width = 0, height = 0;
                            using (var stream = new FileStream(file, FileMode.Open, FileAccess.Read))
                                {
                                var decoder = BitmapDecoder.Create(stream, BitmapCreateOptions.None, BitmapCacheOption.None);
                                width = decoder.Frames[0].PixelWidth;
                                height = decoder.Frames[0].PixelHeight;
                                }

                            // 分析图片颜色
                            var dominantColors = ColorAnalyzer.AnalyzeImage(file);

                            // 创建新的图片信息
                            existingImage = new ImageInfo
                                {
                                FilePath = file,
                                FileName = Path.GetFileName(file),
                                FileSize = GetFileSizeString(fileInfo.Length),
                                Width = width,
                                Height = height,
                                CreationTime = fileInfo.CreationTime,
                                ModificationTime = fileInfo.LastWriteTime,
                                LastAccessTime = DateTime.Now,
                                ImportTime = DateTime.Now,
                                DominantColors = dominantColors,
                                Tags = new List<string>()
                                };
                            _dbService.UpsertImage(existingImage);
                            }

                        // 加载到UI
                        if (Application.Current?.Dispatcher != null)
                            {
                            await Application.Current.Dispatcher.InvokeAsync(() =>
                            {
                                // 检查图片是否已经在列表中
                                if (!Images.Any(img => img.FilePath == existingImage.FilePath))
                                    {
                                    var image = new ImageItem
                                        {
                                        FilePath = existingImage.FilePath,
                                        FileName = existingImage.FileName,
                                        FileSize = existingImage.FileSize,
                                        Width = existingImage.Width,
                                        Height = existingImage.Height,
                                        CreationTime = existingImage.CreationTime,
                                        ModificationTime = existingImage.ModificationTime
                                        };

                                    var bitmap = new BitmapImage();
                                    bitmap.BeginInit();
                                    bitmap.UriSource = new Uri(file);
                                    bitmap.CacheOption = BitmapCacheOption.OnLoad;
                                    bitmap.DecodePixelWidth = 400;
                                    bitmap.EndInit();
                                    bitmap.Freeze();
                                    image.Thumbnail = bitmap;

                                    // 加载标签
                                    if (existingImage.Tags != null)
                                        {
                                        foreach (var tagName in existingImage.Tags)
                                            {
                                            var tag = Tags.FirstOrDefault(t => t.Name == tagName);
                                            if (tag != null)
                                                {
                                                image.Tags.Add(tag);
                                                }
                                            }
                                        }

                                    // 加载分类
                                    if (!string.IsNullOrEmpty(existingImage.Category))
                                        {
                                        var category = Categories.FirstOrDefault(c => c.Name == existingImage.Category);
                                        if (category != null)
                                            {
                                            image.Category = category;
                                            }
                                        }

                                    Images.Add(image);

                                    // 如果是选中的图片，更新颜色分析面板
                                    if (selectedImage != null && selectedImage.FilePath == image.FilePath)
                                        {
                                        UpdateColorAnalysis(existingImage.DominantColors);
                                        }
                                    }

                                // 更新进度
                                loadedFiles++;
                                var progress = (double)loadedFiles / totalFiles * 100;
                                LoadingProgress.Value = progress;
                                LoadingText.Text = $"正在加载图片... ({loadedFiles}/{totalFiles})";
                            });
                            }
                        }
                    catch (Exception ex)
                        {
                        HandyControl.Controls.Growl.Warning($"加载图片失败: {Path.GetFileName(file)} - {ex.Message}");
                        }
                    }

                if (Application.Current?.Dispatcher != null)
                    {
                    await Application.Current.Dispatcher.InvokeAsync(() =>
                    {
                        LoadingMask.Visibility = Visibility.Collapsed;
                    });
                    }
                }
            catch (Exception ex)
                {
                if (Application.Current?.Dispatcher != null)
                    {
                    await Application.Current.Dispatcher.InvokeAsync(() =>
                    {
                        HandyControl.Controls.MessageBox.Error($"加载图片时出错：{ex.Message}", "错误");
                        LoadingMask.Visibility = Visibility.Collapsed;
                    });
                    }
                }
            finally
                {
                isLoading = false;
                }
            }

        private async Task UpdateImage(string file, FileInfo fileInfo, ImageInfo existingImage)
            {
            // 更新图片信息
            existingImage.ModificationTime = fileInfo.LastWriteTime;
            existingImage.FileSize = GetFileSizeString(fileInfo.Length);
            existingImage.DominantColors = ColorAnalyzer.AnalyzeImage(file);
            _dbService.UpsertImage(existingImage);

            await LoadImageToUI(existingImage);
            }

        private async Task AddNewImage(string file, FileInfo fileInfo)
            {
            try
                {
                // 分析图片颜色
                var dominantColors = ColorAnalyzer.AnalyzeImage(file);

                // 创建新的图片信息
                var imageInfo = new ImageInfo
                    {
                    FilePath = file,
                    FileName = Path.GetFileName(file),
                    FileSize = GetFileSizeString(fileInfo.Length),
                    CreationTime = fileInfo.CreationTime,
                    ModificationTime = fileInfo.LastWriteTime,
                    LastAccessTime = DateTime.Now,
                    ImportTime = DateTime.Now,
                    DominantColors = dominantColors
                    };

                // 保存到数据库
                _dbService.UpsertImage(imageInfo);

                await LoadImageToUI(imageInfo);
                }
            catch (Exception ex)
                {
                if (Application.Current?.Dispatcher != null)
                    {
                    await Application.Current.Dispatcher.InvokeAsync(() =>
                    {
                        HandyControl.Controls.Growl.Warning($"添加新图片失败: {Path.GetFileName(file)} - {ex.Message}");
                    });
                    }
                }
            }

        private async Task LoadImageToUI(ImageInfo imageInfo)
            {
            if (Application.Current?.Dispatcher == null) return;

            await Application.Current.Dispatcher.InvokeAsync(() =>
            {
                try
                    {
                    var image = new ImageItem
                        {
                        FilePath = imageInfo.FilePath,
                        FileName = imageInfo.FileName,
                        FileSize = imageInfo.FileSize,
                        Width = imageInfo.Width,
                        Height = imageInfo.Height,
                        CreationTime = imageInfo.CreationTime,
                        ModificationTime = imageInfo.ModificationTime
                        };

                    var bitmap = new BitmapImage();
                    bitmap.BeginInit();
                    bitmap.UriSource = new Uri(imageInfo.FilePath);
                    bitmap.CacheOption = BitmapCacheOption.OnLoad;
                    bitmap.DecodePixelWidth = 400;
                    bitmap.EndInit();
                    bitmap.Freeze();
                    image.Thumbnail = bitmap;

                    // 加载标签
                    if (imageInfo.Tags != null)
                        {
                        foreach (var tagName in imageInfo.Tags)
                            {
                            var tag = Tags.FirstOrDefault(t => t.Name == tagName);
                            if (tag != null)
                                {
                                image.Tags.Add(tag);
                                }
                            }
                        }

                    // 加载分类
                    if (!string.IsNullOrEmpty(imageInfo.Category))
                        {
                        var category = Categories.FirstOrDefault(c => c.Name == imageInfo.Category);
                        if (category != null)
                            {
                            image.Category = category;
                            }
                        }

                    Images.Add(image);
                    }
                catch (Exception ex)
                    {
                    HandyControl.Controls.Growl.Warning($"加载图片到界面失败: {imageInfo.FileName} - {ex.Message}");
                    }
            });
            }

        private string GetFileSizeString(long bytes)
            {
            string[] sizes = { "B", "KB", "MB", "GB" };
            int order = 0;
            double len = bytes;
            while (len >= 1024 && order < sizes.Length - 1)
                {
                order++;
                len = len / 1024;
                }
            return $"{len:0.##} {sizes[order]}";
            }

        private void InitializeCategories()
            {
            Categories = new ObservableCollection<CategoryItem>();
            CategoryList.ItemsSource = Categories;
            }

        private void AddCategory_Click(object sender, RoutedEventArgs e)
            {
            var categoryName = NewCategoryInput.Text?.Trim();
            if (string.IsNullOrWhiteSpace(categoryName))
                {
                HandyControl.Controls.Growl.Warning("请输入分类名称");
                return;
                }

            if (Categories.Any(c => c.Name == categoryName))
                {
                HandyControl.Controls.Growl.Warning("该分类已存在");
                return;
                }

            if (string.IsNullOrEmpty(currentFolderPath))
                {
                HandyControl.Controls.Growl.Warning("请先设置主文件夹");
                return;
                }

            try
                {
                // 创建实际文件夹
                string categoryPath = Path.Combine(currentFolderPath, categoryName);
                if (!Directory.Exists(categoryPath))
                    {
                    Directory.CreateDirectory(categoryPath);
                    }

                // 保存到数据库
                var categoryInfo = new CategoryInfo
                    {
                    Name = categoryName,
                    Path = categoryPath,
                    ImageCount = 0,
                    CreationTime = DateTime.Now
                    };
                _dbService.UpsertCategory(categoryInfo);

                // 更新UI
                var newCategory = new CategoryItem
                    {
                    Name = categoryName,
                    Path = categoryPath,
                    ImageCount = 0
                    };
                Categories.Add(newCategory);

                // 清空输入框
                NewCategoryInput.Clear();
                HandyControl.Controls.Growl.Success($"已添加分类：{categoryName}");
                }
            catch (Exception ex)
                {
                HandyControl.Controls.MessageBox.Error($"创建分类文件夹时出错：{ex.Message}", "错误");
                }
            }

        private void FilterImagesByCategory(CategoryItem category)
            {
            if (category == null) return;

            try
                {
                // 清空当前显示的图片
                Images.Clear();

                // 从数据库获取该分类下的所有图片
                var categoryImages = _dbService.GetImagesByCategory(category.Name);
                if (categoryImages != null && categoryImages.Any())
                    {
                    // 使用同步方式加载图片，避免重复
                    foreach (var imageInfo in categoryImages)
                        {
                        if (File.Exists(imageInfo.FilePath))
                            {
                            var image = new ImageItem
                                {
                                FilePath = imageInfo.FilePath,
                                FileName = imageInfo.FileName,
                                FileSize = imageInfo.FileSize,
                                CreationTime = imageInfo.CreationTime,
                                ModificationTime = imageInfo.ModificationTime
                                };

                            var bitmap = new BitmapImage();
                            bitmap.BeginInit();
                            bitmap.UriSource = new Uri(imageInfo.FilePath);
                            bitmap.CacheOption = BitmapCacheOption.OnLoad;
                            bitmap.DecodePixelWidth = 400;
                            bitmap.EndInit();
                            bitmap.Freeze();
                            image.Thumbnail = bitmap;

                            // 加载标签
                            if (imageInfo.Tags != null)
                                {
                                foreach (var tagName in imageInfo.Tags)
                                    {
                                    var tag = Tags.FirstOrDefault(t => t.Name == tagName);
                                    if (tag != null)
                                        {
                                        image.Tags.Add(tag);
                                        }
                                    }
                                }

                            Images.Add(image);
                            }
                        }
                    }
                else
                    {
                    HandyControl.Controls.Growl.Info($"分类 {category.Name} 中暂无图片");
                    }
                }
            catch (Exception ex)
                {
                HandyControl.Controls.MessageBox.Error($"加载分类图片时出错：{ex.Message}", "错误");
                }
            }

        private void InitializeImages()
            {
            Images = new ObservableCollection<ImageItem>();
            ImageList.ItemsSource = Images;
            }

        private void SelectFolder_Click(object sender, RoutedEventArgs e)
            {
            using (var dialog = new FolderBrowserDialog())
                {
                dialog.Description = "选择图片文件夹";
                if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                    {
                    currentFolderPath = dialog.SelectedPath;
                    FolderPathText.Text = currentFolderPath;
                    RootFolderText.Text = Path.GetFileName(currentFolderPath);

                    // 保存图库路径到设置
                    Properties.Settings.Default.ImageGalleryUrl = currentFolderPath;
                    Properties.Settings.Default.Save();

                    _ = LoadImagesFromFolder(currentFolderPath);
                    }
                }
            }

        private void InitializeTags()
            {
            try
                {
                Tags = new ObservableCollection<TagItem>();
                CommonTags = new ObservableCollection<TagItem>();
                TagsItemsControl.ItemsSource = Tags;
                CommonTagsControl.ItemsSource = CommonTags;

                // 如果数据库已初始化，则加载标签
                if (_dbService != null)
                    {
                    LoadTagsFromDatabase();
                    }
                }
            catch (Exception ex)
                {
                HandyControl.Controls.MessageBox.Error($"初始化标签时出错：{ex.Message}", "错误");
                }
            }

        private void LoadTagsFromDatabase()
            {
            try
                {
                if (_dbService == null) return;

                // 加载所有标签
                var allTags = _dbService.GetAllTags();
                Tags.Clear();
                FilterTags.Clear();
                CommonTags.Clear();

                foreach (var tag in allTags)
                    {
                    var color = new SolidColorBrush((Color)ColorConverter.ConvertFromString(tag.ColorHex));
                    var tagItem = new TagItem { Name = tag.Name, Color = color };

                    // 添加到标签列表
                    Tags.Add(tagItem);

                    // 添加到筛选标签列表
                    if (tag.ImageCount > 0)
                        {
                        FilterTags.Add(new FilterTagItem
                            {
                            Name = tag.Name,
                            Color = color,
                            IsSelected = false
                            });
                        }
                    }

                // 更新常用标签
                UpdateCommonTags();
                }
            catch (Exception ex)
                {
                HandyControl.Controls.MessageBox.Error($"从数据库加载标签时出错：{ex.Message}", "错误");
                }
            }

        private void InitializeFilterTags()
            {
            try
                {
                FilterTags = new ObservableCollection<FilterTagItem>();
                FilterTagsControl.ItemsSource = FilterTags;

                // 如果数据库已初始化，则加载标签
                if (_dbService != null)
                    {
                    UpdateFilterTags();
                    }
                }
            catch (Exception ex)
                {
                HandyControl.Controls.MessageBox.Error($"初始化筛选标签时出错：{ex.Message}", "错误");
                }
            }

        private void UpdateFilterTags()
            {
            try
                {
                if (_dbService == null) return;

                // 获取所有已使用的标签
                var usedTags = _dbService.GetAllTags()
                    .Where(t => t.ImageCount > 0)
                    .OrderByDescending(t => t.ImageCount);

                FilterTags.Clear();
                foreach (var tag in usedTags)
                    {
                    var color = new SolidColorBrush((Color)ColorConverter.ConvertFromString(tag.ColorHex));
                    FilterTags.Add(new FilterTagItem
                        {
                        Name = tag.Name,
                        Color = color,
                        IsSelected = false
                        });
                    }
                }
            catch (Exception ex)
                {
                HandyControl.Controls.MessageBox.Error($"更新标签筛选列表时出错：{ex.Message}", "错误");
                }
            }

        private void AddTagToImage_Click(object sender, RoutedEventArgs e)
            {
            if (selectedImage == null)
                {
                HandyControl.Controls.Growl.Warning("请先选择一张图片");
                return;
                }

            var tagName = ImageTagInput.Text?.Trim();
            if (string.IsNullOrWhiteSpace(tagName))
                {
                HandyControl.Controls.Growl.Warning("请输入标签名称");
                return;
                }

            try
                {
                // 检查标签是否存在，不存在则创建
                var tagInfo = _dbService.GetTag(tagName);
                if (tagInfo == null)
                    {
                    var color = TagColors[currentColorIndex];
                    var colorHex = ((SolidColorBrush)color).Color.ToString();

                    tagInfo = new TagInfo
                        {
                        Name = tagName,
                        ColorHex = colorHex,
                        ImageCount = 0,
                        CreationTime = DateTime.Now
                        };
                    _dbService.UpsertTag(tagInfo);

                    currentColorIndex = (currentColorIndex + 1) % TagColors.Length;
                    }

                // 为图片添加标签
                var imageInfo = _dbService.GetImageByPath(selectedImage.FilePath);
                if (imageInfo != null)
                    {
                    if (!imageInfo.Tags.Contains(tagName))
                        {
                        imageInfo.Tags.Add(tagName);
                        _dbService.UpsertImage(imageInfo);
                        _dbService.UpdateTagImageCount(tagName);

                        // 更新UI
                        var tag = new TagItem
                            {
                            Name = tagName,
                            Color = new SolidColorBrush((Color)ColorConverter.ConvertFromString(tagInfo.ColorHex))
                            };
                        if (!selectedImage.Tags.Any(t => t.Name == tagName))
                            {
                            selectedImage.Tags.Add(tag);
                            }

                        // 更新常用标签和筛选标签列表
                        UpdateCommonTags();
                        UpdateFilterTags();

                        HandyControl.Controls.Growl.Success($"已添加标签：{tagName}");
                        }
                    else
                        {
                        HandyControl.Controls.Growl.Warning("该图片已包含此标签");
                        }
                    }

                // 清空输入框
                ImageTagInput.Clear();
                }
            catch (Exception ex)
                {
                HandyControl.Controls.MessageBox.Error($"添加标签时出错：{ex.Message}", "错误");
                }
            }

        private void RemoveTagFromImage_Click(object sender, System.Windows.Input.MouseButtonEventArgs e)
            {
            if (sender is HandyControl.Controls.Tag tag &&
                tag.DataContext is TagItem tagItem &&
                selectedImage != null)
                {
                try
                    {
                    // 从数据库中移除标签
                    var imageInfo = _dbService.GetImageByPath(selectedImage.FilePath);
                    if (imageInfo != null)
                        {
                        imageInfo.Tags.Remove(tagItem.Name);
                        _dbService.UpsertImage(imageInfo);

                        // 更新标签使用计数
                        _dbService.UpdateTagImageCount(tagItem.Name);

                        // 从UI中移除标签
                        selectedImage.Tags.Remove(tagItem);

                        // 更新常用标签和筛选标签列表
                        UpdateCommonTags();
                        UpdateFilterTags();

                        HandyControl.Controls.Growl.Success($"已移除标签：{tagItem.Name}");
                        }
                    }
                catch (Exception ex)
                    {
                    HandyControl.Controls.MessageBox.Error($"移除标签时出错：{ex.Message}", "错误");
                    }
                }
            }

        private void FilterTag_Click(object sender, System.Windows.Input.MouseButtonEventArgs e)
            {
            try
                {
                if (sender is HandyControl.Controls.Tag tag &&
                    tag.DataContext is FilterTagItem filterTagItem)
                    {
                    // 切换选中状态
                    filterTagItem.IsSelected = !filterTagItem.IsSelected;

                    // 更新标签样式
                    tag.BorderThickness = filterTagItem.IsSelected ? new Thickness(2) : new Thickness(0);
                    tag.BorderBrush = new SolidColorBrush(Colors.White);

                    // 获取所有选中的标签
                    var selectedTags = FilterTags.Where(t => t.IsSelected).Select(t => t.Name).ToList();

                    if (selectedTags.Any())
                        {
                        // 获取包含任一选中标签的图片
                        var images = new List<ImageInfo>();
                        foreach (var tagName in selectedTags)
                            {
                            var tagImages = _dbService.GetImagesByTag(tagName);
                            images.AddRange(tagImages.Where(img => !images.Any(existingImg => existingImg.FilePath == img.FilePath)));
                            }

                        // 更新UI
                        Images.Clear();
                        foreach (var imageInfo in images)
                            {
                            _ = LoadImageToUI(imageInfo);
                            }
                        }
                    else
                        {
                        // 如果没有选中的标签，显示所有图片
                        if (!string.IsNullOrEmpty(currentFolderPath))
                            {
                            _ = LoadImagesFromFolder(currentFolderPath);
                            }
                        }
                    }
                }
            catch (Exception ex)
                {
                HandyControl.Controls.MessageBox.Error($"筛选图片时出错：{ex.Message}", "错误");
                }
            }

        private void Image_Click(object sender, System.Windows.Input.MouseButtonEventArgs e)
            {
            if (sender is Card card && card.DataContext is ImageItem imageItem)
                {
                // 取消之前选中的图片
                if (selectedImage != null)
                    {
                    selectedImage.IsSelected = false;
                    }

                // 选中当前图片
                selectedImage = imageItem;
                imageItem.IsSelected = true;

                // 更新右侧信息面板
                UpdateImageDetails(imageItem);
                }
            }

        private void UpdateImageDetails(ImageItem image)
            {
            if (image == null) return;

            try
                {
                // 更新右侧面板的信息
                ImagePreview.Source = image.Thumbnail;
                FileNameText.Text = image.FileName;
                FileSizeText.Text = image.FileSize;
                DimensionsText.Text = image.Dimensions;
                CreationTimeText.Text = image.CreationTime.ToString("yyyy-MM-dd HH:mm:ss");
                ModificationTimeText.Text = image.ModificationTime.ToString("yyyy-MM-dd HH:mm:ss");

                // 从数据库获取完整信息
                var imageInfo = _dbService.GetImageByPath(image.FilePath);
                if (imageInfo != null)
                    {
                    // 更新标签显示
                    TagsItemsControl.ItemsSource = image.Tags;

                    // 更新颜色分析区域
                    UpdateColorAnalysis(imageInfo.DominantColors);
                    }
                }
            catch (Exception ex)
                {
                HandyControl.Controls.Growl.Error($"更新图片详情时出错：{ex.Message}");
                }
            }

        private void UpdateColorAnalysis(List<Models.ColorInfo> colors)
            {
            ColorAnalysisPanel.Children.Clear();

            // 只取前8个主要颜色，如果不足8个则补充透明色
            var mainColors = colors.OrderByDescending(c => c.Percentage).Take(8).ToList();
            while (mainColors.Count < 8)
                {
                mainColors.Add(new Models.ColorInfo { ColorHex = "#00FFFFFF", Percentage = 0 });
                }

            foreach (var color in mainColors)
                {
                var container = new Grid();

                var border = new Border
                    {
                    Background = new SolidColorBrush((Color)ColorConverter.ConvertFromString(color.ColorHex)),
                    CornerRadius = new CornerRadius(4),
                    Height = 35,
                    Margin = new Thickness(2)
                    };

                var percentageBackground = new Border
                    {
                    Background = new SolidColorBrush(Color.FromArgb(128, 0, 0, 0)),
                    Height = 16,
                    VerticalAlignment = VerticalAlignment.Bottom,
                    CornerRadius = new CornerRadius(0, 0, 4, 4)
                    };

                var percentage = new TextBlock
                    {
                    Text = $"{color.Percentage:P0}",
                    Foreground = Brushes.White,
                    HorizontalAlignment = System.Windows.HorizontalAlignment.Center,
                    VerticalAlignment = VerticalAlignment.Center,
                    FontSize = 10,
                    Effect = new System.Windows.Media.Effects.DropShadowEffect
                        {
                        ShadowDepth = 1,
                        BlurRadius = 2,
                        Color = Colors.Black,
                        Opacity = 0.5
                        }
                    };

                // 为颜色块添加工具提示
                var colorHex = color.ColorHex.ToUpper();
                var tooltip = new System.Windows.Controls.ToolTip
                    {
                    Content = $"颜色值: {colorHex}\n占比: {color.Percentage:P1}"
                    };
                ToolTipService.SetToolTip(border, tooltip);

                // 组装UI元素
                container.Children.Add(border);
                if (color.Percentage > 0)  // 只为有效颜色显示百分比
                    {
                    var percentageContainer = new Grid();
                    percentageContainer.Children.Add(percentageBackground);
                    percentageContainer.Children.Add(percentage);
                    container.Children.Add(percentageContainer);
                    }

                ColorAnalysisPanel.Children.Add(container);
                }
            }

        private async void PerformSearch(string searchText)
            {
            try
                {
                searchText = searchText?.ToLower() ?? string.Empty;

                if (string.IsNullOrWhiteSpace(searchText))
                    {
                    // 如果搜索文本为空，显示所有图片
                    if (!string.IsNullOrEmpty(currentFolderPath))
                        {
                        await LoadImagesFromFolder(currentFolderPath);
                        }
                    return;
                    }

                // 从数据库中搜索图片
                var allImages = Directory.GetFiles(currentFolderPath)
                    .Where(file => supportedExtensions.Contains(Path.GetExtension(file).ToLower()))
                    .Select(file => _dbService.GetImageByPath(file))
                    .Where(img => img != null)
                    .Where(img =>
                        img.FileName.ToLower().Contains(searchText) ||
                        img.Tags.Any(tag => tag.ToLower().Contains(searchText)) ||
                        img.Category?.ToLower().Contains(searchText) == true
                    );

                // 更新UI
                Images.Clear();
                foreach (var imageInfo in allImages)
                    {
                    await LoadImageToUI(imageInfo);
                    }
                }
            catch (Exception ex)
                {
                HandyControl.Controls.Growl.Error($"搜索时出错：{ex.Message}");
                }
            }

        private async void FilterByColor(Color selectedColor)
            {
            try
                {
                if (string.IsNullOrEmpty(currentFolderPath)) return;

                // 将选中的颜色转换为 HSL
                var selectedColorInfo = new Models.ColorInfo
                    {
                    ColorHex = $"#{selectedColor.R:X2}{selectedColor.G:X2}{selectedColor.B:X2}",
                    Hsl = ColorAnalyzer.RgbToHsl(selectedColor.R, selectedColor.G, selectedColor.B)
                    };

                // 获取当前文件夹中的所有图片
                var allImages = Directory.GetFiles(currentFolderPath)
                    .Where(file => supportedExtensions.Contains(Path.GetExtension(file).ToLower()))
                    .Select(file => _dbService.GetImageByPath(file))
                    .Where(img => img != null)
                    .ToList();

                // 筛选包含相似颜色的图片
                var filteredImages = allImages.Where(img =>
                    img.DominantColors.Any(c =>
                    {
                        // 将颜色字符串转换为HSL
                        var hsl = ColorAnalyzer.RgbToHsl(
                            Convert.ToInt32(c.ColorHex.Substring(1, 2), 16),
                            Convert.ToInt32(c.ColorHex.Substring(3, 2), 16),
                            Convert.ToInt32(c.ColorHex.Substring(5, 2), 16)
                        );

                        var colorInfo = new Models.ColorInfo
                            {
                            ColorHex = c.ColorHex,
                            Hsl = hsl,
                            Percentage = c.Percentage
                            };

                        return ColorAnalyzer.AreColorsSimilar(colorInfo, selectedColorInfo);
                    })
                ).ToList();

                // 更新UI
                Images.Clear();
                foreach (var imageInfo in filteredImages)
                    {
                    await LoadImageToUI(imageInfo);
                    }

                // 获取选中颜色最接近的标准颜色名称
                var standardColorName = ColorAnalyzer.GetClosestStandardColor(selectedColorInfo);
                }
            catch (Exception ex)
                {
                HandyControl.Controls.Growl.Error($"颜色筛选时出错：{ex.Message}");
                }
            }

        private bool IsColorSimilar(Color c1, Color c2, double tolerance)
            {
            // 计算RGB分量的差异
            double rDiff = Math.Abs(c1.R - c2.R);
            double gDiff = Math.Abs(c1.G - c2.G);
            double bDiff = Math.Abs(c1.B - c2.B);

            // 如果任何一个分量的差异超过容差，则认为颜色不相似
            return rDiff <= tolerance && gDiff <= tolerance && bDiff <= tolerance;
            }

        private void DeleteImage_Click(object sender, RoutedEventArgs e)
            {
            if (selectedImage == null)
                {
                HandyControl.Controls.Growl.Warning("请先选择一张图片");
                return;
                }

            var result = HandyControl.Controls.MessageBox.Show(
                $"确定要删除图片 {selectedImage.FileName} 吗？",
                "删除确认",
                MessageBoxButton.YesNo,
                MessageBoxImage.Warning);

            if (result == MessageBoxResult.Yes)
                {
                try
                    {
                    // 从数据库中获取完整信息
                    var imageInfo = _dbService.GetImageByPath(selectedImage.FilePath);
                    if (imageInfo != null)
                        {
                        // 更新标签计数
                        foreach (var tag in imageInfo.Tags)
                            {
                            _dbService.UpdateTagImageCount(tag);
                            }

                        // 更新分类计数
                        if (!string.IsNullOrEmpty(imageInfo.Category))
                            {
                            _dbService.UpdateCategoryImageCount(imageInfo.Category);
                            }

                        // 删除文件
                        File.Delete(selectedImage.FilePath);

                        // 从UI中移除
                        Images.Remove(selectedImage);
                        selectedImage = null;

                        // 清空右侧预览
                        ImagePreview.Source = null;
                        FileNameText.Text = "";
                        FileSizeText.Text = "";
                        DimensionsText.Text = "";
                        CreationTimeText.Text = "";
                        ModificationTimeText.Text = "";
                        ColorAnalysisPanel.Children.Clear();

                        HandyControl.Controls.Growl.Success("图片已删除");
                        }
                    }
                catch (Exception ex)
                    {
                    HandyControl.Controls.MessageBox.Error($"删除图片时出错：{ex.Message}", "错误");
                    }
                }
            }

        private void ShowInExplorer_Click(object sender, RoutedEventArgs e)
            {
            if (selectedImage == null)
                {
                HandyControl.Controls.Growl.Warning("请先选择一张图片");
                return;
                }

            try
                {
                System.Diagnostics.Process.Start("explorer.exe", $"/select,\"{selectedImage.FilePath}\"");
                }
            catch (Exception ex)
                {
                HandyControl.Controls.MessageBox.Error($"打开资源管理器时出错：{ex.Message}", "错误");
                }
            }

        private void Image_MouseMove(object sender, System.Windows.Input.MouseEventArgs e)
            {
            if (e.LeftButton == System.Windows.Input.MouseButtonState.Pressed && sender is FrameworkElement element)
                {
                if (element.DataContext is ImageItem imageItem)
                    {
                    var data = new DataObject(DataFormats.FileDrop, new[] { imageItem.FilePath });
                    System.Windows.DragDrop.DoDragDrop(element, data, System.Windows.DragDropEffects.Copy);
                    }
                }
            }

        private void CategoryList_DragEnter(object sender, System.Windows.DragEventArgs e)
            {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
                {
                e.Effects = System.Windows.DragDropEffects.Copy;
                }
            else
                {
                e.Effects = System.Windows.DragDropEffects.None;
                }
            e.Handled = true;
            }

        private void CategoryList_Drop(object sender, System.Windows.DragEventArgs e)
            {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
                {
                var files = (string[])e.Data.GetData(DataFormats.FileDrop);
                if (files != null && files.Length > 0)
                    {
                    var targetElement = e.OriginalSource as FrameworkElement;
                    var categoryItem = targetElement?.DataContext as CategoryItem;

                    if (categoryItem != null)
                        {
                        try
                            {
                            var categoryPath = categoryItem.Path;
                            if (!Directory.Exists(categoryPath))
                                {
                                Directory.CreateDirectory(categoryPath);
                                }

                            foreach (var file in files)
                                {
                                try
                                    {
                                    // 移动文件到分类文件夹
                                    var fileName = Path.GetFileName(file);
                                    var newPath = Path.Combine(categoryPath, fileName);

                                    // 如果目标文件已存在，添加数字后缀
                                    int counter = 1;
                                    string fileNameWithoutExt = Path.GetFileNameWithoutExtension(fileName);
                                    string extension = Path.GetExtension(fileName);
                                    while (File.Exists(newPath))
                                        {
                                        fileName = $"{fileNameWithoutExt}_{counter}{extension}";
                                        newPath = Path.Combine(categoryPath, fileName);
                                        counter++;
                                        }

                                    File.Move(file, newPath);

                                    // 更新数据库
                                    var imageInfo = _dbService.GetImageByPath(file);
                                    if (imageInfo != null)
                                        {
                                        imageInfo.FilePath = newPath;
                                        imageInfo.FileName = fileName;
                                        imageInfo.Category = categoryItem.Name;
                                        _dbService.UpsertImage(imageInfo);
                                        }

                                    // 从当前显示中移除图片
                                    var imageToRemove = Images.FirstOrDefault(img => img.FilePath == file);
                                    if (imageToRemove != null)
                                        {
                                        Images.Remove(imageToRemove);
                                        }
                                    }
                                catch (Exception ex)
                                    {
                                    HandyControl.Controls.Growl.Error($"移动文件失败: {Path.GetFileName(file)} - {ex.Message}");
                                    }
                                }

                            // 更新分类的图片计数
                            var imageCount = Directory.GetFiles(categoryPath)
                                .Count(f => supportedExtensions.Contains(Path.GetExtension(f).ToLower()));
                            categoryItem.ImageCount = imageCount;

                            // 更新数据库中的分类信息
                            var categoryInfo = _dbService.GetCategory(categoryItem.Name);
                            if (categoryInfo != null)
                                {
                                categoryInfo.ImageCount = imageCount;
                                _dbService.UpsertCategory(categoryInfo);
                                }

                            // 如果当前显示的是目标分类，刷新显示
                            if (CategoryTreeView.SelectedItem is TreeViewItem selectedItem &&
                                selectedItem.DataContext is CategoryItem selectedCategory &&
                                selectedCategory.Name == categoryItem.Name)
                                {
                                FilterImagesByCategory(categoryItem);
                                }

                            HandyControl.Controls.Growl.Success($"已将 {files.Length} 张图片移动到分类 {categoryItem.Name}");
                            }
                        catch (Exception ex)
                            {
                            HandyControl.Controls.MessageBox.Error($"移动文件时出错：{ex.Message}", "错误");
                            }
                        }
                    }
                }
            }

        private void InitializeImagePaths()
            {
            ImagePaths = new ObservableCollection<PathItem>();
            PathListBox.ItemsSource = ImagePaths;

            // 初始化设置中的路径集合
            if (Properties.Settings.Default.ImageGalleryPaths == null)
                {
                Properties.Settings.Default.ImageGalleryPaths = new System.Collections.Specialized.StringCollection();
                Properties.Settings.Default.Save();
                }

            // 加载保存的路径列表
            foreach (string path in Properties.Settings.Default.ImageGalleryPaths)
                {
                if (!string.IsNullOrEmpty(path))
                    {
                    ImagePaths.Add(new PathItem { Path = path });
                    }
                }
            }

        private void SaveImagePaths()
            {
            var paths = new System.Collections.Specialized.StringCollection();
            foreach (var pathItem in ImagePaths)
                {
                if (!string.IsNullOrEmpty(pathItem.Path))
                    {
                    paths.Add(pathItem.Path);
                    }
                }
            Properties.Settings.Default.ImageGalleryPaths = paths;
            Properties.Settings.Default.Save();
            }

        private async Task LoadImagesFromAllPaths()
            {
            try
                {
                if (Application.Current?.Dispatcher == null) return;

                await Application.Current.Dispatcher.InvokeAsync(() =>
                {
                    LoadingMask.Visibility = Visibility.Visible;
                    LoadingProgress.Value = 0;
                    Images.Clear();
                });

                var allFiles = new List<string>();

                // 从当前路径加载
                if (!string.IsNullOrEmpty(currentFolderPath) && currentFolderPath != "ImageGalleryUrl")
                    {
                    allFiles.AddRange(Directory.GetFiles(currentFolderPath)
                        .Where(file => supportedExtensions.Contains(Path.GetExtension(file).ToLower())));
                    }

                // 从其他路径加载
                foreach (var pathItem in ImagePaths)
                    {
                    if (Directory.Exists(pathItem.Path))
                        {
                        allFiles.AddRange(Directory.GetFiles(pathItem.Path)
                            .Where(file => supportedExtensions.Contains(Path.GetExtension(file).ToLower())));
                        }
                    }

                var totalFiles = allFiles.Count;
                var loadedFiles = 0;

                foreach (var file in allFiles)
                    {
                    await LoadSingleImage(file, loadedFiles++, totalFiles);
                    }

                if (Application.Current?.Dispatcher != null)
                    {
                    await Application.Current.Dispatcher.InvokeAsync(() =>
                    {
                        LoadingMask.Visibility = Visibility.Collapsed;
                    });
                    }
                }
            catch (Exception ex)
                {
                if (Application.Current?.Dispatcher != null)
                    {
                    await Application.Current.Dispatcher.InvokeAsync(() =>
                    {
                        HandyControl.Controls.MessageBox.Error($"加载图片时出错：{ex.Message}", "错误");
                        LoadingMask.Visibility = Visibility.Collapsed;
                    });
                    }
                }
            }

        private async Task LoadSingleImage(string file, int loadedFiles, int totalFiles)
            {
            try
                {
                var fileInfo = new FileInfo(file);

                // 检查数据库中是否已存在
                var existingImage = _dbService.GetImageByPath(file);
                if (existingImage != null)
                    {
                    // 更新访问时间
                    existingImage.LastAccessTime = DateTime.Now;
                    _dbService.UpsertImage(existingImage);
                    }
                else
                    {
                    // 分析图片颜色
                    var dominantColors = ColorAnalyzer.AnalyzeImage(file);

                    // 创建新的图片信息
                    var imageInfo = new ImageInfo
                        {
                        FilePath = file,
                        FileName = Path.GetFileName(file),
                        FileSize = GetFileSizeString(fileInfo.Length),
                        CreationTime = fileInfo.CreationTime,
                        ModificationTime = fileInfo.LastWriteTime,
                        LastAccessTime = DateTime.Now,
                        ImportTime = DateTime.Now,
                        DominantColors = dominantColors
                        };

                    // 保存到数据库
                    _dbService.UpsertImage(imageInfo);
                    existingImage = imageInfo;
                    }

                if (Application.Current?.Dispatcher != null)
                    {
                    await Application.Current.Dispatcher.InvokeAsync(() =>
                    {
                        try
                            {
                            var image = new ImageItem
                                {
                                FilePath = existingImage.FilePath,
                                FileName = existingImage.FileName,
                                FileSize = existingImage.FileSize,
                                CreationTime = existingImage.CreationTime,
                                ModificationTime = existingImage.ModificationTime
                                };

                            var bitmap = new BitmapImage();
                            bitmap.BeginInit();
                            bitmap.UriSource = new Uri(file);
                            bitmap.CacheOption = BitmapCacheOption.OnLoad;
                            bitmap.DecodePixelWidth = 400;
                            bitmap.EndInit();
                            bitmap.Freeze();
                            image.Width = bitmap.PixelWidth;
                            image.Height = bitmap.PixelHeight;
                            image.Thumbnail = bitmap;

                            // 更新数据库中的尺寸信息
                            if (existingImage.Width != image.Width || existingImage.Height != image.Height)
                                {
                                existingImage.Width = image.Width;
                                existingImage.Height = image.Height;
                                _dbService.UpsertImage(existingImage);
                                }

                            Images.Add(image);

                            // 更新进度
                            var progress = (double)loadedFiles / totalFiles * 100;
                            LoadingProgress.Value = progress;
                            LoadingText.Text = $"正在加载图片... ({loadedFiles}/{totalFiles})";
                            }
                        catch (Exception)
                            {
                            HandyControl.Controls.Growl.Warning($"加载图片失败: {Path.GetFileName(file)}");
                            }
                    });
                    }
                }
            catch (Exception ex)
                {
                if (Application.Current?.Dispatcher != null)
                    {
                    await Application.Current.Dispatcher.InvokeAsync(() =>
                    {
                        HandyControl.Controls.Growl.Warning($"处理图片失败: {Path.GetFileName(file)} - {ex.Message}");
                    });
                    }
                }
            }

        public class PathItem
            {
            public string Path { get; set; }
            }

        private void ValidatePath_Click(object sender, RoutedEventArgs e)
            {
            if (string.IsNullOrEmpty(currentFolderPath))
                {
                UpdatePathStatus(false, "未设置路径");
                return;
                }

            bool isValid = ValidatePath(currentFolderPath);
            UpdatePathStatus(isValid, isValid ? "路径有效" : "路径无效");
            }

        private bool ValidatePath(string path)
            {
            if (string.IsNullOrEmpty(path)) return false;

            try
                {
                if (!Directory.Exists(path))
                    {
                    return false;
                    }

                // 检查是否有图片文件
                var hasImages = Directory.GetFiles(path)
                    .Any(file => supportedExtensions.Contains(Path.GetExtension(file).ToLower()));

                return hasImages;
                }
            catch (Exception)
                {
                return false;
                }
            }

        private void UpdatePathStatus(bool isValid, string message)
            {
            PathStatusIcon.Kind = isValid ?
                PackIconBootstrapIconsKind.CheckCircleFill :
                PackIconBootstrapIconsKind.ExclamationCircleFill;

            PathStatusIcon.Foreground = isValid ?
                FindResource("SuccessBrush") as Brush :
                FindResource("DangerBrush") as Brush;

            PathStatusText.Text = message;
            PathStatusText.Foreground = PathStatusIcon.Foreground;
            }

        private void ResetPath_Click(object sender, RoutedEventArgs e)
            {
            var result = HandyControl.Controls.MessageBox.Show(
                "确定要重置图库路径吗？这将清除当前路径设置。",
                "重置确认",
                MessageBoxButton.YesNo,
                MessageBoxImage.Warning);

            if (result == MessageBoxResult.Yes)
                {
                currentFolderPath = null;
                Properties.Settings.Default.ImageGalleryUrl = "ImageGalleryUrl";
                Properties.Settings.Default.Save();

                FolderPathText.Text = "未设置";
                UpdatePathStatus(false, "未设置路径");
                Images.Clear();

                HandyControl.Controls.Growl.Success("已重置图库路径");
                }
            }

        private void AddPath_Click(object sender, RoutedEventArgs e)
            {
            using (var dialog = new FolderBrowserDialog())
                {
                dialog.Description = "选择图片文件夹";
                if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                    {
                    var path = dialog.SelectedPath;
                    if (ValidatePath(path))
                        {
                        if (!ImagePaths.Any(p => p.Path == path))
                            {
                            ImagePaths.Add(new PathItem { Path = path });
                            SaveImagePaths();
                            HandyControl.Controls.Growl.Success("已添加新路径");
                            }
                        else
                            {
                            HandyControl.Controls.Growl.Warning("该路径已存在");
                            }
                        }
                    else
                        {
                        HandyControl.Controls.Growl.Error("选择的路径无效或不包含图片文件");
                        }
                    }
                }
            }

        private void RemovePath_Click(object sender, RoutedEventArgs e)
            {
            if (sender is Button button && button.DataContext is PathItem pathItem)
                {
                ImagePaths.Remove(pathItem);
                SaveImagePaths();
                HandyControl.Controls.Growl.Success("已移除路径");
                }
            }

        private void OpenSettings_Click(object sender, RoutedEventArgs e)
            {
            SettingsDrawer.IsOpen = true;
            }

        private void CloseDrawer_Click(object sender, RoutedEventArgs e)
            {
            SettingsDrawer.IsOpen = false;
            }

        private void CategoryTreeView_SelectedItemChanged(object sender, RoutedPropertyChangedEventArgs<object> e)
            {
            try
                {
                // 如果新值为空，直接返回
                if (e.NewValue == null) return;

                // 如果是TreeViewItem
                if (e.NewValue is TreeViewItem treeViewItem)
                    {
                    // 只处理主文件夹项
                    if (treeViewItem == RootFolderItem && !string.IsNullOrEmpty(currentFolderPath))
                        {
                        _ = LoadImagesFromFolder(currentFolderPath);
                        }
                    }
                // 如果是CategoryItem
                else if (e.NewValue is CategoryItem categoryItem)
                    {
                    // 防止重复加载
                    if (e.OldValue != e.NewValue)
                        {
                        FilterImagesByCategory(categoryItem);
                        }
                    }
                }
            catch (Exception ex)
                {
                HandyControl.Controls.MessageBox.Error($"切换分类时出错：{ex.Message}", "错误");
                }
            }

        private void UpdateCommonTags()
            {
            try
                {
                // 获取使用次数最多的前10个标签
                var commonTags = _dbService.GetAllTags()
                    .OrderByDescending(t => t.ImageCount)
                    .Take(10)
                    .Select(t => new TagItem
                        {
                        Name = t.Name,
                        Color = new SolidColorBrush((Color)ColorConverter.ConvertFromString(t.ColorHex))
                        });

                CommonTags.Clear();
                foreach (var tag in commonTags)
                    {
                    CommonTags.Add(tag);
                    }
                }
            catch (Exception ex)
                {
                HandyControl.Controls.Growl.Error($"更新常用标签时出错：{ex.Message}");
                }
            }

        private void ImageTagInput_KeyDown(object sender, System.Windows.Input.KeyEventArgs e)
            {
            if (e.Key == System.Windows.Input.Key.Enter)
                {
                AddTagToImage_Click(sender, null);
                }
            }

        private void CommonTag_Click(object sender, System.Windows.Input.MouseButtonEventArgs e)
            {
            if (sender is HandyControl.Controls.Tag tag)
                {
                var tagItem = tag.DataContext as TagItem;
                if (tagItem != null)
                    {
                    ImageTagInput.Text = tagItem.Name;
                    AddTagToImage_Click(sender, null);
                    }
                }
            }

        private void ManageTags_Click(object sender, RoutedEventArgs e)
            {
            // TODO: 实现标签管理窗口
            HandyControl.Controls.MessageBox.Info("标签管理功能正在开发中...", "提示");
            }

        private void Tag_Closing(object sender, CancelEventArgs e)
            {
            if (sender is HandyControl.Controls.Tag tag &&
                tag.DataContext is TagItem tagItem &&
                selectedImage != null)
                {
                try
                    {
                    // 从数据库中获取图片信息
                    var imageInfo = _dbService.GetImageByPath(selectedImage.FilePath);
                    if (imageInfo != null)
                        {
                        // 只从当前图片中移除标签
                        imageInfo.Tags.Remove(tagItem.Name);
                        _dbService.UpsertImage(imageInfo);

                        // 更新标签使用计数
                        _dbService.UpdateTagImageCount(tagItem.Name);

                        // 从UI中移除标签
                        selectedImage.Tags.Remove(tagItem);

                        // 更新UI显示
                        TagsItemsControl.ItemsSource = null;
                        TagsItemsControl.ItemsSource = selectedImage.Tags;

                        // 更新常用标签和筛选标签列表
                        UpdateCommonTags();
                        UpdateFilterTags();

                        HandyControl.Controls.Growl.Success($"已从图片中移除标签：{tagItem.Name}");
                        }
                    }
                catch (Exception ex)
                    {
                    HandyControl.Controls.MessageBox.Error($"移除标签时出错：{ex.Message}", "错误");
                    e.Cancel = true;  // 如果发生错误，取消关闭操作
                    }
                }
            }

        public class TagItem
            {
            public string Name { get; set; }
            public Brush Color { get; set; }
            }

        public class FilterTagItem : INotifyPropertyChanged
            {
            private string name;
            private Brush color;
            private bool isSelected;

            public string Name
                {
                get => name;
                set
                    {
                    name = value;
                    OnPropertyChanged(nameof(Name));
                    }
                }

            public Brush Color
                {
                get => color;
                set
                    {
                    color = value;
                    OnPropertyChanged(nameof(Color));
                    }
                }

            public bool IsSelected
                {
                get => isSelected;
                set
                    {
                    isSelected = value;
                    OnPropertyChanged(nameof(IsSelected));
                    }
                }

            public event PropertyChangedEventHandler PropertyChanged;

            protected virtual void OnPropertyChanged(string propertyName)
                {
                PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
                }
            }

        public class CategoryItem
            {
            public string Name { get; set; }
            public string Path { get; set; }
            public int ImageCount { get; set; }
            }

        public class ImageItem : INotifyPropertyChanged
            {
            private string filePath;
            private string fileName;
            private string fileSize;
            private int width;
            private int height;
            private DateTime creationTime;
            private DateTime modificationTime;
            private BitmapImage thumbnail;
            private ObservableCollection<TagItem> tags;
            private bool isSelected;
            private CategoryItem category;

            public string FilePath
                {
                get => filePath;
                set
                    {
                    filePath = value;
                    OnPropertyChanged(nameof(FilePath));
                    }
                }

            public string FileName
                {
                get => fileName;
                set
                    {
                    fileName = value;
                    OnPropertyChanged(nameof(FileName));
                    }
                }

            public string FileSize
                {
                get => fileSize;
                set
                    {
                    fileSize = value;
                    OnPropertyChanged(nameof(FileSize));
                    }
                }

            public int Width
                {
                get => width;
                set
                    {
                    width = value;
                    OnPropertyChanged(nameof(Width));
                    OnPropertyChanged(nameof(Dimensions));
                    }
                }

            public int Height
                {
                get => height;
                set
                    {
                    height = value;
                    OnPropertyChanged(nameof(Height));
                    OnPropertyChanged(nameof(Dimensions));
                    }
                }

            public string Dimensions => width > 0 && height > 0 ? $"{width:N0} × {height:N0}" : string.Empty;

            public DateTime CreationTime
                {
                get => creationTime;
                set
                    {
                    creationTime = value;
                    OnPropertyChanged(nameof(CreationTime));
                    }
                }

            public DateTime ModificationTime
                {
                get => modificationTime;
                set
                    {
                    modificationTime = value;
                    OnPropertyChanged(nameof(ModificationTime));
                    }
                }

            public BitmapImage Thumbnail
                {
                get => thumbnail;
                set
                    {
                    thumbnail = value;
                    OnPropertyChanged(nameof(Thumbnail));
                    }
                }

            public ObservableCollection<TagItem> Tags
                {
                get => tags;
                set
                    {
                    tags = value;
                    OnPropertyChanged(nameof(Tags));
                    }
                }

            public bool IsSelected
                {
                get => isSelected;
                set
                    {
                    isSelected = value;
                    OnPropertyChanged(nameof(IsSelected));
                    }
                }

            public CategoryItem Category
                {
                get => category;
                set
                    {
                    category = value;
                    OnPropertyChanged(nameof(Category));
                    }
                }

            public ImageItem()
                {
                Tags = new ObservableCollection<TagItem>();
                }

            public event PropertyChangedEventHandler PropertyChanged;

            protected virtual void OnPropertyChanged(string propertyName)
                {
                PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
                }
            }

        private void SearchBar_TextChanged(object sender, TextChangedEventArgs e)
            {
            if (sender is HandyControl.Controls.SearchBar searchBar)
                {
                PerformSearch(searchBar.Text);
                }
            }

        private void ColorFilter_Click(object sender, System.Windows.Input.MouseButtonEventArgs e)
            {
            if (sender is Border colorBorder)
                {
                // 获取边框的背景色
                if (colorBorder.Background is SolidColorBrush brush)
                    {
                    // 设置所有颜色过滤器的边框为0
                    var parent = colorBorder.Parent as ItemsControl;
                    if (parent != null)
                        {
                        foreach (var item in parent.Items)
                            {
                            if (item is Border border)
                                {
                                border.BorderThickness = new Thickness(0);
                                }
                            }
                        }

                    // 设置当前选中的颜色过滤器的边框
                    colorBorder.BorderThickness = new Thickness(2);
                    colorBorder.BorderBrush = new SolidColorBrush(Colors.White);

                    // 执行颜色筛选
                    FilterByColor(brush.Color);
                    }
                }
            }

        private void InsertToPPT_Click(object sender, RoutedEventArgs e)
            {
            var menuItem = sender as MenuItem;
            var contextMenu = menuItem.Parent as ContextMenu;
            var card = contextMenu.PlacementTarget as Card;
            var imageItem = card.DataContext as ImageItem;

            try
                {
                var pptApp = Globals.ThisAddIn.Application;
                var activePresentation = pptApp.ActivePresentation;
                var activeSlide = pptApp.ActiveWindow.View.Slide;

                if (activeSlide != null)
                    {
                    // 获取图片路径
                    string imagePath = imageItem.FilePath;

                    // 计算插入位置（居中）
                    float slideWidth = activeSlide.Master.Width;
                    float slideHeight = activeSlide.Master.Height;
                    float left = (slideWidth - 300) / 2;  // 300为默认宽度
                    float top = (slideHeight - 200) / 2;  // 200为默认高度

                    // 插入图片
                    var shape = activeSlide.Shapes.AddPicture(
                        imagePath,
                        Microsoft.Office.Core.MsoTriState.msoFalse,
                        Microsoft.Office.Core.MsoTriState.msoTrue,
                        left, top, 300, 200);

                    // 选中新插入的图片
                    shape.Select();

                    // 显示成功提示
                    HandyControl.Controls.Growl.Success("图片已成功插入到PPT中");
                    }
                else
                    {
                    HandyControl.Controls.Growl.Warning("请先选择要插入的幻灯片");
                    }
                }
            catch (Exception ex)
                {
                HandyControl.Controls.Growl.Error($"插入图片失败: {ex.Message}");
                }
            }

        private void CopyImage_Click(object sender, RoutedEventArgs e)
            {
            var menuItem = sender as MenuItem;
            var contextMenu = menuItem.Parent as ContextMenu;
            var card = contextMenu.PlacementTarget as Card;
            var imageItem = card.DataContext as ImageItem;

            try
                {
                // 将图片复制到剪贴板
                var bitmap = new BitmapImage(new Uri(imageItem.FilePath));
                Clipboard.SetImage(bitmap);
                HandyControl.Controls.Growl.Success("图片已复制到剪贴板");
                }
            catch (Exception ex)
                {
                HandyControl.Controls.Growl.Error($"复制图片失败: {ex.Message}");
                }
            }
        }
    }