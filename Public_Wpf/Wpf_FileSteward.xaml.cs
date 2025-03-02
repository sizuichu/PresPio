using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Interop;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using HandyControl.Controls;
using Microsoft.Win32;
using Newtonsoft.Json;
using MessageBox = HandyControl.Controls.MessageBox;
using Window = System.Windows.Window;

namespace PresPio
    {
    public class FileSystemItem
        {
        public string Name { get; set; }
        public string Path { get; set; }
        public ImageSource IconPath { get; set; }
        public ObservableCollection<FileSystemItem> Children { get; set; }
        public bool IsDirectory { get; set; }
        public List<string> FileTypes { get; set; }

        public FileSystemItem()
            {
            Children = new ObservableCollection<FileSystemItem>();
            FileTypes = new List<string>();
            }
        }

    public class AppSettings
        {
        public List<FileLibrary> CustomLibraries { get; set; } = new List<FileLibrary>();
        public string LastSelectedLibrary { get; set; }
        public string DefaultViewMode { get; set; } = "List"; // List 或 Card
        }

    public partial class Wpf_FileSteward : Window
        {
        [DllImport("shell32.dll", CharSet = CharSet.Auto)]
        private static extern IntPtr SHGetFileInfo(string pszPath, uint dwFileAttributes, ref SHFILEINFO psfi, uint cbFileInfo, uint uFlags);

        [DllImport("user32.dll", SetLastError = true)]
        private static extern bool DestroyIcon(IntPtr hIcon);

        [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Auto)]
        private struct SHFILEINFO
            {
            public IntPtr hIcon;
            public int iIcon;
            public uint dwAttributes;

            [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 260)]
            public string szDisplayName;

            [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 80)]
            public string szTypeName;
            }

        private const uint SHGFI_ICON = 0x100;
        private const uint SHGFI_LARGEICON = 0x0;
        private const uint SHGFI_SMALLICON = 0x1;
        private const uint FILE_ATTRIBUTE_DIRECTORY = 0x10;

        private ObservableCollection<FileSystemItem> RootItems { get; set; }

        private readonly string settingsPath = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
            "PresPio",
            "FileSteward",
            "settings.json"
        );

        private AppSettings settings;

        private ObservableCollection<FileItem> FileItems { get; set; }
        private ObservableCollection<FileItem> FilteredItems { get; set; }
        private ObservableCollection<FileLibrary> FileLibraries { get; set; }
        private string currentLibraryPath;
        private string currentCategory = "全部";

        public Wpf_FileSteward()
            {
            InitializeComponent();
            LoadSettings();
            InitializeCollections();
            InitializeEventHandlers();
            LoadFileSystem();
            InitializeCategories();
            ApplySettings();
            }

        private void LoadSettings()
            {
            try
                {
                var settingsDir = Path.GetDirectoryName(settingsPath);
                if (!Directory.Exists(settingsDir))
                    {
                    Directory.CreateDirectory(settingsDir);
                    }

                if (File.Exists(settingsPath))
                    {
                    var json = File.ReadAllText(settingsPath);
                    settings = JsonConvert.DeserializeObject<AppSettings>(json);
                    }
                else
                    {
                    settings = new AppSettings();
                    SaveSettings();
                    }
                }
            catch (Exception ex)
                {
                MessageBox.Show($"加载设置时出错：{ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
                settings = new AppSettings();
                }
            }

        private void SaveSettings()
            {
            try
                {
                var json = JsonConvert.SerializeObject(settings, Formatting.Indented);
                File.WriteAllText(settingsPath, json);
                }
            catch (Exception ex)
                {
                MessageBox.Show($"保存设置时出错：{ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }

        private void LoadFileSystem()
            {
            RootItems = new ObservableCollection<FileSystemItem>();

            // 添加默认库
            var defaultLibrary = new FileSystemItem { Name = "默认文件库", IsDirectory = true };
            AddDefaultLibraries(defaultLibrary);
            RootItems.Add(defaultLibrary);

            // 添加自定义库
            var customLibrary = new FileSystemItem { Name = "自定义文件库", IsDirectory = true };
            LoadCustomLibraries(customLibrary);
            RootItems.Add(customLibrary);

            FileLibraryTree.ItemsSource = RootItems;
            }

        private void AddDefaultLibraries(FileSystemItem parent)
            {
            var documentsPath = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            var picturesPath = Environment.GetFolderPath(Environment.SpecialFolder.MyPictures);
            var videosPath = Environment.GetFolderPath(Environment.SpecialFolder.MyVideos);

            parent.Children.Add(CreateLibraryItem("文档", documentsPath, new[] { ".doc", ".docx", ".pdf", ".txt" }));
            parent.Children.Add(CreateLibraryItem("图片", picturesPath, new[] { ".jpg", ".jpeg", ".png", ".gif" }));
            parent.Children.Add(CreateLibraryItem("视频", videosPath, new[] { ".mp4", ".avi", ".mkv" }));
            }

        private FileSystemItem CreateLibraryItem(string name, string path, string[] fileTypes)
            {
            return new FileSystemItem
                {
                Name = name,
                Path = path,
                IsDirectory = true,
                FileTypes = fileTypes.ToList(),
                IconPath = GetFileIcon(path, true)
                };
            }

        private void LoadCustomLibraries(FileSystemItem parent)
            {
            foreach (var library in settings.CustomLibraries)
                {
                parent.Children.Add(CreateLibraryItem(library.Name, library.Path, library.FileTypes.ToArray()));
                }
            }

        private ImageSource GetFileIcon(string path, bool isDirectory)
            {
            try
                {
                SHFILEINFO shinfo = new SHFILEINFO();
                uint flags = SHGFI_ICON | (isDirectory ? SHGFI_LARGEICON : SHGFI_SMALLICON);
                uint attributes = isDirectory ? FILE_ATTRIBUTE_DIRECTORY : 0;

                IntPtr result = SHGetFileInfo(path, attributes, ref shinfo, (uint)Marshal.SizeOf(shinfo), flags);

                if (shinfo.hIcon == IntPtr.Zero)
                    return null;

                ImageSource imageSource = null;
                try
                    {
                    imageSource = Imaging.CreateBitmapSourceFromHIcon(
                        shinfo.hIcon,
                        Int32Rect.Empty,
                        BitmapSizeOptions.FromEmptyOptions());
                    }
                finally
                    {
                    DestroyIcon(shinfo.hIcon);
                    }

                return imageSource;
                }
            catch
                {
                return null;
                }
            }

        private void LoadFolderContents(string path, FileSystemItem parent)
            {
            try
                {
                parent.Children.Clear();

                // 加载子文件夹
                foreach (var dir in Directory.GetDirectories(path))
                    {
                    var dirInfo = new DirectoryInfo(dir);
                    var item = new FileSystemItem
                        {
                        Name = dirInfo.Name,
                        Path = dir,
                        IsDirectory = true,
                        IconPath = GetFileIcon(dir, true)
                        };
                    parent.Children.Add(item);
                    }

                // 加载文件
                foreach (var file in Directory.GetFiles(path))
                    {
                    var fileInfo = new FileInfo(file);
                    var item = new FileSystemItem
                        {
                        Name = fileInfo.Name,
                        Path = file,
                        IsDirectory = false,
                        IconPath = GetFileIcon(file, false)
                        };
                    parent.Children.Add(item);
                    }
                }
            catch (Exception ex)
                {
                MessageBox.Show($"加载文件夹内容时出错：{ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }

        private void FileLibraryTree_SelectedItemChanged(object sender, RoutedPropertyChangedEventArgs<object> e)
            {
            if (e.NewValue is FileSystemItem item)
                {
                if (item.IsDirectory)
                    {
                    currentLibraryPath = item.Path;
                    if (item.Children.Count == 0)
                        {
                        LoadFolderContents(item.Path, item);
                        }
                    UpdateFileList(item.Path);
                    }
                }
            }

        private void ViewMode_SelectionChanged(object sender, SelectionChangedEventArgs e)
            {
            if (cmbViewMode.SelectedItem is ComboBoxItem selectedItem)
                {
                switch (selectedItem.Content.ToString())
                    {
                    case "列表":
                        FileListView.Visibility = Visibility.Visible;
                        CardListView.Visibility = Visibility.Collapsed;
                        WaterfallView.Visibility = Visibility.Collapsed;
                        break;

                    case "卡片":
                        FileListView.Visibility = Visibility.Collapsed;
                        CardListView.Visibility = Visibility.Visible;
                        WaterfallView.Visibility = Visibility.Collapsed;
                        break;

                    case "���布流":
                        FileListView.Visibility = Visibility.Collapsed;
                        CardListView.Visibility = Visibility.Collapsed;
                        WaterfallView.Visibility = Visibility.Visible;
                        break;
                    }
                }
            }

        private void ApplySettings()
            {
            // 应用视图模式
            switch (settings.DefaultViewMode.ToLower())
                {
                case "card":
                    cmbViewMode.SelectedIndex = 1; // 卡片视图
                    break;

                case "waterfall":
                    cmbViewMode.SelectedIndex = 2; // 瀑布流视图
                    break;

                default:
                    cmbViewMode.SelectedIndex = 0; // 列表视图
                    break;
                }
            }

        private void SaveCurrentState()
            {
            // 保存当前视图模式
            if (cmbViewMode.SelectedItem is ComboBoxItem selectedItem)
                {
                switch (selectedItem.Content.ToString())
                    {
                    case "卡片":
                        settings.DefaultViewMode = "Card";
                        break;

                    case "瀑布流":
                        settings.DefaultViewMode = "Waterfall";
                        break;

                    default:
                        settings.DefaultViewMode = "List";
                        break;
                    }
                }

            // 保存当前选中的库
            if (FileLibraryTree.SelectedItem is FileSystemItem selectedLibrary)
                {
                settings.LastSelectedLibrary = selectedLibrary.Name;
                }

            SaveSettings();
            }

        protected override void OnClosing(CancelEventArgs e)
            {
            SaveCurrentState();
            base.OnClosing(e);
            }

        private void InitializeCollections()
            {
            FileItems = new ObservableCollection<FileItem>();
            FilteredItems = new ObservableCollection<FileItem>();
            FileLibraries = new ObservableCollection<FileLibrary>();
            FileListView.ItemsSource = FilteredItems;
            CardListView.ItemsSource = FilteredItems;
            }

        private void InitializeCategories()
            {
            var commonCategories = new Dictionary<string, string[]>
            {
                { "文档", new[] { ".doc", ".docx", ".pdf", ".txt", ".rtf" } },
                { "图片", new[] { ".jpg", ".jpeg", ".png", ".gif", ".bmp" } },
                { "视频", new[] { ".mp4", ".avi", ".mkv", ".mov", ".wmv" } },
                { "音频", new[] { ".mp3", ".wav", ".flac", ".m4a", ".wma" } },
                { "压缩包", new[] { ".zip", ".rar", ".7z", ".tar", ".gz" } }
            };

            foreach (var category in commonCategories)
                {
                var button = new Button
                    {
                    Content = category.Key,
                    Margin = new Thickness(0, 0, 5, 0),
                    Style = Application.Current.FindResource("ButtonDefault") as Style
                    };
                button.Click += CategoryButton_Click;
                CategoryPanel.Children.Add(button);
                }
            }

        private void CategoryButton_Click(object sender, RoutedEventArgs e)
            {
            var button = sender as Button;
            if (button != null)
                {
                // 重置所有按钮样式
                foreach (Button btn in CategoryPanel.Children.OfType<Button>())
                    {
                    btn.Style = Application.Current.FindResource("ButtonDefault") as Style;
                    }

                // 设置选中按钮样式
                button.Style = Application.Current.FindResource("ButtonPrimary") as Style;
                currentCategory = button.Content.ToString();

                // 应用筛选
                ApplyFilter();
                }
            }

        private void ApplyFilter()
            {
            FilteredItems.Clear();
            var items = FileItems.ToList();

            // 应用分类筛选
            if (currentCategory != "全部")
                {
                var extensions = GetCategoryExtensions(currentCategory);
                items = items.Where(f => extensions.Contains(f.FileType.ToLower())).ToList();
                }

            // 应用排序
            if (cmbSort.SelectedItem != null)
                {
                var sortOption = (cmbSort.SelectedItem as ComboBoxItem).Content.ToString();
                switch (sortOption)
                    {
                    case "按名称":
                        items = items.OrderBy(f => f.FileName).ToList();
                        break;

                    case "按大小":
                        items = items.OrderByDescending(f => GetFileSizeInBytes(f.FileSize)).ToList();
                        break;

                    case "按类型":
                        items = items.OrderBy(f => f.FileType).ToList();
                        break;

                    case "按修改时间":
                        items = items.OrderByDescending(f => f.ModifiedDate).ToList();
                        break;
                    }
                }

            foreach (var item in items)
                {
                FilteredItems.Add(item);
                }
            UpdateFileCount();
            }

        private string[] GetCategoryExtensions(string category)
            {
            switch (category)
                {
                case "文档":
                    return new[] { ".doc", ".docx", ".pdf", ".txt", ".rtf" };

                case "图片":
                    return new[] { ".jpg", ".jpeg", ".png", ".gif", ".bmp" };

                case "视频":
                    return new[] { ".mp4", ".avi", ".mkv", ".mov", ".wmv" };

                case "音频":
                    return new[] { ".mp3", ".wav", ".flac", ".m4a", ".wma" };

                case "压缩包":
                    return new[] { ".zip", ".rar", ".7z", ".tar", ".gz" };

                default:
                    return new string[] { };
                }
            }

        private void CmbSort_SelectionChanged(object sender, SelectionChangedEventArgs e)
            {
            ApplyFilter();
            }

        private void FileListView_SelectionChanged(object sender, SelectionChangedEventArgs e)
            {
            // 空方法，保留以备后用
            }

        private void CardListView_SelectionChanged(object sender, SelectionChangedEventArgs e)
            {
            // 空方法，保留以备后用
            }

        private void InitializeEventHandlers()
            {
            FileLibraryTree.SelectedItemChanged += FileLibraryTree_SelectedItemChanged;
            cmbViewMode.SelectionChanged += ViewMode_SelectionChanged;
            cmbSort.SelectionChanged += CmbSort_SelectionChanged;
            }

        private void SearchBar_OnSearchStarted(object sender, HandyControl.Data.FunctionEventArgs<string> e)
            {
            string searchText = e.Info?.ToLower();
            if (string.IsNullOrEmpty(searchText))
                {
                UpdateFileList(currentLibraryPath);
                return;
                }

            var searchResults = FileItems.Where(f =>
                f.FileName.ToLower().Contains(searchText) ||
                f.FilePath.ToLower().Contains(searchText)).ToList();

            FileItems.Clear();
            foreach (var item in searchResults)
                {
                FileItems.Add(item);
                }
            UpdateFileCount();
            }

        private void UpdateFileList(string path)
            {
            if (string.IsNullOrEmpty(path) || !Directory.Exists(path))
                return;

            FileItems.Clear();
            FilteredItems.Clear();
            try
                {
                var files = Directory.GetFiles(path, "*.*", SearchOption.AllDirectories);
                foreach (var file in files)
                    {
                    var fileInfo = new FileInfo(file);
                    var item = new FileItem
                        {
                        FileName = fileInfo.Name,
                        FileSize = FormatFileSize(fileInfo.Length),
                        FileType = fileInfo.Extension,
                        ModifiedDate = fileInfo.LastWriteTime,
                        FilePath = fileInfo.FullName,
                        IconPath = GetFileIcon(file, false)
                        };
                    FileItems.Add(item);
                    }
                ApplyFilter();
                }
            catch (Exception ex)
                {
                MessageBox.Show($"读取文件列表时出错：{ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }

        private ImageSource GeneratePreviewImage(FileInfo fileInfo)
            {
            try
                {
                switch (fileInfo.Extension.ToLower())
                    {
                    case ".jpg":
                    case ".jpeg":
                    case ".png":
                    case ".bmp":
                        var bitmap = new BitmapImage();
                        bitmap.BeginInit();
                        bitmap.UriSource = new Uri(fileInfo.FullName);
                        bitmap.DecodePixelWidth = 200; // 限制预览图大小
                        bitmap.CacheOption = BitmapCacheOption.OnLoad;
                        bitmap.EndInit();
                        return bitmap;

                    default:
                        return null;
                    }
                }
            catch
                {
                return null;
                }
            }

        private void UpdateFileCount()
            {
            int fileCount = FileItems.Count;
            long totalSize = FileItems.Sum(f => GetFileSizeInBytes(f.FileSize));
            var fileTypes = FileItems.GroupBy(f => f.FileType.ToLower())
                                   .ToDictionary(g => g.Key, g => g.Count());

            // 构建文件类型统计信息
            var fileTypeSummary = fileTypes.Count > 0
                ? $"包含: {string.Join(", ", fileTypes.Select(ft => $"{ft.Key}({ft.Value})"))})"
                : string.Empty;

            txtFileCount.Text = $"共 {fileCount} 个文件";
            txtFileSize.Text = $"总大小: {FormatFileSize(totalSize)} {fileTypeSummary}";
            }

        private long GetFileSizeInBytes(string formattedSize)
            {
            try
                {
                var parts = formattedSize.Split(' ');
                if (parts.Length != 2) return 0;

                double size = double.Parse(parts[0]);
                string unit = parts[1].ToUpper();

                switch (unit)
                    {
                    case "B": return (long)size;
                    case "KB": return (long)(size * 1024);
                    case "MB": return (long)(size * 1024 * 1024);
                    case "GB": return (long)(size * 1024 * 1024 * 1024);
                    case "TB": return (long)(size * 1024 * 1024 * 1024 * 1024);
                    default: return 0;
                    }
                }
            catch
                {
                return 0;
                }
            }

        private string FormatFileSize(long bytes)
            {
            string[] sizes = { "B", "KB", "MB", "GB", "TB" };
            int order = 0;
            double len = bytes;
            while (len >= 1024 && order < sizes.Length - 1)
                {
                order++;
                len = len / 1024;
                }
            return $"{len:0.##} {sizes[order]}";
            }

        #region 文件操作

        private void OpenFile_Click(object sender, RoutedEventArgs e)
            {
            var selectedItem = FileListView.SelectedItem as FileItem;
            if (selectedItem != null)
                {
                try
                    {
                    Process.Start(selectedItem.FilePath);
                    }
                catch (Exception ex)
                    {
                    MessageBox.Show($"打开文件时出错：{ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
                    }
                }
            }

        private void OpenFileLocation_Click(object sender, RoutedEventArgs e)
            {
            var selectedItem = FileListView.SelectedItem as FileItem;
            if (selectedItem != null)
                {
                Process.Start("explorer.exe", $"/select,\"{selectedItem.FilePath}\"");
                }
            }

        private void CopyFile_Click(object sender, RoutedEventArgs e)
            {
            var selectedItem = FileListView.SelectedItem as FileItem;
            if (selectedItem != null)
                {
                var dialog = new SaveFileDialog
                    {
                    FileName = selectedItem.FileName,
                    Filter = "All files (*.*)|*.*"
                    };

                if (dialog.ShowDialog() == true)
                    {
                    try
                        {
                        File.Copy(selectedItem.FilePath, dialog.FileName, true);
                        Growl.Success("文件复制成功！");
                        }
                    catch (Exception ex)
                        {
                        MessageBox.Show($"复制文件时出错：{ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
                        }
                    }
                }
            }

        private void MoveFile_Click(object sender, RoutedEventArgs e)
            {
            var selectedItem = FileListView.SelectedItem as FileItem;
            if (selectedItem != null)
                {
                var dialog = new SaveFileDialog
                    {
                    FileName = selectedItem.FileName,
                    Filter = "All files (*.*)|*.*"
                    };

                if (dialog.ShowDialog() == true)
                    {
                    try
                        {
                        File.Move(selectedItem.FilePath, dialog.FileName);
                        FileItems.Remove(selectedItem);
                        UpdateFileCount();
                        Growl.Success("文件移动成功！");
                        }
                    catch (Exception ex)
                        {
                        MessageBox.Show($"移动文件时出错：{ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
                        }
                    }
                }
            }

        private void DeleteFile_Click(object sender, RoutedEventArgs e)
            {
            var selectedItem = FileListView.SelectedItem as FileItem;
            if (selectedItem != null)
                {
                var result = MessageBox.Show(
                    $"确定要删除文件 {selectedItem.FileName} ？",
                    "确认删除",
                    MessageBoxButton.YesNo,
                    MessageBoxImage.Question);

                if (result == MessageBoxResult.Yes)
                    {
                    try
                        {
                        File.Delete(selectedItem.FilePath);
                        FileItems.Remove(selectedItem);
                        UpdateFileCount();
                        Growl.Success("文件删除成功！");
                        }
                    catch (Exception ex)
                        {
                        MessageBox.Show($"删除文件时出错：{ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
                        }
                    }
                }
            }

        private void FileProperties_Click(object sender, RoutedEventArgs e)
            {
            var selectedItem = FileListView.SelectedItem as FileItem;
            if (selectedItem != null)
                {
                try
                    {
                    Process.Start("explorer.exe", $"/properties \"{selectedItem.FilePath}\"");
                    }
                catch (Exception ex)
                    {
                    MessageBox.Show($"显示属性时出错：{ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
                    }
                }
            }

        private void FileListView_Drop(object sender, DragEventArgs e)
            {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
                {
                string[] files = (string[])e.Data.GetData(DataFormats.FileDrop);
                foreach (string file in files)
                    {
                    if (File.Exists(file))
                        {
                        var fileInfo = new FileInfo(file);
                        FileItems.Add(new FileItem
                            {
                            FileName = fileInfo.Name,
                            FileSize = FormatFileSize(fileInfo.Length),
                            FileType = fileInfo.Extension,
                            ModifiedDate = fileInfo.LastWriteTime,
                            FilePath = fileInfo.FullName,
                            IconPath = GetFileIcon(file, false)
                            });
                        }
                    }
                UpdateFileCount();
                }
            }

        private void FileListView_PreviewDragOver(object sender, DragEventArgs e)
            {
            e.Effects = DragDropEffects.Copy;
            e.Handled = true;
            }

        #endregion 文件操作

        #region 文件库操作

        private void AddLibrary_Click(object sender, RoutedEventArgs e)
            {
            var dialog = new System.Windows.Forms.FolderBrowserDialog();
            if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                var library = new FileLibrary
                    {
                    Name = new DirectoryInfo(dialog.SelectedPath).Name,
                    Path = dialog.SelectedPath,
                    FileTypes = new List<string>()
                    };
                FileLibraries.Add(library);

                // 更新TreeView
                var item = new FileSystemItem
                    {
                    Name = library.Name,
                    Path = library.Path,
                    IsDirectory = true,
                    IconPath = GetFileIcon(library.Path, true)
                    };

                if (RootItems.Count > 1 && RootItems[1].Name == "自定义文件库")
                    {
                    RootItems[1].Children.Add(item);
                    }

                // 保存到设置
                settings.CustomLibraries = FileLibraries
                    .Where(l => !IsDefaultLibrary(l.Name))
                    .ToList();
                SaveSettings();

                Growl.Success("文件库添加成功！");
                }
            }

        private bool IsDefaultLibrary(string name)
            {
            return name == "文档" || name == "图片" || name == "视频";
            }

        private void EditLibrary_Click(object sender, RoutedEventArgs e)
            {
            // TODO: 实现编辑文件库功能
            }

        private void DeleteLibrary_Click(object sender, RoutedEventArgs e)
            {
            if (FileLibraryTree.SelectedItem is FileSystemItem item && !IsDefaultLibrary(item.Name))
                {
                var result = MessageBox.Show(
                    $"确定要删除文件库 {item.Name} 吗？\n注意：这不会删除实际的文件夹。",
                    "确认删除",
                    MessageBoxButton.YesNo,
                    MessageBoxImage.Question);

                if (result == MessageBoxResult.Yes)
                    {
                    var library = FileLibraries.FirstOrDefault(l => l.Name == item.Name);
                    if (library != null)
                        {
                        FileLibraries.Remove(library);
                        }

                    if (RootItems.Count > 1 && RootItems[1].Name == "自定义文件库")
                        {
                        RootItems[1].Children.Remove(item);
                        }

                    // 从设置中移除
                    settings.CustomLibraries.RemoveAll(l => l.Name == item.Name);
                    SaveSettings();

                    Growl.Success("文件库删除成功！");
                    }
                }
            else
                {
                MessageBox.Show("默认文件库不能删除！", "提示", MessageBoxButton.OK, MessageBoxImage.Information);
                }
            }

        private void Refresh_Click(object sender, RoutedEventArgs e)
            {
            UpdateFileList(currentLibraryPath);
            }

        #endregion 文件库操作
        }

    public class FileItem : INotifyPropertyChanged
        {
        public string FileName { get; set; }
        public string FileSize { get; set; }
        public string FileType { get; set; }
        public DateTime ModifiedDate { get; set; }
        public string FilePath { get; set; }
        public ImageSource IconPath { get; set; }

        public event PropertyChangedEventHandler PropertyChanged;

        protected virtual void OnPropertyChanged(string propertyName)
            {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
            }
        }

    public class FileLibrary
        {
        public string Name { get; set; }
        public string Path { get; set; }
        public List<string> FileTypes { get; set; }
        }
    }