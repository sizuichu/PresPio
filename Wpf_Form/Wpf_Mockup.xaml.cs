using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Forms;
using HandyControl.Controls;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using MessageBox = System.Windows.Forms.MessageBox;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using System.Xml.Linq;

namespace PresPio
    {
    public class MockupItem
        {
        public string Name { get; set; }
        public string Path { get; set; }
        public DateTime LastModified { get; set; }
        public bool IsDefault { get; set; }
        }

    public partial class Wpf_Mockup
        {
        public PowerPoint.Application app;
        private readonly string mockupFolderPath;
        private readonly string mockupConfigPath;
        private ObservableCollection<MockupItem> mockupItems;
        private ObservableCollection<MockupItem> filteredMockupItems;

        public Wpf_Mockup()
            {
            try
                {
                app = Globals.ThisAddIn.Application;
                app.WindowSelectionChange += new EApplication_WindowSelectionChangeEventHandler(Application_WindowSelectionChange);
                InitializeComponent();

                // 初始化文件夹路径
                mockupFolderPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "MocKup");
                mockupConfigPath = Path.Combine(mockupFolderPath, "mockups.xml");
                EnsureDirectoryExists();

                // 初始化集合
                mockupItems = new ObservableCollection<MockupItem>();
                filteredMockupItems = new ObservableCollection<MockupItem>();

                // 设置数据源
                MockupListView.ItemsSource = filteredMockupItems;

                LoadMockupList();
                LoadNum();
                LoadImg(); // 只加载预览图，不重新生成
                }
            catch (Exception ex)
                {
                Growl.ErrorGlobal($"初始化失败: {ex.Message}");
                }
            }

        private void SaveMockupConfig()
            {
            try
                {
                var doc = new XDocument(
                    new XElement("Mockups",
                        from item in mockupItems
                        select new XElement("Mockup",
                            new XElement("Name", item.Name),
                            new XElement("Path", item.Path),
                            new XElement("IsDefault", item.IsDefault)
                        )
                    )
                );
                doc.Save(mockupConfigPath);
                }
            catch (Exception ex)
                {
                Growl.ErrorGlobal($"保存样机配置失败: {ex.Message}");
                }
            }

        private void LoadMockupConfig()
            {
            try
                {
                if (File.Exists(mockupConfigPath))
                    {
                    var doc = XDocument.Load(mockupConfigPath);
                    foreach (var element in doc.Root.Elements("Mockup"))
                        {
                        string path = element.Element("Path").Value;
                        if (File.Exists(path))
                            {
                            mockupItems.Add(new MockupItem
                                {
                                Name = element.Element("Name").Value,
                                Path = path,
                                LastModified = File.GetLastWriteTime(path),
                                IsDefault = bool.Parse(element.Element("IsDefault").Value)
                                });

                            if (bool.Parse(element.Element("IsDefault").Value))
                                {
                                Properties.Settings.Default.MockupUrl = path;
                                Properties.Settings.Default.Save();
                                }
                            }
                        }
                    }
                }
            catch (Exception ex)
                {
                Growl.ErrorGlobal($"加载样机配置失败: {ex.Message}");
                }
            }

        private void LoadMockupList()
            {
            try
                {
                mockupItems.Clear();
                
                // 加载配置文件中的样机
                LoadMockupConfig();

                // 搜索MocKup文件夹中的新样机文件
                if (Directory.Exists(mockupFolderPath))
                    {
                    var mockupFiles = Directory.GetFiles(mockupFolderPath, "*.pptx");
                    foreach (var file in mockupFiles)
                        {
                        if (!mockupItems.Any(x => x.Path == file))
                            {
                            mockupItems.Add(new MockupItem
                                {
                                Name = Path.GetFileNameWithoutExtension(file),
                                Path = file,
                                LastModified = File.GetLastWriteTime(file),
                                IsDefault = false
                                });
                            }
                        }
                    }

                // 如果没有默认样机，设置第一个为默认
                if (!mockupItems.Any(x => x.IsDefault) && mockupItems.Any())
                    {
                    mockupItems.First().IsDefault = true;
                    Properties.Settings.Default.MockupUrl = mockupItems.First().Path;
                    Properties.Settings.Default.Save();
                    }

                SaveMockupConfig();
                UpdateFilteredList();
                }
            catch (Exception ex)
                {
                Growl.ErrorGlobal($"加载样机列表失败: {ex.Message}");
                }
            }

        private void UpdateFilteredList(string searchText = "")
            {
            filteredMockupItems.Clear();
            var query = mockupItems.AsEnumerable();

            if (!string.IsNullOrWhiteSpace(searchText))
                {
                searchText = searchText.ToLower();
                query = query.Where(item => 
                    item.Name.ToLower().Contains(searchText) || 
                    item.Path.ToLower().Contains(searchText));
                }

            foreach (var item in query.OrderByDescending(x => x.LastModified))
                {
                filteredMockupItems.Add(item);
                }
            }

        private void SearchBox_TextChanged(object sender, System.Windows.Controls.TextChangedEventArgs e)
            {
            UpdateFilteredList(SearchBox.Text);
            }

        private void AddMockup_Click(object sender, RoutedEventArgs e)
            {
            using (OpenFileDialog openFileDialog = new OpenFileDialog())
                {
                openFileDialog.InitialDirectory = mockupFolderPath;
                openFileDialog.Filter = "样机文件 (*.pptx)|*.pptx|All files (*.*)|*.*";
                openFileDialog.Title = "添加样机文件";

                if (openFileDialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                    {
                    string selectedPath = openFileDialog.FileName;
                    string fileName = Path.GetFileName(selectedPath);
                    string targetPath = Path.Combine(mockupFolderPath, fileName);

                    try
                        {
                        if (selectedPath != targetPath)
                            {
                            File.Copy(selectedPath, targetPath, true);
                            }

                        // 添加新样机
                        var newMockup = new MockupItem
                            {
                            Name = Path.GetFileNameWithoutExtension(targetPath),
                            Path = targetPath,
                            LastModified = File.GetLastWriteTime(targetPath),
                            IsDefault = !mockupItems.Any() // 如果是第一个样机，设为默认
                            };

                        mockupItems.Add(newMockup);
                        SaveMockupConfig();
                        LoadMockupList();
                        
                        if (newMockup.IsDefault)
                            {
                            Properties.Settings.Default.MockupUrl = targetPath;
                            Properties.Settings.Default.Save();
                            }

                        Growl.SuccessGlobal("样机添加成功！");
                        }
                    catch (Exception ex)
                        {
                        Growl.ErrorGlobal($"添加样机失败: {ex.Message}");
                        }
                    }
                }
            }

        private void EditMockup_Click(object sender, RoutedEventArgs e)
            {
            var button = sender as System.Windows.Controls.Button;
            string mockupPath = button?.Tag?.ToString();

            if (!string.IsNullOrEmpty(mockupPath))
                {
                try
                    {
                    System.Diagnostics.Process.Start(mockupPath);
                    }
                catch (Exception ex)
                    {
                    Growl.ErrorGlobal($"打开样机文件失败: {ex.Message}");
                    }
                }
            }

        private void DeleteMockup_Click(object sender, RoutedEventArgs e)
            {
            var button = sender as System.Windows.Controls.Button;
            string mockupPath = button?.Tag?.ToString();

            if (!string.IsNullOrEmpty(mockupPath))
                {
                try
                    {
                    if (MessageBox.Show("确定要删除这个样机吗？", "确认删除", 
                        MessageBoxButtons.YesNo, MessageBoxIcon.Question) == System.Windows.Forms.DialogResult.Yes)
                        {
                        File.Delete(mockupPath);
                        LoadMockupList();
                        Growl.SuccessGlobal("样机删除成功！");
                        }
                    }
                catch (Exception ex)
                    {
                    Growl.ErrorGlobal($"删除样机失败: {ex.Message}");
                    }
                }
            }

        private void MockupListView_SelectionChanged(object sender, System.Windows.Controls.SelectionChangedEventArgs e)
            {
            var selectedItem = MockupListView.SelectedItem as MockupItem;
            if (selectedItem != null)
                {
                // 更新默认样机
                foreach (var item in mockupItems)
                    {
                    item.IsDefault = (item == selectedItem);
                    }
                SaveMockupConfig();

                Properties.Settings.Default.MockupUrl = selectedItem.Path;
                Properties.Settings.Default.Save();
                SaveSlidesAsImages(app, false); // 不强制生成预览图
                LoadImg();
                }
            }

        private void EnsureDirectoryExists()
            {
            try
                {
                if (!Directory.Exists(mockupFolderPath))
                    {
                    Directory.CreateDirectory(mockupFolderPath);
                    Growl.InfoGlobal($"已创建样机文件夹: {mockupFolderPath}");
                    }
                }
            catch (Exception ex)
                {
                Growl.ErrorGlobal($"创建文件夹失败: {ex.Message}");
                }
            }

        public static string GetMockup()
            {
            string location = AppDomain.CurrentDomain.BaseDirectory;
            string mockupPath = Properties.Settings.Default.MockupUrl;

            if (!File.Exists(mockupPath))
                {
                using (OpenFileDialog openFileDialog = new OpenFileDialog())
                    {
                    openFileDialog.InitialDirectory = location;
                    openFileDialog.Filter = "样机文件 (*.pptx)|*.pptx|All files (*.*)|*.*";
                    openFileDialog.Title = "选择文件";

                    if (openFileDialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                        {
                        mockupPath = openFileDialog.FileName;
                        Properties.Settings.Default.MockupUrl = mockupPath;
                        Properties.Settings.Default.Save();
                        }
                    }
                }

            return mockupPath;
            }

        public static void CreateMocKupFolder()
            {
            string location = AppDomain.CurrentDomain.BaseDirectory;
            string mocKupPath = Path.Combine(location, "MocKup");

            if (!Directory.Exists(mocKupPath))
                {
                Directory.CreateDirectory(mocKupPath);
                Console.WriteLine($"文件夹 {mocKupPath} 创建成功！");
                }
            else
                {
                Console.WriteLine($"文件夹 {mocKupPath} 已存在。");
                }

            System.Diagnostics.Process.Start("Explorer.exe", location);
            }

        public static void SaveSlidesAsImages(PowerPoint.Application app, bool forceGenerate = false)
            {
            string location = AppDomain.CurrentDomain.BaseDirectory;
            string mockupPath = GetMockup();

            if (!File.Exists(mockupPath))
                {
                Growl.WarningGlobal($"样机文件不存在: {mockupPath}");
                return;
                }

            try
                {
                // 为每个样机创建唯一的预览文件夹
                string mockupName = Path.GetFileNameWithoutExtension(mockupPath);
                string previewFolderPath = Path.Combine(location, "MocKup", "Previews", mockupName);
                
                // 检查是否需要生成预览图
                bool needsGenerate = forceGenerate || !Directory.Exists(previewFolderPath) || 
                                   !Directory.GetFiles(previewFolderPath, "*.png").Any();

                if (!needsGenerate)
                    {
                    // 检查预览图的最后修改时间是否早于样机文件
                    var mockupLastWrite = File.GetLastWriteTime(mockupPath);
                    var previewFiles = Directory.GetFiles(previewFolderPath, "*.png");
                    if (previewFiles.Any())
                        {
                        var previewLastWrite = previewFiles.Max(f => File.GetLastWriteTime(f));
                        needsGenerate = previewLastWrite < mockupLastWrite;
                        }
                    }

                if (needsGenerate)
                    {
                    // 保存当前活动演示文稿的状态
                    Presentation currentPresentation = app.ActivePresentation;
                    if (currentPresentation != null)
                        {
                        currentPresentation.Save();
                        }

                    // 如果预览文件夹已存在，先删除它
                    if (Directory.Exists(previewFolderPath))
                        {
                        Directory.Delete(previewFolderPath, true);
                        }
                    Directory.CreateDirectory(previewFolderPath);

                    Presentation pre = null;
                    try
                        {
                        pre = app.Presentations.Open(mockupPath);
                        int num = pre.Slides.Count;

                        // 导出新的预览图
                        for (int i = 1; i <= num; i++)
                            {
                            string tempImagePath = Path.Combine(previewFolderPath, $"Slide_{i}.png");
                            float slideWidth = pre.SlideMaster.Width;
                            float slideHeight = pre.SlideMaster.Height;
                            pre.Slides[i].Export(tempImagePath, "PNG", (int)slideWidth, (int)slideHeight);
                            }

                        Growl.SuccessGlobal("预览图生成成功！");
                        }
                    finally
                        {
                        if (pre != null)
                            {
                            pre.Close();
                            }
                        }
                    }
                }
            catch (Exception ex)
                {
                Growl.ErrorGlobal($"生成预览图失败: {ex.Message}");
                }
            }

        public void Application_WindowSelectionChange(Selection Sel)
            {
            LoadNum();
            }

        public void LoadNum()
            {
            var app = Globals.ThisAddIn.Application;
            int selectedCount = app.ActiveWindow?.Selection?.SlideRange?.Count ?? 0;
            LabelNum.Content = $"页面数量: {selectedCount}";
            }

        public void DeleMockUp()
            {
            Presentation pre = app.ActivePresentation;
            int count = pre.Slides.Count;
            for (int i = count ; i > 0 ; i--)
                {
                Slide Myslide = pre.Slides[i];
                if (Myslide.Tags["样机"] == "母版样机")
                    {
                    Myslide.Delete();
                    Growl.SuccessGlobal("删除成功！");
                    }
                }
            }

        public void LoadImg()
            {
            try
                {
                string mockupPath = Properties.Settings.Default.MockupUrl;
                if (string.IsNullOrEmpty(mockupPath))
                    {
                    return;
                    }

                string mockupName = Path.GetFileNameWithoutExtension(mockupPath);
                string previewFolderPath = Path.Combine(mockupFolderPath, "Previews", mockupName);

                if (Directory.Exists(previewFolderPath))
                    {
                    var imageFiles = Directory.GetFiles(previewFolderPath, "*.png")
                                            .OrderBy(f => int.Parse(Path.GetFileNameWithoutExtension(f).Split('_')[1]));
                    
                    if (imageFiles.Any())
                        {
                        var imageList = new List<System.Windows.Media.Imaging.BitmapImage>();
                        foreach (var imagePath in imageFiles)
                            {
                            var bitmap = new System.Windows.Media.Imaging.BitmapImage();
                            bitmap.BeginInit();
                            bitmap.CacheOption = System.Windows.Media.Imaging.BitmapCacheOption.OnLoad;
                            bitmap.UriSource = new Uri(imagePath);
                            bitmap.EndInit();
                            bitmap.Freeze(); // 提高性能
                            imageList.Add(bitmap);
                            }
                        
                        // 使用Dispatcher确保在UI线程上更新
                        System.Windows.Application.Current.Dispatcher.Invoke(() =>
                            {
                            CoverFlowMain.Items.Clear();
                            foreach (var image in imageList)
                                {
                                CoverFlowMain.Items.Add(image);
                                }
                            if (CoverFlowMain.Items.Count > 0)
                                {
                                CoverFlowMain.SelectedIndex = 0;
                                CoverFlowMain.ScrollIntoView(CoverFlowMain.SelectedItem);
                                }
                            });
                        }
                    else
                        {
                        System.Windows.Application.Current.Dispatcher.Invoke(() =>
                            {
                            CoverFlowMain.Items.Clear();
                            });
                        Growl.InfoGlobal("当前样机没有预览图");
                        }
                    }
                }
            catch (Exception ex)
                {
                Growl.ErrorGlobal($"加载预览图失败: {ex.Message}");
                }
            }

        private void Button_Click(object sender, RoutedEventArgs e)
            {
            var presentation = app.ActivePresentation;

            if (presentation == null || presentation.Slides.Count == 0)
                {
                Growl.ErrorGlobal("没有可导出的幻灯片！");
                return;
                }

            // 检查是否已选择样机
            if (string.IsNullOrEmpty(Properties.Settings.Default.MockupUrl))
                {
                Growl.ErrorGlobal("请先选择样机！");
                return;
                }

            List<string> imagePaths = new List<string>();

            if (SelAll.IsChecked == true)
                {
                SlideRange slideRange = presentation.Slides.Range();
                string[] paths = ExportSlideImages(slideRange);
                if (paths != null && paths.Length > 0)
                    {
                    imagePaths.AddRange(paths);
                    }
                else
                    {
                    Growl.ErrorGlobal("导出所有幻灯片失败，请重试！");
                    return;
                    }
                }
            else if (SelBtn.IsChecked == true)
                {
                var selectedSlides = app.ActiveWindow.Selection.SlideRange;
                if (selectedSlides == null || selectedSlides.Count == 0)
                    {
                    Growl.ErrorGlobal("请先选择幻灯片页面!");
                    return;
                    }
                string[] paths = ExportSlideImages(selectedSlides);
                imagePaths.AddRange(paths);
                }

            if (imagePaths.Count == 0)
                {
                Growl.ErrorGlobal("导出幻灯片失败，请重试！");
                return;
                }

            // 使用选中的样机索引
            int selectedIndex = CoverFlowMain.SelectedIndex + 1;
            if (selectedIndex <= 0) selectedIndex = 1;

            InsertTemplateAndImages(app, selectedIndex, imagePaths.ToArray());

            if (presentation.Slides.Count > 0)
                {
                Slide lastSlide = presentation.Slides[presentation.Slides.Count];
                app.ActiveWindow.View.GotoSlide(lastSlide.SlideIndex);
                }
            }

        public static string[] ExportSlideImages(SlideRange selectedSlides)
            {
            string tempDir = Path.Combine(Path.GetTempPath(), "selectedSlideImages");
            if (!Directory.Exists(tempDir))
                {
                Directory.CreateDirectory(tempDir);
                }

            foreach (string file in Directory.GetFiles(tempDir))
                {
                File.Delete(file);
                }

            List<string> imagePaths = new List<string>();
            int slideIndex = 1;

            foreach (Slide slide in selectedSlides)
                {
                string tempImagePath = Path.Combine(tempDir, $"slide_{slideIndex}.png");
                float slideWidth = slide.Master.Width;
                float slideHeight = slide.Master.Height;
                slide.Export(tempImagePath, "PNG", (int)slideWidth, (int)slideHeight);
                imagePaths.Add(tempImagePath);
                slideIndex++;
                }

            return imagePaths.ToArray();
            }

        public static void InsertTemplateAndImages(Microsoft.Office.Interop.PowerPoint.Application app, int index, string[] imagePaths)
            {
            try
                {
                Presentation pre = app.ActivePresentation;
                string templatePath = Properties.Settings.Default.MockupUrl;

                if (!File.Exists(templatePath))
                    {
                    Growl.WarningGlobal("默认样机文件不存在，请修复！");
                    return;
                    }

                int num = pre.Slides.Count;

                // 检查是否已经存在样机
                bool foundExisting = false;
                for (int i = num; i > 0; i--)
                    {
                    try
                        {
                        if (pre.Slides[i].Tags["样机"] == "母版样机")
                            {
                            // 如果存在，先删除旧的样机
                            pre.Slides[i].Delete();
                            foundExisting = true;
                            break;
                            }
                        }
                    catch { }
                    }

                // 如果没有找到现有样机，则在指定位置插入
                if (!foundExisting)
                    {
                    pre.Slides.InsertFromFile(templatePath, num, index, index);
                    Slide newSlide = pre.Slides[num + 1];
                    InsertImagesToSlide(newSlide, imagePaths);
                    newSlide.Tags.Add("样机", "母版样机");
                    }
                else
                    {
                    // 如果找到并删除了现有样机，在原位置插入新的
                    pre.Slides.InsertFromFile(templatePath, num - 1, index, index);
                    Slide newSlide = pre.Slides[num];
                    InsertImagesToSlide(newSlide, imagePaths);
                    newSlide.Tags.Add("样机", "母版样机");
                    }

                Growl.SuccessGlobal("样机生成成功！");
                }
            catch (Exception ex)
                {
                Growl.ErrorGlobal($"插入样机失败: {ex.Message}");
                }
            }

        private static void InsertImagesToSlide(Slide slide, string[] imagePaths)
            {
            foreach (var imagePath in imagePaths)
                {
                slide.Shapes.AddPicture(imagePath, MsoTriState.msoFalse, MsoTriState.msoCTrue, 0, 0);
                }
            }

        private void SelAll_Checked(object sender, RoutedEventArgs e)
            {
            Growl.InfoGlobal("注意：全选页面可能导致插入图片过多，建议选择部分需要页面即可！");
            }

        private void Button_Click_1(object sender, RoutedEventArgs e)
            {
            DeleMockUp();
            }

        private void OpenFile_Click(object sender, RoutedEventArgs e)
            {
            string mocKupPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "MocKup");
            System.Diagnostics.Process.Start("Explorer.exe", mocKupPath);
            }

        private void SeltMockup_Click(object sender, RoutedEventArgs e)
            {
            string location = AppDomain.CurrentDomain.BaseDirectory;
            string mockupUrl = Properties.Settings.Default.MockupUrl;

            using (OpenFileDialog openFileDialog = new OpenFileDialog())
                {
                openFileDialog.InitialDirectory = location;
                openFileDialog.Filter = "样机文件 (*.pptx)|*.pptx|All files (*.*)|*.*";
                openFileDialog.Title = "选择演示稿文件";

                if (openFileDialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                    {
                    mockupUrl = openFileDialog.FileName;
                    Properties.Settings.Default.MockupUrl = mockupUrl;
                    Properties.Settings.Default.Save();
                    SaveSlidesAsImages(app, true); // 选择新样机时强制生成预览图
                    LoadImg();
                    }
                }
            }
        }
    }
