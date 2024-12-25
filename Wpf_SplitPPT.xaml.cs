using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using System.Windows;
using Microsoft.Office.Interop.PowerPoint;
using System.Windows.Forms;
using System.Threading;
using System.ComponentModel;
using System.Windows.Controls;
using System.Windows.Media.Imaging;
using System.IO;
using System.Linq;
using Microsoft.Office.Core;
using Application = Microsoft.Office.Interop.PowerPoint.Application;
using Shape = Microsoft.Office.Interop.PowerPoint.Shape;
using MessageBox = System.Windows.MessageBox;

namespace PresPio.Wpf_Form
{
    /// <summary>
    /// PPT拆分工具窗体
    /// </summary>
    public partial class Wpf_SplitPPT : Window
    {
        #region 字段

        // PowerPoint相关
        private Application _application;
        private Presentation _presentation;
        private int _currentSlideIndex = 1;
        private CancellationTokenSource _cancellationTokenSource;

        #endregion

        #region 初始化

        public Wpf_SplitPPT()
        {
            InitializeComponent();
            InitializeData();
        }

        private void InitializeData()
        {
            _application = Globals.ThisAddIn.Application;
            _cancellationTokenSource = new CancellationTokenSource();
        }

        protected override void OnClosing(CancelEventArgs e)
        {
            base.OnClosing(e);
            CleanupResources();
        }

        private void CleanupResources()
        {
            _cancellationTokenSource?.Cancel();
            
            if (_presentation != null)
            {
                try
                {
                    _presentation.Close();
                }
                catch { }
            }
        }

        #endregion

        #region 文件加载和预览

        private void LoadPresentation(string filePath)
        {
            try
            {
                if (_presentation != null)
                {
                    _presentation.Close();
                }

                _presentation = _application.Presentations.Open(
                    filePath,
                    MsoTriState.msoFalse,
                    MsoTriState.msoFalse,
                    MsoTriState.msoFalse
                );

                _currentSlideIndex = 1;
                LoadSlidePreview(_currentSlideIndex);
                UpdatePageNumber();
            }
            catch (Exception ex)
            {
                ShowError($"打开文件失败：{ex.Message}");
            }
        }

        private async Task LoadSlidePreviewAsync(int slideIndex)
        {
            try
            {
                if (_presentation == null || slideIndex < 1 || slideIndex > _presentation.Slides.Count)
                    return;

                var slide = _presentation.Slides[slideIndex];
                string tempPath = Path.Combine(Path.GetTempPath(), $"preview_{Guid.NewGuid()}.png");
                
                // 使用PPT默认的16:9比例导出
                slide.Export(tempPath, "PNG", 1920, 1080);

                await Task.Run(() =>
                {
                    var bitmap = new BitmapImage();
                    bitmap.BeginInit();
                    bitmap.CacheOption = BitmapCacheOption.OnLoad;
                    bitmap.UriSource = new Uri(tempPath);
                    bitmap.EndInit();
                    bitmap.Freeze();

                    Dispatcher.Invoke(() =>
                    {
                        PreviewImage.Source = bitmap;
                    });

                    Thread.Sleep(1000);
                    try { File.Delete(tempPath); } catch { }
                });
            }
            catch (Exception ex)
            {
                ShowError($"预览失败：{ex.Message}");
            }
        }

        private void LoadSlidePreview(int slideIndex)
        {
            _ = LoadSlidePreviewAsync(slideIndex);
        }

        private void UpdatePageNumber()
        {
            if (_presentation != null && _presentation.Slides.Count > 0)
            {
                PageNumberText.Text = $"{_currentSlideIndex}/{_presentation.Slides.Count}";
                PrevButton.IsEnabled = _currentSlideIndex > 1;
                NextButton.IsEnabled = _currentSlideIndex < _presentation.Slides.Count;
            }
            else
            {
                PageNumberText.Text = "0/0";
                PrevButton.IsEnabled = false;
                NextButton.IsEnabled = false;
            }
        }

        #endregion

        #region 导航事件处理

        private void PrevButton_Click(object sender, RoutedEventArgs e)
        {
            if (_currentSlideIndex > 1)
            {
                _currentSlideIndex--;
                LoadSlidePreview(_currentSlideIndex);
                UpdatePageNumber();
            }
        }

        private void NextButton_Click(object sender, RoutedEventArgs e)
        {
            if (_presentation != null && _currentSlideIndex < _presentation.Slides.Count)
            {
                _currentSlideIndex++;
                LoadSlidePreview(_currentSlideIndex);
                UpdatePageNumber();
            }
        }

        #endregion

        #region 拆分功能

        private void AddFilesButton_Click(object sender, RoutedEventArgs e)
        {
            var dialog = new Microsoft.Win32.OpenFileDialog
            {
                Filter = "PowerPoint文件|*.ppt;*.pptx",
                Multiselect = true
            };

            if (dialog.ShowDialog() == true)
            {
                foreach (string file in dialog.FileNames)
                {
                    if (!FileListView.Items.Cast<string>().Contains(file))
                    {
                        FileListView.Items.Add(file);
                    }
                }

                // 加载第一个文件进行预览
                if (FileListView.Items.Count > 0 && _presentation == null)
                {
                    LoadPresentation(FileListView.Items[0].ToString());
                }
            }
        }

        private void ClearButton_Click(object sender, RoutedEventArgs e)
        {
            FileListView.Items.Clear();
            if (_presentation != null)
            {
                _presentation.Close();
                _presentation = null;
            }
            PreviewImage.Source = null;
            UpdatePageNumber();
        }

        private async void StartButton_Click(object sender, RoutedEventArgs e)
        {
            if (FileListView.Items.Count == 0)
            {
                ShowWarning("请先选择要拆分的PPT文件！");
                return;
            }

            try
            {
                var dialog = new FolderBrowserDialog
                {
                    Description = "请选择拆分文件存放文件夹:"
                };

                if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    StartButton.IsEnabled = false;
                    CancelButton.IsEnabled = true;
                    string exportPath = dialog.SelectedPath;

                    await ProcessFiles(exportPath);

                    CompleteButton.Visibility = Visibility.Visible;
                    StartButton.IsEnabled = true;
                    CancelButton.IsEnabled = false;
                    ShowSuccess("拆分完成！");
                }
            }
            catch (Exception ex)
            {
                ShowError($"拆分失败：{ex.Message}");
                StartButton.IsEnabled = true;
                CancelButton.IsEnabled = false;
            }
        }

        private async Task ProcessFiles(string exportPath)
        {
            int totalSlides = 0;
            int processedSlides = 0;

            // 首先计算总幻灯片数
            foreach (string filePath in FileListView.Items)
            {
                var presentation = _application.Presentations.Open(
                    filePath,
                    MsoTriState.msoFalse,
                    MsoTriState.msoFalse,
                    MsoTriState.msoFalse
                );
                
                if (SinglePageRadio.IsChecked == true)
                {
                    totalSlides++;
                }
                else if (CustomRangeRadio.IsChecked == true)
                {
                    var ranges = ParseRanges(CustomRangeTextBox.Text);
                    if (ranges != null)
                    {
                        totalSlides += ranges.Sum(r => r.Item2 - r.Item1 + 1);
                    }
                }
                else
                {
                    totalSlides += presentation.Slides.Count;
                }
                
                presentation.Close();
            }

            foreach (string filePath in FileListView.Items)
            {
                if (_cancellationTokenSource.Token.IsCancellationRequested)
                    return;

                try
                {
                    var presentation = _application.Presentations.Open(
                        filePath,
                        MsoTriState.msoFalse,
                        MsoTriState.msoFalse,
                        MsoTriState.msoFalse
                    );

                    string fileName = Path.GetFileNameWithoutExtension(filePath);
                    string fileExportPath = Path.Combine(exportPath, fileName);
                    Directory.CreateDirectory(fileExportPath);

                    if (SinglePageRadio.IsChecked == true)
                    {
                        // 导出当前页
                        string slidePath = Path.Combine(fileExportPath, $"{fileName}_第{_currentSlideIndex}页.pptx");
                        presentation.Slides[_currentSlideIndex].Export(slidePath, "pptx", 0, 0);
                        processedSlides++;
                        UpdateProgress(processedSlides, totalSlides);
                    }
                    else if (CustomRangeRadio.IsChecked == true)
                    {
                        // 导出自定义范围
                        var ranges = ParseRanges(CustomRangeTextBox.Text);
                        if (ranges != null)
                        {
                            var tasks = new List<Task>();
                            foreach (var range in ranges)
                            {
                                int start = range.Item1;
                                int end = range.Item2;
                                tasks.Add(Task.Run(() =>
                                {
                                    for (int i = start; i <= end; i++)
                                    {
                                        if (_cancellationTokenSource.Token.IsCancellationRequested)
                                            return;

                                        string slidePath = Path.Combine(fileExportPath, $"{fileName}_第{i}页.pptx");
                                        presentation.Slides[i].Export(slidePath, "pptx", 0, 0);
                                        Interlocked.Increment(ref processedSlides);
                                        Dispatcher.Invoke(() => UpdateProgress(processedSlides, totalSlides));
                                    }
                                }));
                            }
                            await Task.WhenAll(tasks);
                        }
                    }
                    else
                    {
                        // 导出所有页面
                        int slideCount = presentation.Slides.Count;
                        int threadsCount = Environment.ProcessorCount;
                        int slidesPerThread = (slideCount + threadsCount - 1) / threadsCount;
                        var tasks = new List<Task>();

                        for (int i = 0; i < threadsCount; i++)
                        {
                            int startIndex = i * slidesPerThread + 1;
                            int endIndex = Math.Min(startIndex + slidesPerThread - 1, slideCount);

                            tasks.Add(Task.Run(() =>
                            {
                                for (int j = startIndex; j <= endIndex; j++)
                                {
                                    if (_cancellationTokenSource.Token.IsCancellationRequested)
                                        return;

                                    string slidePath = Path.Combine(fileExportPath, $"{fileName}_第{j}页.pptx");
                                    presentation.Slides[j].Export(slidePath, "pptx", 0, 0);
                                    Interlocked.Increment(ref processedSlides);
                                    Dispatcher.Invoke(() => UpdateProgress(processedSlides, totalSlides));
                                }
                            }));
                        }

                        await Task.WhenAll(tasks);
                    }

                    presentation.Close();
                }
                catch (Exception ex)
                {
                    ShowError($"处理文件 {Path.GetFileName(filePath)} 时出错: {ex.Message}");
                }
            }
        }

        private void UpdateProgress(int current, int total)
        {
            double progress = (double)current / total * 100;
            ProgressBar.Value = progress;
            ProgressText.Text = $"{current}/{total}";
        }

        private List<Tuple<int, int>> ParseRanges(string input)
        {
            try
            {
                var ranges = new List<Tuple<int, int>>();
                var parts = input.Split(new[] { ',', '，', ' ' }, StringSplitOptions.RemoveEmptyEntries);

                foreach (var part in parts)
                {
                    var range = part.Trim().Split(new[] { '-', '~' });
                    if (range.Length == 1)
                    {
                        int page = int.Parse(range[0]);
                        if (page > 0 && page <= _presentation.Slides.Count)
                        {
                            ranges.Add(new Tuple<int, int>(page, page));
                        }
                    }
                    else if (range.Length == 2)
                    {
                        int start = int.Parse(range[0]);
                        int end = int.Parse(range[1]);
                        if (start <= end && start > 0 && end <= _presentation.Slides.Count)
                        {
                            ranges.Add(new Tuple<int, int>(start, end));
                        }
                    }
                }

                return ranges.Any() ? ranges : null;
            }
            catch
            {
                return null;
            }
        }

        #endregion

        #region 通用事件处理

        private void CancelButton_Click(object sender, RoutedEventArgs e)
        {
            _cancellationTokenSource?.Cancel();
            _cancellationTokenSource = new CancellationTokenSource();
            CancelButton.IsEnabled = false;
            StartButton.IsEnabled = true;
            ShowInfo("已取消拆分操作");
        }

        private void CompleteButton_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }

        #endregion

        #region 消息提示

        private void ShowError(string message)
        {
            MessageBox.Show(message, "错误", MessageBoxButton.OK, MessageBoxImage.Error);
        }

        private void ShowWarning(string message)
        {
            MessageBox.Show(message, "警告", MessageBoxButton.OK, MessageBoxImage.Warning);
        }

        private void ShowSuccess(string message)
        {
            MessageBox.Show(message, "成功", MessageBoxButton.OK, MessageBoxImage.Information);
        }

        private void ShowInfo(string message)
        {
            MessageBox.Show(message, "提示", MessageBoxButton.OK, MessageBoxImage.Information);
        }

        #endregion

        private void FileListView_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (FileListView.SelectedItem != null)
            {
                LoadPresentation(FileListView.SelectedItem.ToString());
            }
        }
    }
}
