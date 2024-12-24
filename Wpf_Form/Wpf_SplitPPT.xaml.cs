using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Threading.Tasks;
using System.Windows;
using Microsoft.Office.Interop.PowerPoint;
using System.Windows.Forms;
using HandyControl.Controls;
using System.Threading;
using System.ComponentModel;
using System.Windows.Controls;
using System.Windows.Media.Imaging;
using System.IO;
using System.Linq;
using System.Windows.Media;
using Microsoft.Win32;
using Microsoft.Office.Core;
using System.Windows.Input;

namespace PresPio.Wpf_Form
{
    public class PPTFile : INotifyPropertyChanged
    {
        private string _name;
        private string _status = "等待处理";
        private double _progress;
        private string _fullPath;

        public string Name
        {
            get => _name;
            set
            {
                _name = value;
                OnPropertyChanged(nameof(Name));
            }
        }

        public string Status
        {
            get => _status;
            set
            {
                _status = value;
                OnPropertyChanged(nameof(Status));
            }
        }

        public double Progress
        {
            get => _progress;
            set
            {
                _progress = value;
                OnPropertyChanged(nameof(Progress));
            }
        }

        public string FullPath
        {
            get => _fullPath;
            set
            {
                _fullPath = value;
                OnPropertyChanged(nameof(FullPath));
            }
        }

        public event PropertyChangedEventHandler PropertyChanged;

        protected virtual void OnPropertyChanged(string propertyName)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }

    public partial class Wpf_SplitPPT : Window, INotifyPropertyChanged
    {
        private CancellationTokenSource _cancellationTokenSource;
        private Presentation _presentation;
        private Application _application;
        private string _withoutExtension;
        private string _fileName;
        private int _currentSlideIndex = 1;
        private ObservableCollection<PPTFile> _pptFiles = new ObservableCollection<PPTFile>();

        public event PropertyChangedEventHandler PropertyChanged;

        public string FileName
        {
            get => _fileName;
            set
            {
                _fileName = value;
                OnPropertyChanged(nameof(FileName));
            }
        }

        public Wpf_SplitPPT(Presentation presentation)
        {
            InitializeComponent();
            DataContext = this;
            
            _presentation = presentation;
            _application = presentation.Application;
            _withoutExtension = System.IO.Path.GetFileNameWithoutExtension(presentation.Name);
            FileName = presentation.Name;
            _cancellationTokenSource = new CancellationTokenSource();

            // 初始化UI
            FileListView.ItemsSource = _pptFiles;
            CustomSplitRadio.Checked += (s, e) => CustomSplitPanel.Visibility = Visibility.Visible;
            SinglePageRadio.Checked += (s, e) => CustomSplitPanel.Visibility = Visibility.Collapsed;
            
            // 加载第一页预览
            LoadSlidePreview(_currentSlideIndex);
            UpdatePageNumber();
        }

        private void OnPropertyChanged(string propertyName)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }

        private async Task LoadSlidePreviewAsync(int slideIndex)
        {
            try
            {
                if (slideIndex < 1 || slideIndex > _presentation.Slides.Count)
                    return;

                var slide = _presentation.Slides[slideIndex];
                string tempPath = Path.Combine(Path.GetTempPath(), $"preview_{Guid.NewGuid()}.png");
                
                // 使用更高的分辨率导出
                slide.Export(tempPath, "PNG", 1920, 1080);

                var bitmap = new BitmapImage();
                bitmap.BeginInit();
                bitmap.CacheOption = BitmapCacheOption.OnLoad;
                bitmap.UriSource = new Uri(tempPath);
                bitmap.EndInit();

                PreviewImage.Source = bitmap;

                // 清理临时文件
                await Task.Run(() =>
                {
                    Thread.Sleep(1000);
                    try { File.Delete(tempPath); } catch { }
                });
            }
            catch (Exception ex)
            {
                Growl.ErrorGlobal($"预览失败：{ex.Message}");
            }
        }

        private void LoadSlidePreview(int slideIndex)
        {
            _ = LoadSlidePreviewAsync(slideIndex);
        }

        private void UpdatePageNumber()
        {
            PageNumberText.Text = $"{_currentSlideIndex}/{_presentation.Slides.Count}";
            PrevButton.IsEnabled = _currentSlideIndex > 1;
            NextButton.IsEnabled = _currentSlideIndex < _presentation.Slides.Count;
        }

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
            if (_currentSlideIndex < _presentation.Slides.Count)
            {
                _currentSlideIndex++;
                LoadSlidePreview(_currentSlideIndex);
                UpdatePageNumber();
            }
        }

        private async void StartButton_Click(object sender, RoutedEventArgs e)
        {
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
                    
                    if (SinglePageRadio.IsChecked == true)
                    {
                        await SplitSinglePage(exportPath);
                    }
                    else
                    {
                        await SplitCustomRange(exportPath);
                    }
                }
            }
            catch (Exception ex)
            {
                Growl.ErrorGlobal($"拆分失败：{ex.Message}");
            }
            finally
            {
                StartButton.IsEnabled = true;
            }
        }

        private async Task SplitSinglePage(string exportPath)
        {
            int slideCount = _presentation.Slides.Count;
            ProgressText.Text = $"0/{slideCount}";
            PercentText.Text = "0%";

            try
            {
                for (int i = 1; i <= slideCount; i++)
                {
                    if (_cancellationTokenSource.Token.IsCancellationRequested)
                        return;

                    var newPresentation = _application.Presentations.Add(MsoTriState.msoFalse);
                    _presentation.Slides[i].Copy();
                    newPresentation.Slides.Paste();
                    
                    string slidePath = Path.Combine(exportPath, $"{_withoutExtension}_第{i}页.pptx");
                    newPresentation.SaveAs(slidePath);
                    newPresentation.Close();

                    double progress = (double)i / slideCount * 100;
                    ProgressBar.Value = progress;
                    ProgressText.Text = $"{i}/{slideCount}";
                    PercentText.Text = $"{Math.Round(progress)}%";

                    await Task.Delay(100); // 避免UI卡顿
                }

                CompleteButton.Visibility = Visibility.Visible;
                StartButton.Visibility = Visibility.Collapsed;
                Growl.SuccessGlobal("拆分完成！");
            }
            catch (Exception ex)
            {
                Growl.ErrorGlobal($"拆分过程中出错：{ex.Message}");
            }
        }

        private async Task SplitCustomRange(string exportPath)
        {
            var ranges = ParseRanges(CustomSplitTextBox.Text);
            if (ranges == null || !ranges.Any())
            {
                Growl.WarningGlobal("请输入有效的页码范围！");
                return;
            }

            int totalRanges = ranges.Count;
            int current = 0;

            try
            {
                foreach (var range in ranges)
                {
                    if (_cancellationTokenSource.Token.IsCancellationRequested)
                        return;

                    var newPresentation = _application.Presentations.Add(MsoTriState.msoFalse);
                    
                    for (int i = range.Item1; i <= range.Item2; i++)
                    {
                        _presentation.Slides[i].Copy();
                        newPresentation.Slides.Paste();
                    }

                    string slidePath = Path.Combine(exportPath, $"{_withoutExtension}_第{range.Item1}-{range.Item2}页.pptx");
                    newPresentation.SaveAs(slidePath);
                    newPresentation.Close();

                    current++;
                    double progress = (double)current / totalRanges * 100;
                    ProgressBar.Value = progress;
                    ProgressText.Text = $"{current}/{totalRanges}";
                    PercentText.Text = $"{Math.Round(progress)}%";

                    await Task.Delay(100);
                }

                CompleteButton.Visibility = Visibility.Visible;
                StartButton.Visibility = Visibility.Collapsed;
                Growl.SuccessGlobal("拆分完成！");
            }
            catch (Exception ex)
            {
                Growl.ErrorGlobal($"拆分过程中出错：{ex.Message}");
            }
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

        private void SelectFilesButton_Click(object sender, RoutedEventArgs e)
        {
            var dialog = new Microsoft.Win32.OpenFileDialog
            {
                Multiselect = true,
                Filter = "PowerPoint文件|*.ppt;*.pptx"
            };

            if (dialog.ShowDialog() == true)
            {
                foreach (string file in dialog.FileNames)
                {
                    if (!_pptFiles.Any(p => p.FullPath == file))
                    {
                        _pptFiles.Add(new PPTFile
                        {
                            Name = Path.GetFileName(file),
                            FullPath = file,
                            Status = "等待处理",
                            Progress = 0
                        });
                    }
                }
            }
        }

        private async void StartBatchButton_Click(object sender, RoutedEventArgs e)
        {
            if (!_pptFiles.Any())
            {
                Growl.WarningGlobal("请先选择要处理的PPT文件！");
                return;
            }

            var dialog = new FolderBrowserDialog
            {
                Description = "请选择拆分���件存放文件夹:"
            };

            if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                StartBatchButton.IsEnabled = false;
                SelectFilesButton.IsEnabled = false;

                try
                {
                    foreach (var pptFile in _pptFiles)
                    {
                        if (_cancellationTokenSource.Token.IsCancellationRequested)
                            break;

                        pptFile.Status = "处理中";
                        var presentation = _application.Presentations.Open(
                            pptFile.FullPath,
                            MsoTriState.msoFalse,
                            MsoTriState.msoFalse,
                            MsoTriState.msoFalse
                        );

                        string exportFolder = Path.Combine(dialog.SelectedPath, Path.GetFileNameWithoutExtension(pptFile.Name));
                        Directory.CreateDirectory(exportFolder);

                        int slideCount = presentation.Slides.Count;
                        for (int i = 1; i <= slideCount; i++)
                        {
                            if (_cancellationTokenSource.Token.IsCancellationRequested)
                                break;

                            var newPresentation = _application.Presentations.Add(MsoTriState.msoFalse);
                            presentation.Slides[i].Copy();
                            newPresentation.Slides.Paste();

                            string slidePath = Path.Combine(exportFolder, $"第{i}页.pptx");
                            newPresentation.SaveAs(slidePath);
                            newPresentation.Close();

                            pptFile.Progress = (double)i / slideCount * 100;
                            await Task.Delay(50);
                        }

                        presentation.Close();
                        pptFile.Status = "已完成";
                        pptFile.Progress = 100;
                    }

                    if (!_cancellationTokenSource.Token.IsCancellationRequested)
                    {
                        Growl.SuccessGlobal("批量拆分完成！");
                    }
                }
                catch (Exception ex)
                {
                    Growl.ErrorGlobal($"批量拆分失败：{ex.Message}");
                }
                finally
                {
                    StartBatchButton.IsEnabled = true;
                    SelectFilesButton.IsEnabled = true;
                }
            }
        }

        private void CancelButton_Click(object sender, RoutedEventArgs e)
        {
            _cancellationTokenSource.Cancel();
            Close();
        }

        private void CompleteButton_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }

        private void PreviewBorder_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            if (_currentSlideIndex < 1 || _currentSlideIndex > _presentation.Slides.Count)
                return;

            try
            {
                var previewWindow = new Window
                {
                    Title = $"预览 - 第{_currentSlideIndex}页",
                    Width = 1024,
                    Height = 768,
                    WindowStartupLocation = WindowStartupLocation.CenterScreen,
                    Background = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#F0F2F5"))
                };

                var border = new Border
                {
                    Margin = new Thickness(20),
                    Background = Brushes.White,
                    CornerRadius = new CornerRadius(8),
                    Effect = TryFindResource("ShadowEffect") as DropShadowEffect
                };

                var image = new Image
                {
                    Source = PreviewImage.Source,
                    Stretch = Stretch.Uniform,
                    Margin = new Thickness(20)
                };
                RenderOptions.SetBitmapScalingMode(image, BitmapScalingMode.HighQuality);

                border.Child = image;
                previewWindow.Content = border;
                previewWindow.ShowDialog();
            }
            catch (Exception ex)
            {
                Growl.ErrorGlobal($"预览失败：{ex.Message}");
            }
        }
    }
} 