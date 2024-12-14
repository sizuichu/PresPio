using HandyControl.Controls;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using System;
using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.Windows;
using System.Windows.Forms;
using System.Windows.Threading;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using System.Threading.Tasks;
using System.Threading;
using MessageBox = HandyControl.Controls.MessageBox;
using Window = System.Windows.Window;
using static Emgu.CV.VideoCapture;

namespace PresPio
{
    public partial class Wpf_Publish : Window
    {
        private PowerPoint.Application _app;
        private CancellationTokenSource _cts;
        private readonly IProgress<int> _progress;
        private DateTime startTime;

        public Wpf_Publish()
        {
            InitializeComponent();
            _app = Globals.ThisAddIn.Application;
            _progress = new Progress<int>(value => 
            {
                ExportProgress.Value = value;
                ExportButton.Content = $"导出中 {value}%";
            });

            InitializeEvents();
            LoadDefaultSettings();
        }

        private void InitializeEvents()
        {
            BrowseButton.Click += OnBrowseClick;
            ExportButton.Click += OnExportClick;
            Loaded += OnWindowLoaded;
        }

        private void LoadDefaultSettings()
        {
            ExportPathBox.Text = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            LoadCheckBoxSettings();
        }

        private void LoadCheckBoxSettings()
        {
            PptxCheck.IsChecked = Properties.Settings.Default.Publish_Chek1;
            PpsxCheck.IsChecked = Properties.Settings.Default.Publish_Chek2;
            ThemeCheck.IsChecked = Properties.Settings.Default.Publish_Chek3;
            FontCheck.IsChecked = Properties.Settings.Default.Publish_Chek4;
            PdfCheck.IsChecked = Properties.Settings.Default.Publish_Chek5;
            ImageCheck.IsChecked = Properties.Settings.Default.Publish_Chek7;
            VideoCheck.IsChecked = Properties.Settings.Default.Publish_Chek8;
        }

        private void OnWindowLoaded(object sender, RoutedEventArgs e)
        {
            if (_app?.ActivePresentation != null)
            {
                SourceFileText.Text = _app.ActivePresentation.Name;
            }
        }

        private void OnBrowseClick(object sender, RoutedEventArgs e)
        {
            var dialog = new FolderBrowserDialog
            {
                Description = "选择导出位置",
                ShowNewFolderButton = true
            };

            if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                ExportPathBox.Text = dialog.SelectedPath;
            }
        }

        private async void OnExportClick(object sender, RoutedEventArgs e)
        {
            if (!ValidateExport())
                return;

            try
            {
                ExportButton.IsEnabled = false;
                startTime = DateTime.Now;
                _cts = new CancellationTokenSource();
                
                await ExportFilesAsync(_cts.Token);
                
                MessageBox.Show("导出完成！", "成功", MessageBoxButton.OK, MessageBoxImage.Information);
                
                if (MessageBox.Show("是否打开导出文件夹？", "提示", 
                    MessageBoxButton.YesNo, MessageBoxImage.Question) == MessageBoxResult.Yes)
                {
                    Process.Start("explorer.exe", ExportPathBox.Text);
                }
            }
            catch (OperationCanceledException)
            {
                MessageBox.Show("导出已取消", "提示");
            }
            catch (Exception ex)
            {
                MessageBox.Show($"导出过程中出现错误：{ex.Message}", "错误", 
                    MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                ExportButton.IsEnabled = true;
                ExportButton.Content = "开始导出";
                _cts?.Dispose();
            }
        }

        private bool ValidateExport()
        {
            if (string.IsNullOrEmpty(ExportPathBox.Text))
            {
                MessageBox.Show("请选择导出位置", "提示");
                return false;
            }

            if (!Directory.Exists(ExportPathBox.Text))
            {
                try
                {
                    Directory.CreateDirectory(ExportPathBox.Text);
                }
                catch
                {
                    MessageBox.Show("无法创建导出目录", "错误");
                    return false;
                }
            }

            return true;
        }

        private async Task ExportFilesAsync(CancellationToken token)
        {
            var presentation = _app.ActivePresentation;
            var basePath = ExportPathBox.Text;
            var fileName = Path.GetFileNameWithoutExtension(presentation.Name);
            int progress = 0;
            int totalTasks = GetTotalTasks();

            // 创建导出目录结构
            var exportPath = Path.Combine(basePath, $"{fileName}_导出文件");
            var presentationPath = Path.Combine(exportPath, "01-演示文件");
            var previewPath = Path.Combine(exportPath, "02-预览文件");
            var themePath = Path.Combine(exportPath, "03-主题文件");
            var printPath = Path.Combine(exportPath, "04-印刷文件");
            var videoPath = Path.Combine(exportPath, "05-演示视频");

            // 创建所需的目录
            Directory.CreateDirectory(exportPath);
            Directory.CreateDirectory(presentationPath);
            Directory.CreateDirectory(previewPath);
            Directory.CreateDirectory(themePath);
            Directory.CreateDirectory(printPath);
            Directory.CreateDirectory(videoPath);

            // 保存导出路径到设置
            Properties.Settings.Default.Publish_URL = exportPath;
            Properties.Settings.Default.Save();

            // 导出演示文件
            if (PptxCheck.IsChecked == true)
            {
                var pptxPath = Path.Combine(presentationPath, fileName);
                await Task.Run(() =>
                {
                    _app.Presentations[1].SaveCopyAs(
                        pptxPath,
                        PpSaveAsFileType.ppSaveAsPresentation, MsoTriState.msoFalse
                        );
                    UpdateProgress(ref progress, totalTasks);
                }, token);
            }

            // 导出放映文件
            if (PpsxCheck.IsChecked == true)
            {
                var ppsxPath = Path.Combine(presentationPath, fileName);
                await Task.Run(() =>
                {
                    presentation.SaveCopyAs(
                        ppsxPath,
                        PpSaveAsFileType.ppSaveAsShow,
                        FontCheck.IsChecked == true ? MsoTriState.msoTrue : MsoTriState.msoFalse
                    );
                    UpdateProgress(ref progress, totalTasks);
                }, token);
            }

            // 导出主题文件
            if (ThemeCheck.IsChecked == true)
            {
                var themxPath = Path.Combine(themePath, fileName);
                await Task.Run(() =>
                {
                    presentation.SaveCopyAs(
                        themxPath,
                        PpSaveAsFileType.ppSaveAsOpenXMLTheme,
                        MsoTriState.msoFalse
                    );
                    UpdateProgress(ref progress, totalTasks);
                }, token);
            }

            // 导出PDF文件
            if (PdfCheck.IsChecked == true)
            {
                var pdfPath = Path.Combine(printPath, fileName);
                await Task.Run(() =>
                {
                    presentation.SaveAs(
                        pdfPath,
                        PpSaveAsFileType.ppSaveAsPDF
                    );
                    UpdateProgress(ref progress, totalTasks);
                }, token);
            }

            // 导出图片
            if (ImageCheck.IsChecked == true)
            {
                await Task.Run(() =>
                {
                    int slideCount = presentation.Slides.Count;
                    for (int i = 1; i <= slideCount; i++)
                    {
                        float height = presentation.SlideMaster.Height * 2;
                        float width = presentation.SlideMaster.Width * 2;
                        string imagePath = Path.Combine(previewPath, $"{fileName}_{i}.png");
                        presentation.Slides[i].Export(imagePath, "PNG", (int)width, (int)height);
                    }
                    UpdateProgress(ref progress, totalTasks);
                }, token);
            }

            // 导出视频
            if (VideoCheck.IsChecked == true)
            {
                var videoFilePath = Path.Combine(videoPath, fileName);
                await Task.Run(() =>
                {
                    presentation.CreateVideo(videoPath, false, 5, 1080, 60, 100);
                    UpdateProgress(ref progress, totalTasks);
                }, token);
            }
        }

        private int GetTotalTasks()
        {
            int count = 0;
            if (PptxCheck.IsChecked == true) count++;
            if (PpsxCheck.IsChecked == true) count++;
            if (ThemeCheck.IsChecked == true) count++;
            if (PdfCheck.IsChecked == true) count++;
            if (ImageCheck.IsChecked == true) count++;
            if (VideoCheck.IsChecked == true) count++;
            return Math.Max(1, count);
        }

        private void UpdateProgress(ref int progress, int total)
        {
            progress++;
            var percentage = (int)((double)progress / total * 100);
            _progress.Report(percentage);
        }

        protected override void OnClosing(CancelEventArgs e)
        {
            _cts?.Cancel();
            base.OnClosing(e);
        }

        #region CheckBox Event Handlers
        private void PptxCheck_Click(object sender, RoutedEventArgs e)
        {
            Properties.Settings.Default.Publish_Chek1 = PptxCheck.IsChecked ?? false;
            Properties.Settings.Default.Save();
        }

        private void PpsxCheck_Click(object sender, RoutedEventArgs e)
        {
            Properties.Settings.Default.Publish_Chek2 = PpsxCheck.IsChecked ?? false;
            Properties.Settings.Default.Save();
        }

        private void ThemeCheck_Click(object sender, RoutedEventArgs e)
        {
            Properties.Settings.Default.Publish_Chek3 = ThemeCheck.IsChecked ?? false;
            Properties.Settings.Default.Save();
        }

        private void FontCheck_Click(object sender, RoutedEventArgs e)
        {
            Properties.Settings.Default.Publish_Chek4 = FontCheck.IsChecked ?? false;
            Properties.Settings.Default.Save();
        }

        private void PdfCheck_Click(object sender, RoutedEventArgs e)
        {
            Properties.Settings.Default.Publish_Chek5 = PdfCheck.IsChecked ?? false;
            Properties.Settings.Default.Save();
        }

        private void ImageCheck_Click(object sender, RoutedEventArgs e)
        {
            Properties.Settings.Default.Publish_Chek7 = ImageCheck.IsChecked ?? false;
            Properties.Settings.Default.Save();
        }

        private void VideoCheck_Click(object sender, RoutedEventArgs e)
        {
            Properties.Settings.Default.Publish_Chek8 = VideoCheck.IsChecked ?? false;
            Properties.Settings.Default.Save();
        }
        #endregion
    }
}
