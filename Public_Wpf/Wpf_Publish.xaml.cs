using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Forms;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using ComboBox = System.Windows.Controls.ComboBox;
using ComboBoxItem = System.Windows.Controls.ComboBoxItem;
using MessageBox = HandyControl.Controls.MessageBox;

using PowerPoint = Microsoft.Office.Interop.PowerPoint;

using Window = System.Windows.Window;

namespace PresPio
    {
    public partial class Wpf_Publish : Window
        {
        private PowerPoint.Application _app;
        private CancellationTokenSource _cts;
        private readonly IProgress<int> _progress;
        private DateTime startTime;
        private ComboBox _imageFormatComboBox;
        private ComboBox _imageQualityComboBox;

        public Wpf_Publish()
            {
            InitializeComponent();
            _app = Globals.ThisAddIn.Application;
            _progress = new Progress<int>(value =>
            {
                System.Windows.Application.Current.Dispatcher.Invoke(() =>
                {
                    ExportProgress.Value = value;
                    ExportButton.Content = $"导出中 {value}%";
                });
            });

            // 初始化 ComboBox 引用
            _imageFormatComboBox = ImageFormatComboBox;
            _imageQualityComboBox = ImageQualityComboBox;

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
            var exportPath = ExportPathBox.Text;
            var fileName = Path.GetFileNameWithoutExtension(presentation.Name);
            int progress = 0;
            int totalTasks = GetTotalTasks();

            // 在进入后台线程前获取所有 UI 控件的状态
            var isPptxChecked = PptxCheck.IsChecked == true;
            var isPpsxChecked = PpsxCheck.IsChecked == true;
            var isThemeChecked = ThemeCheck.IsChecked == true;
            var isFontChecked = FontCheck.IsChecked == true;
            var isPdfChecked = PdfCheck.IsChecked == true;
            var isImageChecked = ImageCheck.IsChecked == true;
            var isVideoChecked = VideoCheck.IsChecked == true;

            // 创建导出目录结构
            Directory.CreateDirectory(exportPath);

            // 动态创建文件夹路径
            var folderPaths = new Dictionary<string, string>();
            int folderIndex = 1;

            // 演示文件（包含 pptx 和 ppsx）
            if (isPptxChecked || isPpsxChecked)
                {
                folderPaths["presentation"] = Path.Combine(exportPath, $"{folderIndex:D2}-演示文件");
                folderIndex++;
                }

            // 预览文件（图片）
            if (isImageChecked)
                {
                folderPaths["preview"] = Path.Combine(exportPath, $"{folderIndex:D2}-预览文件");
                folderIndex++;
                }

            // 主题文件
            if (isThemeChecked)
                {
                folderPaths["theme"] = Path.Combine(exportPath, $"{folderIndex:D2}-主题文件");
                folderIndex++;
                }

            // 印刷文件（PDF）
            if (isPdfChecked)
                {
                folderPaths["print"] = Path.Combine(exportPath, $"{folderIndex:D2}-印刷文件");
                folderIndex++;
                }

            // 演示视频
            if (isVideoChecked)
                {
                folderPaths["video"] = Path.Combine(exportPath, $"{folderIndex:D2}-演示视频");
                }

            // 创建所需的目录
            foreach (var path in folderPaths.Values)
                {
                Directory.CreateDirectory(path);
                }

            // 保存导出路径到设置
            Properties.Settings.Default.Publish_URL = exportPath;
            Properties.Settings.Default.Save();

            // 导出演示文件
            if (isPptxChecked)
                {
                var pptxPath = Path.Combine(folderPaths["presentation"], $"{fileName}_编辑版.pptx");
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
            if (isPpsxChecked)
                {
                var ppsxPath = Path.Combine(folderPaths["presentation"], $"{fileName}_放映版.ppsx");
                await Task.Run(() =>
                {
                    presentation.SaveCopyAs(
                        ppsxPath,
                        PpSaveAsFileType.ppSaveAsShow,
                        isFontChecked ? MsoTriState.msoTrue : MsoTriState.msoFalse
                    );
                    UpdateProgress(ref progress, totalTasks);
                }, token);
                }

            // 导出主题文件
            if (isThemeChecked)
                {
                var themxPath = Path.Combine(folderPaths["theme"], $"{fileName}.thmx");
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
            if (isPdfChecked)
                {
                var pdfPath = Path.Combine(folderPaths["print"], $"{fileName}.pdf");
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
            if (isImageChecked)
                {
                // 获取图片格式和质量设置
                string imageFormat = "PNG";

                System.Windows.Application.Current.Dispatcher.Invoke(() =>
                {
                    if (_imageFormatComboBox != null)
                        {
                        var selectedFormat = (_imageFormatComboBox.SelectedItem as ComboBoxItem)?.Content.ToString();
                        imageFormat = selectedFormat ?? "PNG";
                        }

                    if (_imageQualityComboBox != null)
                        {
                        var selectedQuality = (_imageQualityComboBox.SelectedItem as ComboBoxItem)?.Content.ToString();
                        switch (selectedQuality)
                            {
                            case "高":
                                break;

                            case "中":
                                break;

                            case "低":
                                break;

                            default:
                                break;
                            }
                        }
                });

                await Task.Run(() =>
                {
                    int slideCount = presentation.Slides.Count;
                    string numberFormat = new string('0', slideCount.ToString().Length);

                    for (int i = 1 ; i <= slideCount ; i++)
                        {
                        float height = presentation.PageSetup.SlideHeight * 2;
                        float width = presentation.PageSetup.SlideWidth * 2;

                        // 设置正确的文件扩展名和导出格式
                        string extension;
                        string exportFormat;
                        if (imageFormat.Equals("JPEG", StringComparison.OrdinalIgnoreCase))
                            {
                            extension = "jpg";
                            exportFormat = "JPG";
                            }
                        else
                            {
                            extension = "png";
                            exportFormat = "PNG";
                            }

                        string imagePath = Path.Combine(folderPaths["preview"], $"{fileName}_幻灯片{i.ToString(numberFormat)}.{extension}");
                        presentation.Slides[i].Export(imagePath, exportFormat, (int)width, (int)height);
                        }
                    UpdateProgress(ref progress, totalTasks);
                }, token);
                }

            // 导出视频
            if (isVideoChecked)
                {
                var videoFilePath = Path.Combine(folderPaths["video"], $"{fileName}.mp4");
                await Task.Run(() =>
                {
                    presentation.CreateVideo(videoFilePath, false, 5, 1080, 60, 100);
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
            System.Windows.Application.Current.Dispatcher.Invoke(() =>
            {
                _progress.Report(percentage);
            });
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

        #endregion CheckBox Event Handlers
        }
    }