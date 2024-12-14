using System;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Forms;
using HandyControl.Controls;
using Microsoft.Office.Interop.PowerPoint;
using PowerPoint= Microsoft.Office.Interop.PowerPoint;

namespace PresPio
    {
    /// <summary>
    /// Wpf_ImageExport.xaml 的交互逻辑
    /// </summary>
    public partial class Wpf_ImageExport
        {
        public PowerPoint.Application app { get; set; }

        public Wpf_ImageExport()
            {
            app = Globals.ThisAddIn.Application;
            InitializeComponent();
            LoadDpi();
            LoadFilePath();
            SizeTextBox.Text = GetSize();
            }

        private void RadioButton_Checked(object sender, RoutedEventArgs e)
            {
            if (sender is System.Windows.Controls.RadioButton radioButton && radioButton.IsChecked == true)
                {
                Growl.Success("选中的按钮内容为: " + radioButton.Name);
                }
            }

        public void LoadFilePath()
            {
            string pptFilePath = app.ActivePresentation.FullName;
            string pptFileName = Path.GetFileNameWithoutExtension(pptFilePath);
            string exportFolderPath = Path.Combine(Path.GetDirectoryName(pptFilePath), $"{pptFileName}_HD");

            Properties.Settings.Default.exportFolderPath = exportFolderPath;
            Properties.Settings.Default.Save();
            filePath.Text = exportFolderPath;
            }

        public void LoadDpi()
            {
            DpiComBox.Items.Clear();
            int defaultDpi = GetDPI();
            double[] dpiValues = { defaultDpi, 75, 100, 150, 300, 350, 400, 450, 500, 800, 1000 };

            foreach (var dpi in dpiValues)
                {
                DpiComBox.Items.Add(dpi);
                }

            if (DpiComBox.Items.Count > 0)
                {
                DpiComBox.SelectedItem = DpiComBox.Items[0];
                }
            }

        public int GetDPI()
            {
            string regPath = @"HKEY_CURRENT_USER\Software\Microsoft\Office";
            double version = double.Parse(app.Version, CultureInfo.InvariantCulture);
            int[] versions = { 16, 15, 14 };
            string[] versionPaths = { @"\16.0\PowerPoint\Options", @"\15.0\PowerPoint\Options", @"\14.0\PowerPoint\Options" };

            for (int i = 0 ; i < versions.Length ; i++)
                {
                if (version == versions[i])
                    {
                    regPath += versionPaths[i];
                    break;
                    }
                }

            object dpiValue = Microsoft.Win32.Registry.GetValue(regPath, "ExportBitmapResolution", null);
            return dpiValue is int dpi ? dpi : 0;
            }


        public void OutImg(string format)
            {
            Selection sel = app.ActiveWindow.Selection;
            if (sel.Type != PpSelectionType.ppSelectionSlides)
                {
                Growl.Warning("请选择幻灯片页面后再试");
                return;
                }

            string pptFilePath = app.ActivePresentation.FullName;
            string pptFileName = Path.GetFileNameWithoutExtension(pptFilePath);
            string exportFolderPath = Properties.Settings.Default.exportFolderPath;

            // 确保导出文件夹存在
            EnsureExportFolderExists(exportFolderPath);

            float dpi = GetDPI();
            string imageFormat = GetImageFormat(format);

            // 使用并行处理导出每个幻灯片
            ExportSlides(sel, pptFileName, exportFolderPath, dpi, imageFormat);

            // 打开导出文件夹
            Process.Start(exportFolderPath);
            }

        private void EnsureExportFolderExists(string exportFolderPath)
            {
            if (!Directory.Exists(exportFolderPath))
                {
                Directory.CreateDirectory(exportFolderPath);
                }
            }

        private string GetImageFormat(string format)
            {
            if (format == "JPEG")
                return "JPG";
            else if (format == "PNG")
                return "PNG";
            else if (format == "GIF")
                return "GIF";
            else if (format == "TIFF")
                return "TIF";
            else if (format == "BMP")
                return "BMP";
            else
                return "PNG"; // 默认格式
            }

        private void ExportSlides(Selection sel, string pptFileName, string exportFolderPath, float dpi, string imageFormat)
            {
            Parallel.ForEach(sel.SlideRange.Cast<Slide>(), slide =>
            {
                int width = (int)(slide.Master.Width / 72 * dpi);
                int height = (int)(slide.Master.Height / 72 * dpi);
                slide.Export(Path.Combine(exportFolderPath, $"{pptFileName}_Page{slide.SlideNumber}.{imageFormat.ToLower()}"), imageFormat, width, height);
            });
            }

        private void Button_Click(object sender, RoutedEventArgs e)
            {
            if (RadioButtonJpg.IsChecked == true)
                {
                OutImg("JPEG");
                }
            else if (RadioButtonPng.IsChecked == true)
                {
                OutImg("PNG");
                }
            else if (RadioButtonGif.IsChecked == true)
                {
                OutImg("GIF");
                }
            else if (RadioButtonTif.IsChecked == true)
                {
                OutImg("TIFF");
                }
            else if (RadioButtonBmp.IsChecked == true)
                {
                OutImg("BMP");
                }
            }

        private void DpiComBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
            {
            if (DpiComBox.SelectedItem is double selectedDpi)
                {
                Properties.Settings.Default.PPtDPI = (int)selectedDpi;
                Properties.Settings.Default.Save();
                DpiSetting();
                SizeTextBox.Text = GetSize();
                Growl.Success($"DPI设置成功为{selectedDpi}");
                }
            }

        public void DpiSetting()
            {
            int dpi = Properties.Settings.Default.PPtDPI;
            string regPath = @"HKEY_CURRENT_USER\Software\Microsoft\Office";
            double version = double.Parse(app.Version, CultureInfo.InvariantCulture);

            // 根据 PowerPoint 版本构建注册表路径
            if (version == 16)
                {
                regPath += @"\16.0\PowerPoint\Options";
                }
            else if (version == 15)
                {
                regPath += @"\15.0\PowerPoint\Options";
                }
            else if (version == 14)
                {
                regPath += @"\14.0\PowerPoint\Options";
                }
            else
                {
                throw new InvalidOperationException($"无法处理 PowerPoint {version}。");
                }

            // 设置注册表值
            Microsoft.Win32.Registry.SetValue(regPath, "ExportBitmapResolution", dpi, Microsoft.Win32.RegistryValueKind.DWord);
            }


        public string GetSize()
            {
            Slide slide = app.ActiveWindow.View.Slide;
            float dpi = GetDPI();
            int width = (int)(slide.Master.Width / 72 * dpi);
            int height = (int)(slide.Master.Height / 72 * dpi);
            return $"{width}*{height}";
            }

        private void FolderSelect_Click(object sender, RoutedEventArgs e)
            {
            string exportFolder = Properties.Settings.Default.exportFolderPath;
            using (var folderBrowserDialog = new FolderBrowserDialog())
                {
                folderBrowserDialog.Description = "请选择导出文件夹";
                folderBrowserDialog.ShowNewFolderButton = true;
                folderBrowserDialog.SelectedPath = exportFolder;

                if (folderBrowserDialog.ShowDialog() == System.Windows.Forms.DialogResult.OK && !string.IsNullOrWhiteSpace(folderBrowserDialog.SelectedPath))
                    {
                    exportFolder = folderBrowserDialog.SelectedPath;
                    Properties.Settings.Default.exportFolderPath = exportFolder;
                    filePath.Text = exportFolder;
                    }
                }
            }
        }
    }
