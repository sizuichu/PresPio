using HandyControl.Controls;
using HandyControl.Data;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Documents;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using Color = System.Windows.Media.Color;
using DocumentTextRange = System.Windows.Documents.TextRange;
using MediaFontFamily = System.Windows.Media.FontFamily;

using PowerPoint = Microsoft.Office.Interop.PowerPoint;

using Window = HandyControl.Controls.Window;

namespace PresPio
    {
    public class SlideItem : INotifyPropertyChanged
        {
        public int Index { get; set; }
        public string Title { get; set; }
        public BitmapImage Thumbnail { get; set; }
        private bool isSelected;

        public bool IsSelected
            {
            get => isSelected;
            set
                {
                if (isSelected != value)
                    {
                    isSelected = value;
                    OnPropertyChanged(nameof(IsSelected));
                    }
                }
            }

        public event PropertyChangedEventHandler PropertyChanged;

        protected virtual void OnPropertyChanged(string propertyName)
            {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
            }
        }

    public partial class Wpf_Manuscript : Window
        {
        private PowerPoint.Application pptApplication;
        private PowerPoint.Presentation currentPresentation;
        private int currentSlideIndex = 0;
        private int totalSlides = 0;
        private bool isModified = false;
        private bool isExternalPPT = false;
        private ObservableCollection<SlideItem> slides = new ObservableCollection<SlideItem>();

        public Wpf_Manuscript()
            {
            InitializeComponent();
            InitializeControls();
            }

        private void InitializeControls()
            {
            BtnPrev.IsEnabled = false;
            BtnNext.IsEnabled = false;
            BtnPreview.IsEnabled = false;
            BtnPrint.IsEnabled = false;
            BtnExportSelected.IsEnabled = false;
            BtnExportAll.IsEnabled = false;

            SlideList.ItemsSource = slides;
            FontSizeCombo.SelectedIndex = 0;
            }

        private void LoadSlideList()
            {
            slides.Clear();
            for (int i = 1 ; i <= totalSlides ; i++)
                {
                var slide = currentPresentation.Slides[i];
                string tempPath = Path.Combine(Path.GetTempPath(), $"thumb_{i}.png");
                slide.Export(tempPath, "PNG", 80, 60);

                var bitmap = new BitmapImage();
                bitmap.BeginInit();
                bitmap.CacheOption = BitmapCacheOption.OnLoad;
                bitmap.UriSource = new Uri(tempPath);
                bitmap.EndInit();

                slides.Add(new SlideItem
                    {
                    Index = i,
                    Title = $"第 {i} 页",
                    Thumbnail = bitmap,
                    IsSelected = false
                    });

                try { File.Delete(tempPath); } catch { }
                }
            UpdateSelectedCount();
            }

        private void BtnUseActive_Click(object sender, RoutedEventArgs e)
            {
            try
                {
                CleanupPPT();
                pptApplication = Marshal.GetActiveObject("PowerPoint.Application") as PowerPoint.Application;
                if (pptApplication != null && pptApplication.Presentations.Count > 0)
                    {
                    currentPresentation = pptApplication.ActivePresentation;
                    totalSlides = currentPresentation.Slides.Count;
                    currentSlideIndex = pptApplication.ActiveWindow.Selection.SlideRange.SlideNumber;
                    isExternalPPT = true;

                    LoadSlideList();
                    UpdateUI();
                    LoadCurrentSlide();

                    EnableControls(true);
                    StatusText.Text = "已连接到活动PPT";
                    }
                else
                    {
                    Growl.Warning("未找到打开的PPT文件");
                    }
                }
            catch (Exception ex)
                {
                Growl.Error($"连接PPT时出错：{ex.Message}");
                CleanupPPT();
                }
            }

        private void EnableControls(bool enabled)
            {
            BtnPrev.IsEnabled = enabled;
            BtnNext.IsEnabled = enabled;
            BtnPreview.IsEnabled = enabled;
            BtnPrint.IsEnabled = enabled;
            BtnExportSelected.IsEnabled = enabled;
            BtnExportAll.IsEnabled = enabled;
            }

        private void BtnImport_Click(object sender, RoutedEventArgs e)
            {
            Microsoft.Win32.OpenFileDialog dlg = new Microsoft.Win32.OpenFileDialog
                {
                DefaultExt = ".pptx",
                Filter = "PowerPoint 文件|*.pptx;*.ppt"
                };

            if (dlg.ShowDialog() == true)
                {
                try
                    {
                    CleanupPPT();
                    pptApplication = new PowerPoint.Application();
                    currentPresentation = pptApplication.Presentations.Open(dlg.FileName, MsoTriState.msoTrue, MsoTriState.msoFalse, MsoTriState.msoFalse);
                    totalSlides = currentPresentation.Slides.Count;
                    currentSlideIndex = 1;
                    isExternalPPT = false;

                    LoadSlideList();
                    UpdateUI();
                    LoadCurrentSlide();

                    EnableControls(true);
                    StatusText.Text = "已加载PPT文";
                    }
                catch (Exception ex)
                    {
                    Growl.Error($"加载PPT时出错：{ex.Message}");
                    CleanupPPT();
                    }
                }
            }

        private void LoadCurrentSlide()
            {
            if (currentPresentation == null || currentSlideIndex < 1 || currentSlideIndex > totalSlides)
                {
                NoPreviewText.Visibility = Visibility.Visible;
                PreviewImage.Source = null;
                return;
                }

            try
                {
                NoPreviewText.Visibility = Visibility.Collapsed;
                PowerPoint.Slide slide = currentPresentation.Slides[currentSlideIndex];

                float slideWidth = currentPresentation.PageSetup.SlideWidth;
                float slideHeight = currentPresentation.PageSetup.SlideHeight;
                float aspectRatio = slideWidth / slideHeight;

                int exportWidth = 1024;
                int exportHeight = (int)(exportWidth / aspectRatio);

                string tempPath = Path.Combine(Path.GetTempPath(), $"slide_{currentSlideIndex}.png");
                slide.Export(tempPath, "PNG", exportWidth, exportHeight);

                BitmapImage bitmap = new BitmapImage();
                bitmap.BeginInit();
                bitmap.CacheOption = BitmapCacheOption.OnLoad;
                bitmap.UriSource = new Uri(tempPath);
                bitmap.EndInit();
                PreviewImage.Source = bitmap;

                NotesTextBox.Document.Blocks.Clear();
                try
                    {
                    // 确保备注页存在
                    var notesPage = slide.NotesPage;
                    if (notesPage != null)
                        {
                        // 获取备注形状
                        var notesShape = notesPage.Shapes.Placeholders[2];
                        if (notesShape != null && notesShape.HasTextFrame == MsoTriState.msoTrue)
                            {
                            var textRange = notesShape.TextFrame.TextRange;

                            // 获取格式化的文本
                            string notes = textRange.Text.Trim();
                            if (!string.IsNullOrEmpty(notes))
                                {
                                var paragraph = new Paragraph();
                                var run = new Run(notes);

                                // 应用PPT中的格式
                                if (textRange.Font.Bold == MsoTriState.msoTrue)
                                    run.FontWeight = FontWeights.Bold;
                                if (textRange.Font.Italic == MsoTriState.msoTrue)
                                    run.FontStyle = FontStyles.Italic;
                                if (textRange.Font.Underline == MsoTriState.msoTrue)
                                    run.TextDecorations = TextDecorations.Underline;

                                // 应用字体大小
                                if (textRange.Font.Size > 0)
                                    run.FontSize = textRange.Font.Size;

                                // 应用字体颜色
                                int rgb = textRange.Font.Color.RGB;
                                if (rgb != 0)
                                    {
                                    byte r = (byte)((rgb >> 16) & 0xFF);
                                    byte g = (byte)((rgb >> 8) & 0xFF);
                                    byte b = (byte)(rgb & 0xFF);
                                    run.Foreground = new SolidColorBrush(Color.FromRgb(r, g, b));
                                    }

                                paragraph.Inlines.Add(run);
                                NotesTextBox.Document.Blocks.Add(paragraph);

                                // 应用段落格式
                                var range = new DocumentTextRange(NotesTextBox.Document.ContentStart, NotesTextBox.Document.ContentEnd);
                                range.ApplyPropertyValue(TextElement.FontFamilyProperty, new MediaFontFamily("微软雅黑"));
                                range.ApplyPropertyValue(Paragraph.LineHeightProperty, 1.5);
                                }
                            }
                        }
                    }
                catch (Exception ex)
                    {
                    Growl.Warning($"读取备注时出错：{ex.Message}");
                    }

                try { File.Delete(tempPath); } catch { }

                UpdateUI();
                UpdateWordCount();
                SlideList.SelectedIndex = currentSlideIndex - 1;

                if (SlideList.SelectedItem != null)
                    {
                    SlideList.ScrollIntoView(SlideList.SelectedItem);
                    }
                }
            catch (Exception ex)
                {
                NoPreviewText.Visibility = Visibility.Visible;
                PreviewImage.Source = null;
                Growl.Error($"加载幻灯片时出错：{ex.Message}");
                }
            }

        private void SaveCurrentNotes()
            {
            if (currentPresentation == null || currentSlideIndex < 1 || currentSlideIndex > totalSlides)
                return;

            try
                {
                PowerPoint.Slide slide = currentPresentation.Slides[currentSlideIndex];

                // 确保备注页存在
                var notesPage = slide.NotesPage;
                if (notesPage != null)
                    {
                    // 获取备注形状
                    var notesShape = notesPage.Shapes.Placeholders[2];
                    if (notesShape != null && notesShape.HasTextFrame == MsoTriState.msoTrue)
                        {
                        var textRange = notesShape.TextFrame.TextRange;
                        textRange.Text = string.Empty; // 清空现有内容

                        // 遍历RichTextBox中的内容，保持格式
                        var document = NotesTextBox.Document;
                        foreach (var block in document.Blocks)
                            {
                            if (block is Paragraph paragraph)
                                {
                                string paragraphText = string.Empty;
                                foreach (var inline in paragraph.Inlines)
                                    {
                                    if (inline is Run run)
                                        {
                                        // 获取当前文本范围
                                        var currentRange = textRange.InsertAfter(run.Text);

                                        // 应用字体样式
                                        if (run.FontWeight == FontWeights.Bold)
                                            currentRange.Font.Bold = MsoTriState.msoTrue;
                                        if (run.FontStyle == FontStyles.Italic)
                                            currentRange.Font.Italic = MsoTriState.msoTrue;
                                        if (run.TextDecorations == TextDecorations.Underline)
                                            currentRange.Font.Underline = MsoTriState.msoTrue;

                                        // 应用字体大小
                                        if (run.FontSize > 0)
                                            currentRange.Font.Size = (float)run.FontSize;

                                        // 应用字体颜色
                                        if (run.Foreground is SolidColorBrush brush)
                                            {
                                            currentRange.Font.Color.RGB =
                                                (brush.Color.R << 16) | (brush.Color.G << 8) | brush.Color.B;
                                            }
                                        }
                                    }
                                textRange.InsertAfter("\n");
                                }
                            }

                        // 设置段落对齐方式
                        textRange.ParagraphFormat.Alignment = PpParagraphAlignment.ppAlignLeft;

                        // 设置默认字体
                        textRange.Font.Name = "微软雅黑";
                        if (textRange.Font.Size < 12)
                            textRange.Font.Size = 14;
                        }
                    }
                isModified = true;
                }
            catch (Exception ex)
                {
                Growl.Error($"保存备注时出错：{ex.Message}");
                }
            }

        private void ExportToPDF(string filePath, List<int> slideIndices = null)
            {
            SaveCurrentNotes(); // 导出前保存当前备注

            // 创建临时演示文稿
            PowerPoint.Presentation tempPres = null;
            try
                {
                // 创建新的演示文稿
                tempPres = pptApplication.Presentations.Add(MsoTriState.msoFalse);

                // 设置页面大小为A4
                tempPres.PageSetup.SlideSize = PowerPoint.PpSlideSizeType.ppSlideSizeA4Paper;

                // 设置页面方向和布局
                bool isPortrait = PortraitRadio.IsChecked == true;
                bool isDoublePage = LandscapeDoubleRadio.IsChecked == true;

                if (!isPortrait)
                    {
                    tempPres.PageSetup.SlideWidth = 960;  // A4横向宽度
                    tempPres.PageSetup.SlideHeight = 720; // A4横向高度
                    }
                else
                    {
                    tempPres.PageSetup.SlideWidth = 720;  // A4纵向宽度
                    tempPres.PageSetup.SlideHeight = 960; // A4纵向高度
                    }

                // 获取要导出的幻灯片索引
                var indices = slideIndices ?? Enumerable.Range(1, currentPresentation.Slides.Count).ToList();

                // 为每个幻灯片创建新的页面
                for (int i = 0 ; i < indices.Count ; i += (isDoublePage ? 2 : 1))
                    {
                    // 添加新的空白幻灯片
                    PowerPoint.Slide newSlide = tempPres.Slides.Add(tempPres.Slides.Count + 1, PowerPoint.PpSlideLayout.ppLayoutBlank);

                    // 计算布局参数
                    float slideWidth = tempPres.PageSetup.SlideWidth;
                    float slideHeight = tempPres.PageSetup.SlideHeight;
                    float margin = 20;
                    float imageTop = margin;

                    if (isDoublePage)
                        {
                        // 横向双张布局
                        float contentWidth = (slideWidth - margin * 3) / 2; // 每个内容区域的宽度
                        float imageHeight = contentWidth * 0.75f;  // 保持4:3比例
                        float textTop = imageTop + imageHeight + margin;
                        float textHeight = slideHeight - textTop - margin;

                        // 添加第一张幻灯片
                        AddSlideContent(currentPresentation.Slides[indices[i]], newSlide,
                            margin, imageTop, contentWidth, imageHeight,
                            margin, textTop, contentWidth, textHeight);

                        // 添加第二张幻灯片（如果存在）
                        if (i + 1 < indices.Count)
                            {
                            float secondLeft = margin * 2 + contentWidth;
                            AddSlideContent(currentPresentation.Slides[indices[i + 1]], newSlide,
                                secondLeft, imageTop, contentWidth, imageHeight,
                                secondLeft, textTop, contentWidth, textHeight);
                            }
                        }
                    else
                        {
                        // 横向单张或纵向布局
                        float contentWidth = slideWidth - margin * 2;
                        float imageHeight;
                        if (isPortrait)
                            {
                            imageHeight = contentWidth * 0.75f; // 保持4:3比例
                            if (imageHeight > slideHeight * 0.4f)
                                {
                                imageHeight = slideHeight * 0.4f;
                                contentWidth = imageHeight / 0.75f;
                                }
                            }
                        else
                            {
                            imageHeight = contentWidth * 0.75f; // 保持4:3比例
                            if (imageHeight > slideHeight * 0.6f)
                                {
                                imageHeight = slideHeight * 0.6f;
                                contentWidth = imageHeight / 0.75f;
                                }
                            }

                        float imageLeft = (slideWidth - contentWidth) / 2;
                        float textTop = imageTop + imageHeight + margin;
                        float textHeight = slideHeight - textTop - margin;

                        AddSlideContent(currentPresentation.Slides[indices[i]], newSlide,
                            imageLeft, imageTop, contentWidth, imageHeight,
                            imageLeft, textTop, contentWidth, textHeight);
                        }
                    }

                // 导出为PDF
                tempPres.SaveAs(filePath, PowerPoint.PpSaveAsFileType.ppSaveAsPDF, MsoTriState.msoTrue);

                System.Diagnostics.Process.Start(filePath);
                Growl.Success("导出完成");
                }
            catch (Exception ex)
                {
                Growl.Error($"导出时出错：{ex.Message}");
                }
            finally
                {
                if (tempPres != null)
                    {
                    try
                        {
                        tempPres.Close();
                        Marshal.ReleaseComObject(tempPres);
                        }
                    catch { }
                    }
                }
            }

        private void AddSlideContent(PowerPoint.Slide originalSlide, PowerPoint.Slide newSlide,
            float imageLeft, float imageTop, float imageWidth, float imageHeight,
            float textLeft, float textTop, float textWidth, float textHeight)
            {
            // 导出原始幻灯片为图片，保持原始比例
            string tempImagePath = Path.Combine(Path.GetTempPath(), $"slide_{originalSlide.SlideNumber}.png");

            // 获取原始幻灯片的尺寸
            float originalWidth = originalSlide.Master.Width;
            float originalHeight = originalSlide.Master.Height;
            float originalRatio = originalWidth / originalHeight;

            // 根据原始比例调整图片尺寸
            float adjustedWidth = imageWidth;
            float adjustedHeight = imageWidth / originalRatio;

            if (adjustedHeight > imageHeight)
                {
                adjustedHeight = imageHeight;
                adjustedWidth = imageHeight * originalRatio;
                }

            // 计算居中位置
            float adjustedLeft = imageLeft + (imageWidth - adjustedWidth) / 2;
            float adjustedTop = imageTop + (imageHeight - adjustedHeight) / 2;

            // 导出图片，使用较大的尺寸以保持质量
            originalSlide.Export(tempImagePath, "PNG", (int)(adjustedWidth * 2), (int)(adjustedHeight * 2));

            // 添加图片
            PowerPoint.Shape imageShape = newSlide.Shapes.AddPicture(
                tempImagePath, MsoTriState.msoFalse, MsoTriState.msoTrue,
                adjustedLeft, adjustedTop, adjustedWidth, adjustedHeight);

            // 删除临时图片文件
            try { File.Delete(tempImagePath); } catch { }

            // 获取并添加备注
            var notesPage = originalSlide.NotesPage;
            if (notesPage != null)
                {
                var notesShape = notesPage.Shapes.Placeholders[2];
                if (notesShape != null && notesShape.HasTextFrame == MsoTriState.msoTrue)
                    {
                    var originalTextRange = notesShape.TextFrame.TextRange;
                    string notes = originalTextRange.Text.Trim();

                    if (!string.IsNullOrEmpty(notes))
                        {
                        // 添加备注文本框
                        PowerPoint.Shape textShape = newSlide.Shapes.AddTextbox(
                            MsoTextOrientation.msoTextOrientationHorizontal,
                            textLeft, textTop, textWidth, textHeight);

                        // 复制文本和格式
                        var textRange = textShape.TextFrame.TextRange;
                        textRange.Text = originalTextRange.Text;

                        // 复制段落格式
                        textRange.ParagraphFormat.Alignment = originalTextRange.ParagraphFormat.Alignment;
                        textRange.ParagraphFormat.SpaceAfter = originalTextRange.ParagraphFormat.SpaceAfter;
                        textRange.ParagraphFormat.SpaceBefore = originalTextRange.ParagraphFormat.SpaceBefore;
                        textRange.ParagraphFormat.LineRuleWithin = MsoTriState.msoTrue;
                        textRange.ParagraphFormat.SpaceWithin = 1.5f;  // 1.5倍行距

                        // 复制字体格式
                        textRange.Font.Name = originalTextRange.Font.Name;
                        textRange.Font.Size = originalTextRange.Font.Size;
                        textRange.Font.Bold = originalTextRange.Font.Bold;
                        textRange.Font.Italic = originalTextRange.Font.Italic;
                        textRange.Font.Underline = originalTextRange.Font.Underline;
                        if (originalTextRange.Font.Color.RGB != 0)
                            {
                            textRange.Font.Color.RGB = originalTextRange.Font.Color.RGB;
                            }

                        // 设置文本框属性
                        textShape.TextFrame.WordWrap = MsoTriState.msoTrue;
                        textShape.TextFrame.AutoSize = PowerPoint.PpAutoSize.ppAutoSizeNone;
                        }
                    }
                }
            }

        private void BtnExportSelected_Click(object sender, RoutedEventArgs e)
            {
            if (currentPresentation != null && SlideList.SelectedItems.Count > 0)
                {
                var saveDialog = new Microsoft.Win32.SaveFileDialog
                    {
                    DefaultExt = ".pdf",
                    Filter = "PDF文件|*.pdf",
                    FileName = $"{Path.GetFileNameWithoutExtension(currentPresentation.Name)}_备注页面"
                    };

                if (saveDialog.ShowDialog() == true)
                    {
                    var selectedIndices = new List<int>();
                    foreach (SlideItem item in SlideList.SelectedItems)
                        {
                        selectedIndices.Add(item.Index);
                        }
                    ExportToPDF(saveDialog.FileName, selectedIndices);
                    }
                }
            }

        private void BtnExportAll_Click(object sender, RoutedEventArgs e)
            {
            if (currentPresentation != null)
                {
                var saveDialog = new Microsoft.Win32.SaveFileDialog
                    {
                    DefaultExt = ".pdf",
                    Filter = "PDF文件|*.pdf",
                    FileName = $"{Path.GetFileNameWithoutExtension(currentPresentation.Name)}_备注页面"
                    };

                if (saveDialog.ShowDialog() == true)
                    {
                    ExportToPDF(saveDialog.FileName);
                    }
                }
            }

        private void BtnPreview_Click(object sender, RoutedEventArgs e)
            {
            if (currentPresentation != null)
                {
                string tempFile = Path.Combine(Path.GetTempPath(), "preview.pdf");
                ExportToPDF(tempFile);
                }
            }

        private void PrintDocument(List<int> slideIndices = null)
            {
            if (currentPresentation != null)
                {
                try
                    {
                    SaveCurrentNotes(); // 打印前保存当前备注

                    // 设置打印选项
                    currentPresentation.PrintOptions.OutputType = PowerPoint.PpPrintOutputType.ppPrintOutputNotesPages; // 幻灯片和备注
                    currentPresentation.PrintOptions.FitToPage = MsoTriState.msoTrue;
                    currentPresentation.PrintOptions.PrintHiddenSlides = MsoTriState.msoFalse;
                    currentPresentation.PrintOptions.HighQuality = MsoTriState.msoTrue;

                    // 显示打印对话框并打印
                    currentPresentation.PrintOut();
                    }
                catch (Exception ex)
                    {
                    Growl.Error($"打印时出错：{ex.Message}");
                    }
                }
            }

        private void BtnPrint_Click(object sender, RoutedEventArgs e)
            {
            if (currentPresentation != null)
                {
                try
                    {
                    SaveCurrentNotes(); // 打印前保存当前备注

                    // 设置打印选项
                    currentPresentation.PrintOptions.OutputType = PowerPoint.PpPrintOutputType.ppPrintOutputNotesPages; // 幻灯片和备注
                    currentPresentation.PrintOptions.FitToPage = MsoTriState.msoTrue;
                    currentPresentation.PrintOptions.PrintHiddenSlides = MsoTriState.msoFalse;
                    currentPresentation.PrintOptions.HighQuality = MsoTriState.msoTrue;

                    // 显示打印对话框并打印
                    currentPresentation.PrintOut();
                    }
                catch (Exception ex)
                    {
                    Growl.Error($"打印时出错：{ex.Message}");
                    }
                }
            }

        private void SlideList_SelectionChanged(object sender, SelectionChangedEventArgs e)
            {
            if (SlideList.SelectedItem is SlideItem item)
                {
                SaveCurrentNotes(); // 切换页面时保存当前备注
                currentSlideIndex = item.Index;
                LoadCurrentSlide();
                }
            }

        private void BtnBold_Click(object sender, RoutedEventArgs e)
            {
            var selection = NotesTextBox.Selection;
            if (selection.IsEmpty)
                return;

            var textRange = new DocumentTextRange(selection.Start, selection.End);
            var fontWeight = textRange.GetPropertyValue(TextElement.FontWeightProperty);
            textRange.ApplyPropertyValue(TextElement.FontWeightProperty,
                fontWeight.Equals(FontWeights.Bold) ? FontWeights.Normal : FontWeights.Bold);
            }

        private void BtnItalic_Click(object sender, RoutedEventArgs e)
            {
            var selection = NotesTextBox.Selection;
            if (selection.IsEmpty)
                return;

            var textRange = new DocumentTextRange(selection.Start, selection.End);
            var fontStyle = textRange.GetPropertyValue(TextElement.FontStyleProperty);
            textRange.ApplyPropertyValue(TextElement.FontStyleProperty,
                fontStyle.Equals(FontStyles.Italic) ? FontStyles.Normal : FontStyles.Italic);
            }

        private void BtnUnderline_Click(object sender, RoutedEventArgs e)
            {
            var selection = NotesTextBox.Selection;
            if (selection.IsEmpty)
                return;

            var textRange = new DocumentTextRange(selection.Start, selection.End);
            var textDecorations = textRange.GetPropertyValue(Inline.TextDecorationsProperty);
            textRange.ApplyPropertyValue(Inline.TextDecorationsProperty,
                textDecorations.Equals(TextDecorations.Underline) ? null : TextDecorations.Underline);
            }

        private void FontSizeCombo_SelectionChanged(object sender, SelectionChangedEventArgs e)
            {
            if (NotesTextBox == null || FontSizeCombo.SelectedItem == null)
                return;

            var selection = NotesTextBox.Selection;
            if (selection.IsEmpty)
                return;

            var fontSize = double.Parse((FontSizeCombo.SelectedItem as ComboBoxItem).Tag.ToString());
            var textRange = new DocumentTextRange(selection.Start, selection.End);
            textRange.ApplyPropertyValue(TextElement.FontSizeProperty, fontSize);
            }

        private void BtnPrev_Click(object sender, RoutedEventArgs e)
            {
            if (currentSlideIndex > 1)
                {
                SaveCurrentNotes(); // 切换页面时保存当前备注
                currentSlideIndex--;
                LoadCurrentSlide();
                }
            }

        private void BtnNext_Click(object sender, RoutedEventArgs e)
            {
            if (currentSlideIndex < totalSlides)
                {
                SaveCurrentNotes(); // 切换页面时保存当前备注
                currentSlideIndex++;
                LoadCurrentSlide();
                }
            }

        private void UpdateUI()
            {
            PageInfo.Text = $"第 {currentSlideIndex}/{totalSlides} 页";
            BtnPrev.IsEnabled = currentSlideIndex > 1;
            BtnNext.IsEnabled = currentSlideIndex < totalSlides;
            }

        private void NotesTextBox_TextChanged(object sender, TextChangedEventArgs e)
            {
            if (currentPresentation != null)
                {
                UpdateWordCount();
                }
            }

        private void UpdateWordCount()
            {
            var text = new DocumentTextRange(NotesTextBox.Document.ContentStart, NotesTextBox.Document.ContentEnd).Text;
            int count = text.Replace("\r", "").Replace("\n", "").Length;
            WordCount.Text = $"{count} 字";
            }

        private void UpdateAutoSaveStatus()
            {
            AutoSaveStatus.Text = "已自动保存";
            AutoSaveStatus.Foreground = new SolidColorBrush(Colors.Green);

            // 2秒后隐藏保存状态
            var timer = new System.Windows.Threading.DispatcherTimer();
            timer.Interval = TimeSpan.FromSeconds(2);
            timer.Tick += (s, e) =>
            {
                AutoSaveStatus.Text = "";
                timer.Stop();
            };
            timer.Start();
            }

        private void CleanupPPT()
            {
            if (currentPresentation != null)
                {
                if (isModified && !isExternalPPT)
                    {
                    try
                        {
                        currentPresentation.Save();
                        }
                    catch (Exception ex)
                        {
                        Growl.Error($"保存PPT时出错：{ex.Message}");
                        }
                    }
                try
                    {
                    if (!isExternalPPT)
                        {
                        currentPresentation.Close();
                        }
                    Marshal.ReleaseComObject(currentPresentation);
                    }
                catch { }
                currentPresentation = null;
                }

            if (pptApplication != null && !isExternalPPT)
                {
                try
                    {
                    pptApplication.Quit();
                    Marshal.ReleaseComObject(pptApplication);
                    }
                catch { }
                pptApplication = null;
                }

            slides.Clear();
            GC.Collect();
            GC.WaitForPendingFinalizers();
            }

        protected override void OnClosing(System.ComponentModel.CancelEventArgs e)
            {
            CleanupPPT();
            base.OnClosing(e);
            }

        private void BtnClearFormat_Click(object sender, RoutedEventArgs e)
            {
            var selection = NotesTextBox.Selection;
            if (selection.IsEmpty)
                return;

            var textRange = new DocumentTextRange(selection.Start, selection.End);
            textRange.ClearAllProperties();
            textRange.ApplyPropertyValue(TextElement.FontFamilyProperty, NotesTextBox.FontFamily);
            textRange.ApplyPropertyValue(TextElement.FontSizeProperty, NotesTextBox.FontSize);
            textRange.ApplyPropertyValue(TextElement.ForegroundProperty, NotesTextBox.Foreground);
            }

        private void BtnAlignLeft_Click(object sender, RoutedEventArgs e)
            {
            var selection = NotesTextBox.Selection;
            if (selection.IsEmpty)
                return;

            var textRange = new DocumentTextRange(selection.Start, selection.End);
            textRange.ApplyPropertyValue(Paragraph.TextAlignmentProperty, TextAlignment.Left);
            }

        private void BtnAlignCenter_Click(object sender, RoutedEventArgs e)
            {
            var selection = NotesTextBox.Selection;
            if (selection.IsEmpty)
                return;

            var textRange = new DocumentTextRange(selection.Start, selection.End);
            textRange.ApplyPropertyValue(Paragraph.TextAlignmentProperty, TextAlignment.Center);
            }

        private void BtnAlignRight_Click(object sender, RoutedEventArgs e)
            {
            var selection = NotesTextBox.Selection;
            if (selection.IsEmpty)
                return;

            var textRange = new DocumentTextRange(selection.Start, selection.End);
            textRange.ApplyPropertyValue(Paragraph.TextAlignmentProperty, TextAlignment.Right);
            }

        private void BtnAutoFormat_Click(object sender, RoutedEventArgs e)
            {
            if (currentPresentation == null)
                return;

            try
                {
                var textRange = new DocumentTextRange(NotesTextBox.Document.ContentStart, NotesTextBox.Document.ContentEnd);

                // 设置默认格式
                textRange.ApplyPropertyValue(TextElement.FontFamilyProperty, new MediaFontFamily("微软雅黑"));
                textRange.ApplyPropertyValue(TextElement.FontSizeProperty, 14.0);
                textRange.ApplyPropertyValue(Paragraph.TextAlignmentProperty, TextAlignment.Left);
                textRange.ApplyPropertyValue(Paragraph.LineHeightProperty, 1.5);
                textRange.ApplyPropertyValue(Paragraph.MarginProperty, new Thickness(0, 5, 0, 5));

                // 自动分段
                string text = textRange.Text;
                NotesTextBox.Document.Blocks.Clear();
                var paragraphs = text.Split(new[] { "\r\n", "\r", "\n" }, StringSplitOptions.None);
                foreach (var para in paragraphs)
                    {
                    if (string.IsNullOrWhiteSpace(para))
                        continue;

                    var paragraph = new Paragraph(new Run(para.Trim()));
                    paragraph.TextAlignment = TextAlignment.Left;
                    paragraph.LineHeight = 1.5;
                    paragraph.Margin = new Thickness(0, 5, 0, 5);
                    NotesTextBox.Document.Blocks.Add(paragraph);
                    }

                SaveCurrentNotes();
                Growl.Success("自动排版完成");
                }
            catch (Exception ex)
                {
                Growl.Error($"自动排版时出错：{ex.Message}");
                }
            }

        private void BtnClearNotes_Click(object sender, RoutedEventArgs e)
            {
            if (currentPresentation == null)
                return;

            if (HandyControl.Controls.MessageBox.Show("确定要清空当前页面的备注吗？", "确认",
                MessageBoxButton.YesNo, MessageBoxImage.Question) == MessageBoxResult.Yes)
                {
                NotesTextBox.Document.Blocks.Clear();
                SaveCurrentNotes();
                Growl.Success("备注已清空");
                }
            }

        private void SlideList_MouseDoubleClick(object sender, System.Windows.Input.MouseButtonEventArgs e)
            {
            if (SlideList.SelectedItem is SlideItem item)
                {
                SaveCurrentNotes();
                currentSlideIndex = item.Index;
                LoadCurrentSlide();
                Growl.Info($"已跳转到 {currentSlideIndex} 页");
                }
            }

        private void SearchBox_TextChanged(object sender, TextChangedEventArgs e)
            {
            string searchText = SearchBox.Text?.Trim().ToLower() ?? "";
            var listBox = SlideList;
            if (listBox == null) return;

            if (string.IsNullOrEmpty(searchText))
                {
                // 显示所有页面
                foreach (SlideItem item in slides)
                    {
                    var container = listBox.ItemContainerGenerator.ContainerFromItem(item) as ListBoxItem;
                    if (container != null)
                        {
                        container.Visibility = Visibility.Visible;
                        }
                    }
                }
            else
                {
                // 只显示匹配的页面
                foreach (SlideItem item in slides)
                    {
                    var container = listBox.ItemContainerGenerator.ContainerFromItem(item) as ListBoxItem;
                    if (container != null)
                        {
                        container.Visibility = item.Title.ToLower().Contains(searchText)
                            ? Visibility.Visible
                            : Visibility.Collapsed;
                        }
                    }
                }
            }

        private void BtnSelectAll_Click(object sender, RoutedEventArgs e)
            {
            foreach (SlideItem item in slides)
                {
                item.IsSelected = true;
                }
            UpdateSelectedCount();
            }

        private void BtnInvertSelect_Click(object sender, RoutedEventArgs e)
            {
            foreach (SlideItem item in slides)
                {
                item.IsSelected = !item.IsSelected;
                }
            UpdateSelectedCount();
            }

        private void BtnClearSelect_Click(object sender, RoutedEventArgs e)
            {
            foreach (SlideItem item in slides)
                {
                item.IsSelected = false;
                }
            UpdateSelectedCount();
            }

        private void SlideCheckBox_Changed(object sender, RoutedEventArgs e)
            {
            UpdateSelectedCount();
            }

        private void UpdateSelectedCount()
            {
            int selectedCount = slides.Count(item => item.IsSelected);
            if (selectedCount > 0)
                {
                StatusText.Text = $"已选择 {selectedCount} 页";
                BtnExportSelected.IsEnabled = true;
                }
            else
                {
                StatusText.Text = "就绪";
                BtnExportSelected.IsEnabled = currentPresentation != null;
                }
            }

        private void FontColorPicker_SelectedColorChanged(object sender, FunctionEventArgs<Color> e)
            {
            if (NotesTextBox == null)
                return;

            var selection = NotesTextBox.Selection;
            if (selection.IsEmpty)
                return;

            var brush = new SolidColorBrush(e.Info);
            var textRange = new DocumentTextRange(selection.Start, selection.End);
            textRange.ApplyPropertyValue(TextElement.ForegroundProperty, brush);
            }

        private void BtnFontColor_Click(object sender, RoutedEventArgs e)
            {
            if (NotesTextBox == null)
                return;

            var selection = NotesTextBox.Selection;
            if (selection.IsEmpty)
                return;

            // 获取当前选中文本的颜色
            var textRange = new DocumentTextRange(selection.Start, selection.End);
            var currentBrush = textRange.GetPropertyValue(TextElement.ForegroundProperty) as SolidColorBrush;

            // 创建颜色对话框
            var dialog = new System.Windows.Forms.ColorDialog();
            if (currentBrush != null)
                {
                dialog.Color = System.Drawing.Color.FromArgb(
                    currentBrush.Color.A,
                    currentBrush.Color.R,
                    currentBrush.Color.G,
                    currentBrush.Color.B);
                }

            // 显示颜色对话框
            if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                var color = Color.FromArgb(
                    dialog.Color.A,
                    dialog.Color.R,
                    dialog.Color.G,
                    dialog.Color.B);

                var brush = new SolidColorBrush(color);
                textRange.ApplyPropertyValue(TextElement.ForegroundProperty, brush);

                // 更新颜色指示器
                ColorIndicator.Fill = brush;
                }
            }

        private void ManuscriptWindow_Closed(object sender, EventArgs e)
            {
            // 启用功能区按钮
            MyRibbon RB = Globals.Ribbons.Ribbon1;
            RB.splitButton15.Enabled = true;
            }
        }
    }