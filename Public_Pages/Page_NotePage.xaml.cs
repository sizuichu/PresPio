using HandyControl.Controls;
using Microsoft.Office.Interop.PowerPoint;
using System;
using System.IO;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Documents;

using PowerPoint = Microsoft.Office.Interop.PowerPoint;

using TextRange = System.Windows.Documents.TextRange;

namespace PresPio
    {
    /// <summary>
    /// Page_NotePage.xaml 的交互逻辑
    /// </summary>
    public partial class Page_NotePage
        {
        public Microsoft.Office.Interop.PowerPoint.Application app;
        public Microsoft.Office.Interop.PowerPoint.Application pptApp = Globals.ThisAddIn.Application;
        private Presentation pptPresentation;
        private Slide pptSlide;

        public Page_NotePage()
            {
            InitializeComponent();
            try
                {
                pptPresentation = pptApp.ActivePresentation;
                pptSlide = pptPresentation.Slides[pptApp.ActiveWindow.Selection.SlideRange.SlideNumber];

                // 更新状态栏信息
                UpdateSlideInfo();

                // 将当前幻灯片的备注文本导出到 RichTextBox 控件中
                ExportSlideNotesToRichTextBox();

                // 注册幻灯片选择更改事件处理程序
                pptApp.SlideSelectionChanged += new EApplication_SlideSelectionChangedEventHandler(App_SlideSelectionChanged);

                // 初始化字体列表
                FontFamily.ItemsSource = System.Windows.Media.Fonts.SystemFontFamilies;

                // 绑定格式化按钮事件
                BoldBtn.Click += (s, e) => ApplyTextFormat(TextFormats.Bold);
                ItalicBtn.Click += (s, e) => ApplyTextFormat(TextFormats.Italic);
                UnderlineBtn.Click += (s, e) => ApplyTextFormat(TextFormats.Underline);

                // 初始化撤销重做功能
                RichTextBox1.IsUndoEnabled = true;
                UndoBtn.Click += (s, e) => { if (RichTextBox1.CanUndo) RichTextBox1.Undo(); };
                RedoBtn.Click += (s, e) => { if (RichTextBox1.CanRedo) RichTextBox1.Redo(); };

                // 其他现有初始化代码...
                UpdateSlideInfo();

                // 初始化字体大小改变事件
                FontSize.SelectionChanged += FontSize_SelectionChanged;
                FontFamily.SelectionChanged += FontFamily_SelectionChanged;

                // 初始化查找按钮事件
                FindBtn.Click += FindBtn_Click;
                }
            catch (Exception ex)
                {
                Growl.ErrorGlobal($"初始化时发生错误: {ex.Message}");
                }
            }

        private void App_SlideSelectionChanged(SlideRange SldRange)
            {
            try
                {
                // 检查是否选择了幻灯片页面
                if (pptApp.ActiveWindow.Selection.Type == PowerPoint.PpSelectionType.ppSelectionSlides)
                    {
                    // 获取当前幻灯片
                    pptSlide = pptPresentation.Slides[pptApp.ActiveWindow.Selection.SlideRange.SlideNumber];
                    // 更新状态栏信息
                    UpdateSlideInfo();
                    // 将当前幻灯片的备注文本导出到 RichTextBox 控件中
                    ExportSlideNotesToRichTextBox();
                    }
                }
            catch (Exception ex)
                {
                Growl.ErrorGlobal($"切换幻灯片时发生错误: {ex.Message}");
                }
            }

        private void ExportSlideNotesToRichTextBox()
            {
            RichTextBox1.Document.Blocks.Clear();

            try
                {
                PowerPoint.Shape notesPlaceholder = pptSlide.NotesPage.Shapes.Placeholders[2];
                if (notesPlaceholder != null)
                    {
                    string slideNotes = notesPlaceholder.TextFrame.TextRange.Text;
                    if (string.IsNullOrWhiteSpace(slideNotes))
                        {
                        RichTextBox1.Document.Blocks.Add(new Paragraph(new Run("当前页面无备注")));
                        }
                    else
                        {
                        // 保持原有格式导入
                        RichTextBox1.Document.Blocks.Add(new Paragraph(new Run(slideNotes)));

                        // 更新状态栏信息
                        UpdateSlideInfo();
                        RichTextBox1_TextChanged(null, null);
                        }
                    }
                }
            catch (Exception ex)
                {
                RichTextBox1.Document.Blocks.Add(new Paragraph(new Run($"发生错误: {ex.Message}")));
                }
            }

        private void AddBtn_Click(object sender, RoutedEventArgs e)
            {
            PowerPoint.Selection selection = Globals.ThisAddIn.Application.ActiveWindow.Selection;
            if (selection.Type == PowerPoint.PpSelectionType.ppSelectionShapes)
                {
                StringBuilder stringBuilder = new StringBuilder();
                foreach (PowerPoint.Shape shape in selection.ShapeRange)
                    {
                    if (shape.HasTextFrame == Microsoft.Office.Core.MsoTriState.msoTrue)
                        {
                        string shapeText = shape.TextFrame.TextRange.Text;
                        if (!string.IsNullOrWhiteSpace(shapeText))
                            {
                            // 将形状文本添加到RichTextBox1中
                            RichTextBox1.AppendText(shapeText + Environment.NewLine);
                            }
                        }
                    }

                PowerPoint.TextRange textRange = selection.TextRange;
                if (textRange != null)
                    {
                    // 将选择的文本添加到RichTextBox1中
                    RichTextBox1.AppendText(textRange.Text + Environment.NewLine);
                    }

                string selectedText = stringBuilder.ToString().Trim();
                if (!string.IsNullOrWhiteSpace(selectedText))
                    {
                    PowerPoint.Slide slide = Globals.ThisAddIn.Application.ActiveWindow.View.Slide;
                    if (slide != null)
                        {
                        PowerPoint.Shape notesShape = slide.NotesPage.Shapes.Placeholders[2];
                        if (notesShape != null)
                            {
                            PowerPoint.TextRange notesTextRange = notesShape.TextFrame.TextRange;
                            if (notesTextRange != null)
                                {
                                // 在备注页的文本范围后插入所选文本
                                notesTextRange.InsertAfter(selectedText);
                                // 更新RichTextBox1以包含新插入的文本
                                RichTextBox1.AppendText(notesTextRange.Text + Environment.NewLine);
                                }
                            }
                        }
                    }
                }
            }

        private void DeleBtn_Click(object sender, RoutedEventArgs e)
            {
            // 清除 RichTextBox 控件中的文本
            RichTextBox1.Document.Blocks.Clear();
            }

        private void WriteBtn_Click(object sender, RoutedEventArgs e)
            {
            // 获取当前演示文稿
            PowerPoint.Presentation presentation = Globals.ThisAddIn.Application.ActivePresentation;
            var openFileDialog = new System.Windows.Forms.OpenFileDialog();
            openFileDialog.Filter = "文本文件 (*.txt)|*.txt";
            openFileDialog.Title = "选择之前导出的备注文本文件";

            if (openFileDialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                string notesFilePath = openFileDialog.FileName;

                try
                    {
                    // 读取备注文本文件的内容
                    string allNotesText = File.ReadAllText(notesFilePath);

                    // 分割备注文本为每页的备注
                    string[] pageNotes = allNotesText.Split(new string[] { "Slide " }, StringSplitOptions.RemoveEmptyEntries);

                    // 循环遍历每页的备注
                    for (int i = 1 ; i <= pageNotes.Length ; i++)
                        {
                        // 查找对应编号的幻灯片
                        PowerPoint.Slide slide = presentation.Slides[i];

                        // 获取备注文本
                        string notesText = pageNotes[i - 1].Trim();

                        // 如果幻灯片有备注页
                        if (slide.HasNotesPage == Microsoft.Office.Core.MsoTriState.msoTrue)
                            {
                            // 获取备注页
                            var notesPage = slide.NotesPage;

                            // 设置备注文本
                            notesPage.Shapes[2].TextFrame.TextRange.Text = notesText;
                            }
                        else
                            {
                            // 如果幻灯片没有备注页，则创建一个
                            var notesSlide = slide.NotesPage;
                            PowerPoint.TextRange textRange = notesSlide.Shapes[2].TextFrame.TextRange;
                            textRange.Text = notesText;
                            }
                        }

                    HandyControl.Controls.Growl.SuccessGlobal("备注导入完成。");
                    }
                catch (Exception ex)
                    {
                    HandyControl.Controls.Growl.ErrorGlobal("导入备注时发生错误：" + ex.Message);
                    }
                }
            else
                {
                HandyControl.Controls.Growl.WarningGlobal("未选择备注文本文件。");
                }
            }

        //导入TXT文本
        private void ImportFromTxt(string filePath)
            {
            if (!string.IsNullOrWhiteSpace(filePath))
                {
                using (StreamReader reader = new StreamReader(filePath))
                    {
                    string fileContent = reader.ReadToEnd();
                    RichTextBox1.Document.Blocks.Add(new System.Windows.Documents.Paragraph(new Run(fileContent)));
                    }
                }
            }

        //导入word文件
        private void ImportAndParseWord(string filePath)
            {
            // 创建 Word 应用程序对象
            Microsoft.Office.Interop.Word.Application wordApp = new Microsoft.Office.Interop.Word.Application();

            // 打开指定的 Word 文档
            Microsoft.Office.Interop.Word.Document wordDoc = wordApp.Documents.Open(filePath);

            // 获取文档的内容
            Microsoft.Office.Interop.Word.Range range = wordDoc.Content;

            // 将内容转换为纯文本格式
            string text = range.Text;

            // 将文本显示在 RichTextBox 中
            RichTextBox1.Document.Blocks.Add(new System.Windows.Documents.Paragraph(new Run(text)));

            // 关闭 Word 文档和应用程序对象
            wordDoc.Close();
            wordApp.Quit();
            }

        private void SaveBtn_Click(object sender, RoutedEventArgs e)
            {
            try
                {
                var range = new TextRange(RichTextBox1.Document.ContentStart, RichTextBox1.Document.ContentEnd);
                string text = range.Text;

                if (pptSlide != null && pptSlide.NotesPage != null)
                    {
                    var placeholders = pptSlide.NotesPage.Shapes.Placeholders;
                    if (placeholders.Count > 1)
                        {
                        var textFrame = placeholders[2].TextFrame;
                        if (textFrame != null)
                            {
                            textFrame.TextRange.Text = text;
                            Growl.SuccessGlobal("保存成功");
                            }
                        }
                    }
                }
            catch (Exception ex)
                {
                Growl.ErrorGlobal($"保存失败: {ex.Message}");
                }
            }

        private void importBtn_Click(object sender, RoutedEventArgs e)
            {
            try
                {
                System.Windows.Forms.FolderBrowserDialog folderBrowserDialog = new System.Windows.Forms.FolderBrowserDialog();
                folderBrowserDialog.Description = "选择要保存备注文件的文件夹路径";

                if (folderBrowserDialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                    {
                    string saveFolderPath = folderBrowserDialog.SelectedPath;

                    // 获取演示文稿的文件名（不含扩展名）
                    string pptFileName = Path.GetFileNameWithoutExtension(pptPresentation.Name);

                    // 构造备注文件的完整路径
                    string notesFilePath = Path.Combine(saveFolderPath, pptFileName + "_Notes.txt");

                    // 创建一个字符串来存储所有备注文本
                    string allNotesText = "";

                    // 循环遍历每个幻灯片
                    foreach (PowerPoint.Slide slide in pptPresentation.Slides)
                        {
                        // 如果幻灯片有备注页
                        if (slide.NotesPage != null && slide.NotesPage.Shapes.Placeholders.Count > 1)
                            {
                            // 获取幻灯片的备注文本
                            string slideNotesText = slide.NotesPage.Shapes.Placeholders[2].TextFrame.TextRange.Text;

                            // 将幻灯片的备注文本添加到总的备注文本中
                            allNotesText += $"幻灯片 {slide.SlideNumber} 备注:\n{slideNotesText}\n\n";
                            }
                        }

                    // 将所有备注文本写入到文件中
                    File.WriteAllText(notesFilePath, allNotesText);

                    Growl.SuccessGlobal($"备注导出完成。文件保存在：{notesFilePath}");
                    }
                else
                    {
                    Growl.WarningGlobal("未选择文件夹路径。");
                    }
                }
            catch (Exception ex)
                {
                Growl.ErrorGlobal($"导出失败: {ex.Message}");
                }
            }

        private void Window_Closed(object sender, EventArgs e)
            {
            try
                {
                // 取消事件订阅
                pptApp.SlideSelectionChanged -= App_SlideSelectionChanged;

                // 启用功能区按钮
                MyRibbon RB = Globals.Ribbons.Ribbon1;
                RB.button124.Enabled = true;
                }
            catch (Exception ex)
                {
                Growl.ErrorGlobal($"关闭窗口时发生错误: {ex.Message}");
                }
            }

        private void RichTextBox1_TextChanged(object sender, TextChangedEventArgs e)
            {
            // 更新字数统计
            var textRange = new TextRange(RichTextBox1.Document.ContentStart, RichTextBox1.Document.ContentEnd);
            int wordCount = textRange.Text.Length;
            WordCountText.Text = $"字数: {wordCount - 2}";
            }

        private void UpdateSlideInfo()
            {
            try
                {
                if (pptSlide != null)
                    {
                    int currentSlide = pptSlide.SlideNumber;
                    int totalSlides = pptPresentation.Slides.Count;
                    SlideInfoText.Text = $"当前幻灯片: 第{currentSlide}/{totalSlides}页";
                    }
                }
            catch (Exception ex)
                {
                Growl.ErrorGlobal($"更新幻灯片信息时发生错误: {ex.Message}");
                SlideInfoText.Text = "当前幻灯片: 获取失败";
                }
            }

        private enum TextFormats
            { Bold, Italic, Underline }

        private void ApplyTextFormat(TextFormats format)
            {
            var selection = RichTextBox1.Selection;
            if (selection.IsEmpty) return;

            var textRange = new System.Windows.Documents.TextRange(selection.Start, selection.End);

            switch (format)
                {
                case TextFormats.Bold:
                    textRange.ApplyPropertyValue(TextElement.FontWeightProperty,
                        ((ToggleButton)BoldBtn).IsChecked == true ? FontWeights.Bold : FontWeights.Normal);
                    break;

                case TextFormats.Italic:
                    textRange.ApplyPropertyValue(TextElement.FontStyleProperty,
                        ((ToggleButton)ItalicBtn).IsChecked == true ? FontStyles.Italic : FontStyles.Normal);
                    break;

                case TextFormats.Underline:
                    textRange.ApplyPropertyValue(Inline.TextDecorationsProperty,
                        ((ToggleButton)UnderlineBtn).IsChecked == true ? TextDecorations.Underline : null);
                    break;
                }
            }

        private void FontSize_SelectionChanged(object sender, SelectionChangedEventArgs e)
            {
            if (RichTextBox1.Selection.IsEmpty || FontSize.SelectedItem == null) return;

            var size = double.Parse(((ComboBoxItem)FontSize.SelectedItem).Content.ToString());
            var textRange = new System.Windows.Documents.TextRange(RichTextBox1.Selection.Start, RichTextBox1.Selection.End);
            textRange.ApplyPropertyValue(TextElement.FontSizeProperty, size);
            }

        private void FontFamily_SelectionChanged(object sender, SelectionChangedEventArgs e)
            {
            if (RichTextBox1.Selection.IsEmpty || FontFamily.SelectedItem == null) return;

            var fontFamily = (System.Windows.Media.FontFamily)FontFamily.SelectedItem;
            var textRange = new System.Windows.Documents.TextRange(RichTextBox1.Selection.Start, RichTextBox1.Selection.End);
            textRange.ApplyPropertyValue(TextElement.FontFamilyProperty, fontFamily);
            }

        private void FindBtn_Click(object sender, RoutedEventArgs e)
            {
            string searchText = SearchBox.Text;
            if (string.IsNullOrEmpty(searchText))
                {
                Growl.WarningGlobal("请输入要查找的内容");
                return;
                }

            TextRange documentRange = new TextRange(RichTextBox1.Document.ContentStart, RichTextBox1.Document.ContentEnd);
            string text = documentRange.Text;
            int index = text.IndexOf(searchText, StringComparison.OrdinalIgnoreCase);

            if (index == -1)
                {
                Growl.WarningGlobal("未找到匹配内容");
                return;
                }

            // 查找文本位置
            TextPointer start = documentRange.Start.GetPositionAtOffset(index);
            TextPointer end = start.GetPositionAtOffset(searchText.Length);

            // 选中找到的文本
            RichTextBox1.Selection.Select(start, end);
            RichTextBox1.Focus();
            }
        }
    }