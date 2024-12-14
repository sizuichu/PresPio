using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Forms;
using HandyControl.Controls;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using MessageBox = System.Windows.Forms.MessageBox;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace PresPio
    {
    public partial class Wpf_Mockup
        {
        public PowerPoint.Application app; //加载PPT项目
        public string MockupUrl = null;

        public Wpf_Mockup()
            {
            app = Globals.ThisAddIn.Application;
            app.WindowSelectionChange += new EApplication_WindowSelectionChangeEventHandler(Application_WindowSelectionChange);
            InitializeComponent();
            LoadImg();
            LoadNum(); //加载页面计数
            SaveSlidesAsImages(app);
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
                        mockupPath = openFileDialog.FileName; // 赋值给 mockupUrl
                        }
                    }
                }

            if (!string.IsNullOrEmpty(mockupPath))
                {
                Properties.Settings.Default.MockupUrl = mockupPath;
                Properties.Settings.Default.Save();
                }

            return mockupPath; // 返回选择的文件路径
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

        public static void SaveSlidesAsImages(PowerPoint.Application app)
            {
            string location = AppDomain.CurrentDomain.BaseDirectory;
            string mockupPath = GetMockup();

            if (!File.Exists(mockupPath))
                {
                Console.WriteLine($"文件 {mockupPath} 不存在。");
                return;
                }

            // 保存当前活动演示文稿的状态
            Presentation currentPresentation = app.ActivePresentation;
            if (currentPresentation != null)
                {
                currentPresentation.Save(); // 保存当前演示文稿
                }

            Presentation pre = app.Presentations.Open(mockupPath);
            string mocKupPath = Path.Combine(location, "MocKup");
            Directory.CreateDirectory(mocKupPath);

            int num = pre.Slides.Count;
            bool needsExport = Enumerable.Range(1, num).Any(i => !File.Exists(Path.Combine(mocKupPath, $"Slide_{i}.png")));

            if (needsExport)
                {
                for (int i = 1 ; i <= num ; i++)
                    {
                    string tempImagePath = Path.Combine(mocKupPath, $"Slide_{i}.png");
                    float slideWidth = pre.SlideMaster.Width;
                    float slideHeight = pre.SlideMaster.Height;
                    pre.Slides[i].Export(tempImagePath, "PNG", (int)slideWidth, (int)slideHeight);
                    }
                Console.WriteLine("所有幻灯片已成功导出为图片！");
                }
            else  
                {
                Console.WriteLine("图片已存在，跳过导出。");
                }

            pre.Close(); // 关闭新打开的演示文稿
            }

        public void Application_WindowSelectionChange(Selection Sel)
            {
            LoadNum();
            }

        public void LoadNum()
            {
            var app = Globals.ThisAddIn.Application;
            int selectedCount = app.ActiveWindow?.Selection?.SlideRange?.Count ?? 0;
            LabelNum.Content = $"当前选择页面数量: {selectedCount}";
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
            string mocKupPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "MocKup");
            var imageFiles = Directory.GetFiles(mocKupPath, "*.png");
            CoverFlowMain.AddRange(imageFiles.Select(imagePath => new Uri(imagePath)));
            CoverFlowMain.PageIndex = 1;
            }

        private void Button_Click(object sender, RoutedEventArgs e)
            {
            var presentation = app.ActivePresentation;

            if (presentation == null || presentation.Slides.Count == 0)
                {
                Growl.ErrorGlobal("没有可导出的幻灯片！");
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

            int pageIndex = CoverFlowMain.PageIndex + 1;
            InsertTemplateAndImages(app, pageIndex, imagePaths.ToArray());

            if (presentation.Slides.Count > 0)
                {
                Slide lastSlide = presentation.Slides[presentation.Slides.Count];
                app.ActiveWindow.View.GotoSlide(lastSlide.SlideIndex);
                }

            Growl.SuccessGlobal("生成样机成功！");
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
            Presentation pre = app.ActivePresentation;
            string templatePath = Properties.Settings.Default.MockupUrl;

            if (!File.Exists(templatePath))
                {
                Growl.WarningGlobal("默认样机文件不存在，请修复！");
                return;
                }

            int num = pre.Slides.Count;
            pre.Slides.InsertFromFile(templatePath, num, index, index);
            Slide newSlide = pre.Slides[num + 1];
            InsertImagesToSlide(newSlide, imagePaths);
            newSlide.Tags.Add("样机", "母版样机");
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
                openFileDialog.Title = "选择演示文稿文件";

                if (openFileDialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                    {
                    mockupUrl = openFileDialog.FileName;
                    Properties.Settings.Default.MockupUrl = mockupUrl;
                    Properties.Settings.Default.Save();
                    }
                else
                    {
                    return;
                    }
                }

            SaveSlidesAsImages(app);
            }
        }
    }
