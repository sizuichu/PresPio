using System;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Net.NetworkInformation;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Forms;
using HandyControl.Controls;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.Office.Tools.Ribbon;
using PresPio.Properties;
using static PresPio.ThisAddIn;
using Color = System.Drawing.Color;
using Control = System.Windows.Forms.Control;
using MessageBox = System.Windows.Forms.MessageBox;
using Growl = HandyControl.Controls.Growl;
using Office = Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

using Shape = Microsoft.Office.Interop.PowerPoint.Shape;

namespace PresPio
    {
    public partial class MyRibbon
        {
        public PowerPoint.Application app; //加载PPT项目

        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
            {
            new App().InitializeComponent(); //加载WPF

            app = Globals.ThisAddIn.Application;//初始化项目

            //设置
            checkBox1.Checked = Properties.Settings.Default.Replace;//保留格式选项
            checkBox2.Checked = Properties.Settings.Default.TextBoxAuto;                                                    //菜单设置
            //菜单设置
            Globals.Ribbons.Ribbon1.group2.Visible = Properties.Settings.Default.Group2;
            Globals.Ribbons.Ribbon1.group3.Visible = Properties.Settings.Default.Group3;
            Globals.Ribbons.Ribbon1.group4.Visible = Properties.Settings.Default.Group4;
            Globals.Ribbons.Ribbon1.group5.Visible = Properties.Settings.Default.Group5;
            Globals.Ribbons.Ribbon1.tab2.Visible = Properties.Settings.Default.Group6;
            //Globals.Ribbons.Ribbon1.group7.Visible = Properties.Settings.Default.Group7;
            }

        //以下为公用字段
        /// <summary>
        /// Temp_Path为临时文件夹所在
        /// </summary>
        public string Temp_Path()
            {
            string Temp_Path = Properties.Settings.Default.Temp_Path;
            return Temp_Path;
            }

        //以下为事件
        private void button2_Click(object sender, RibbonControlEventArgs e)
            {
            Wpf_Colortheif wpf_Colortheif = new Wpf_Colortheif();
            wpf_Colortheif.Show();
            }

        private void button3_Click(object sender, RibbonControlEventArgs e)
            {
            app.CommandBars.ExecuteMso("ThemeColorsCreateNew");//新建配色方案
            }

        private void button4_Click(object sender, RibbonControlEventArgs e)
            {
            app.CommandBars.ExecuteMso("ThemeFontsCreateNew");//新建字体方案
            }

        private void button5_Click(object sender, RibbonControlEventArgs e)
            {
            MyFunction F = new MyFunction();
            Selection sel = app.ActiveWindow.Selection;
            sel.ShapeRange.Fill.ForeColor.RGB = F.RGB2Int(237, 125, 49);
            sel.ShapeRange.Fill.Transparency = 0.5f;
            }

        private void button6_Click(object sender, RibbonControlEventArgs e)
            {
            string Location = AppDomain.CurrentDomain.BaseDirectory; //获取插件安装位置
            System.Diagnostics.Process.Start("Explorer.exe", Location); //打开安装位置
            }

        private void button1_Click(object sender, RibbonControlEventArgs e)
            {
            Wpf_About wpf_About = new Wpf_About();
            wpf_About.ShowDialog();
            }

        private void button8_Click(object sender, RibbonControlEventArgs e)
            {
            PowerPoint.Selection sel = app.ActiveWindow.Selection;
            PowerPoint.Slide slide = app.ActiveWindow.View.Slide;
            if (sel.Type == PowerPoint.PpSelectionType.ppSelectionNone)
                {
                Growl.Warning("请选中要合并的文本框，或者幻灯片", "温馨提示");
                }
            else if (sel.Type == PowerPoint.PpSelectionType.ppSelectionSlides)
                {
                PowerPoint.Shape text = slide.Shapes.AddTextbox(Office.MsoTextOrientation.msoTextOrientationHorizontal, -200, 0, 200, 200);
                text.TextFrame.AutoSize = PowerPoint.PpAutoSize.ppAutoSizeShapeToFitText;
                text.Name = "textcount";
                text.TextFrame.TextRange.Text = "";
                foreach (PowerPoint.Slide item in sel.SlideRange)
                    {
                    for (int i = 1 ; i <= item.Shapes.Count ; i++)
                        {
                        PowerPoint.Shape shape = item.Shapes[i];
                        if (shape.Type == Office.MsoShapeType.msoGroup)
                            {
                            for (int j = 1 ; j <= shape.GroupItems.Count ; j++)
                                {
                                if (shape.GroupItems[j].HasTextFrame == Office.MsoTriState.msoTrue)
                                    {
                                    if (shape.GroupItems[j].Name != "textcount")
                                        {
                                        text.TextFrame.TextRange.Text = text.TextFrame.TextRange.Text + Environment.NewLine + shape.GroupItems[j].TextFrame.TextRange.Text;
                                        }
                                    }
                                }
                            }
                        else
                            {
                            if (shape.HasTextFrame == Office.MsoTriState.msoTrue)
                                {
                                if (shape.Name != "textcount")
                                    {
                                    text.TextFrame.TextRange.Text = text.TextFrame.TextRange.Text + Environment.NewLine + shape.TextFrame.TextRange.Text;
                                    }
                                }
                            }
                        }
                    }
                text.Select();
                }
            else if (sel.Type == PowerPoint.PpSelectionType.ppSelectionShapes)
                {
                PowerPoint.Shape text = slide.Shapes.AddTextbox(Office.MsoTextOrientation.msoTextOrientationHorizontal, -200, 0, 200, 200);
                text.TextFrame.AutoSize = PowerPoint.PpAutoSize.ppAutoSizeShapeToFitText;
                text.Name = "textcount";
                text.TextFrame.TextRange.Text = "";
                PowerPoint.ShapeRange range = sel.ShapeRange;
                if (sel.HasChildShapeRange)
                    {
                    range = sel.ChildShapeRange;
                    }
                for (int i = 1 ; i <= range.Count ; i++)
                    {
                    PowerPoint.Shape shape = range[i];
                    if (shape.Type == Office.MsoShapeType.msoGroup)
                        {
                        for (int j = 1 ; j <= shape.GroupItems.Count ; j++)
                            {
                            if (shape.GroupItems[j].HasTextFrame == Office.MsoTriState.msoTrue)
                                {
                                if (shape.GroupItems[j].Name != "textcount")
                                    {
                                    text.TextFrame.TextRange.Text = text.TextFrame.TextRange.Text + Environment.NewLine + shape.GroupItems[j].TextFrame.TextRange.Text;
                                    }
                                }
                            }
                        }
                    else
                        {
                        if (shape.HasTextFrame == Office.MsoTriState.msoTrue)
                            {
                            if (shape.Name != "textcount")
                                {
                                text.TextFrame.TextRange.Text = text.TextFrame.TextRange.Text + Environment.NewLine + shape.TextFrame.TextRange.Text;
                                }
                            }
                        }
                    }
                }
            }

        private void button9_Click(object sender, RibbonControlEventArgs e)
            {
            PowerPoint.Selection sel = app.ActiveWindow.Selection;
            PowerPoint.Slide slide = app.ActiveWindow.View.Slide;
            if (sel.Type == PowerPoint.PpSelectionType.ppSelectionNone)
                {
                Growl.Warning("请选中至少1个文本框", "温馨提示");
                }
            else
                {
                PowerPoint.ShapeRange range = sel.ShapeRange;
                if (sel.HasChildShapeRange)
                    {
                    range = sel.ChildShapeRange;
                    }
                int count = range.Count;
                for (int i = 1 ; i <= count ; i++)
                    {
                    PowerPoint.Shape shape = range[i];
                    string txt = shape.TextEffect.Text;
                    if (txt.Contains("\r") || txt.Contains("\v"))
                        {
                        String[] arr = txt.Split(char.Parse("\r"), char.Parse("\v")).ToArray();
                        int tcount = arr.Count();
                        shape.PickUp();
                        for (int j = 1 ; j <= tcount ; j++)
                            {
                            PowerPoint.Shape text = slide.Shapes.AddTextbox(Office.MsoTextOrientation.msoTextOrientationHorizontal, shape.Left + shape.Width, shape.Top + shape.Height * (j - 1) / tcount, shape.Width, shape.Height);
                            text.TextFrame.TextRange.Text = arr[j - 1];
                            text.Apply();
                            text.TextFrame.AutoSize = PowerPoint.PpAutoSize.ppAutoSizeShapeToFitText;
                            }
                        }
                    else
                        {
                        Growl.Warning("存在没有分段的文本框", "温馨提示");
                        }
                    }
                }
            }

        private void button10_Click(object sender, RibbonControlEventArgs e)
            {
            PowerPoint.Selection sel = app.ActiveWindow.Selection;
            PowerPoint.Slide slide = app.ActiveWindow.View.Slide;
            if (sel.Type == PowerPoint.PpSelectionType.ppSelectionNone)
                {
                Growl.Warning("请选中至少1个文本框", "温馨提示");
                }
            else
                {
                PowerPoint.ShapeRange range = sel.ShapeRange;
                if (sel.HasChildShapeRange)
                    {
                    range = sel.ChildShapeRange;
                    }
                int count = range.Count;
                for (int i = 1 ; i <= count ; i++)
                    {
                    PowerPoint.Shape shape = range[i];
                    string txt = shape.TextEffect.Text;
                    int tcount = txt.Length;
                    shape.PickUp();
                    for (int j = 1 ; j <= tcount ; j++)
                        {
                        PowerPoint.Shape text = slide.Shapes.AddTextbox(Office.MsoTextOrientation.msoTextOrientationHorizontal, shape.Left + shape.Width + 24 * (j - 1), shape.Top, 24, 20);
                        text.TextFrame.TextRange.Text = txt.Substring(j - 1, 1);
                        text.Apply();
                        text.TextFrame2.TextRange.Font.Size = shape.TextFrame2.TextRange.Characters[j].Font.Size;
                        text.TextFrame2.TextRange.Font.Name = shape.TextFrame2.TextRange.Characters[j].Font.Name;
                        text.TextFrame2.TextRange.Font.NameFarEast = shape.TextFrame2.TextRange.Characters[j].Font.NameFarEast;
                        text.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = shape.TextFrame2.TextRange.Characters[j].Font.Fill.ForeColor.RGB;
                        text.TextFrame2.TextRange.Font.Fill.Transparency = shape.TextFrame2.TextRange.Characters[j].Font.Fill.Transparency;
                        text.TextFrame.AutoSize = PowerPoint.PpAutoSize.ppAutoSizeShapeToFitText;
                        }
                    }
                }
            }

        private void button11_Click(object sender, RibbonControlEventArgs e)
            {
            PowerPoint.Selection sel = app.ActiveWindow.Selection;
            PowerPoint.Slide slide = app.ActiveWindow.View.Slide;
            if (sel.Type == PowerPoint.PpSelectionType.ppSelectionNone || sel.TextRange.Count < 2)
                {
                Growl.Warning("请选中至少2个文本框", "温馨提示");
                }
            else
                {
                PowerPoint.ShapeRange range = sel.ShapeRange;
                if (sel.HasChildShapeRange)
                    {
                    range = sel.ChildShapeRange;
                    }
                PowerPoint.Shape text = slide.Shapes.AddTextbox(Office.MsoTextOrientation.msoTextOrientationHorizontal, sel.ShapeRange[sel.ShapeRange.Count].Left + sel.ShapeRange[sel.ShapeRange.Count].Width, sel.ShapeRange[1].Top, sel.ShapeRange[1].Width * sel.ShapeRange.Count / 2, sel.ShapeRange[1].Height);
                PowerPoint.TextFrame2 tframe = text.TextFrame2;
                int count = sel.ShapeRange.Count;
                for (int i = 1 ; i <= count ; i++)
                    {
                    tframe.TextRange.Text = tframe.TextRange.Text + range[i].TextFrame2.TextRange.Text;
                    tframe.AutoSize = Office.MsoAutoSize.msoAutoSizeShapeToFitText;
                    }
                }
            }

        private void button12_Click(object sender, RibbonControlEventArgs e)
            {
            }

        private void button13_Click(object sender, RibbonControlEventArgs e)
            {
            PowerPoint.Selection sel = Globals.ThisAddIn.Application.ActiveWindow.Selection;
            if (sel.Type == PowerPoint.PpSelectionType.ppSelectionShapes)
                {
                PowerPoint.ShapeRange shapeRange = sel.ShapeRange;
                PowerPoint.Slide slide = Globals.ThisAddIn.Application.ActiveWindow.View.Slide;
                float w = slide.Master.Width;
                float h = slide.Master.Height;

                foreach (PowerPoint.Shape shape in shapeRange)
                    {
                    // 检查形状是否为可以进行布尔操作的类型
                    if (shape.Type == MsoShapeType.msoAutoShape ||
                        shape.Type == MsoShapeType.msoFreeform ||
                        shape.Type == MsoShapeType.msoTextBox ||
                        shape.Type == MsoShapeType.msoPicture ||
                        shape.Type == MsoShapeType.msoGroup ||
                        shape.Type == MsoShapeType.msoPlaceholder ||
                        shape.Type == MsoShapeType.msoEmbeddedOLEObject)
                        {
                        // 为所有形状添加相交矩形
                        PowerPoint.Shape newShape = slide.Shapes.AddShape(Office.MsoAutoShapeType.msoShapeRectangle, 0f, 0f, w, h);
                        shape.Select(Office.MsoTriState.msoTrue); // 选择当前形状
                        newShape.Select(Office.MsoTriState.msoFalse); // 取消选择新添加的矩形

                        string strCmd = "ShapesIntersect";
                        Globals.ThisAddIn.Application.CommandBars.ExecuteMso(strCmd); // 执行相交操作
                        }
                    }
                }
            }

        private void button14_Click(object sender, RibbonControlEventArgs e)
            {
            }

        private void splitButton1_Click(object sender, RibbonControlEventArgs e)
            {
            Wpf_Publish wpf_Publish = new Wpf_Publish();
            wpf_Publish.ShowDialog();
            }

        public static Microsoft.Office.Tools.CustomTaskPane taskPane;

        private void button15_Click(object sender, RibbonControlEventArgs e)
            {
            }

        private void splitButton2_Click(object sender, RibbonControlEventArgs e)
            {
            button16_Click(sender, e);
            }

        private void button16_Click(object sender, RibbonControlEventArgs e)
            {
            PowerPoint.Selection sel = app.ActiveWindow.Selection;
            PowerPoint.Slide slide = app.ActiveWindow.View.Slide;

            if (sel.Type == PowerPoint.PpSelectionType.ppSelectionNone)
                {
                Growl.WarningGlobal("请选中元素或幻灯片，且做好备份！");
                }
            else if (sel.Type == PowerPoint.PpSelectionType.ppSelectionShapes)
                {
                ProcessShapes(sel.ShapeRange, slide);
                }
            else if (sel.Type == PowerPoint.PpSelectionType.ppSelectionSlides)
                {
                ProcessSlides(sel.SlideRange);
                }

            void ProcessShapes(PowerPoint.ShapeRange shapeRange, PowerPoint.Slide targetSlide)
                {
                foreach (PowerPoint.Shape item in shapeRange)
                    {
                    item.Copy();
                    PowerPoint.Shape pic = targetSlide.Shapes.PasteSpecial(PowerPoint.PpPasteDataType.ppPastePNG)[1];
                    CenterShape(pic, item);
                    item.Delete();
                    }
                }

            void ProcessSlides(PowerPoint.SlideRange slideRange)
                {
                foreach (PowerPoint.Slide item in slideRange)
                    {
                    for (int i = item.Shapes.Count ; i >= 1 ; i--)
                        {
                        item.Shapes[i].Copy();
                        PowerPoint.Shape pic = item.Shapes.PasteSpecial(PowerPoint.PpPasteDataType.ppPastePNG)[1];
                        CenterShape(pic, item.Shapes[i]);
                        item.Shapes[i].Delete();
                        }
                    }
                Growl.SuccessGlobal("已将所选页面中的所有元素转为png图片");
                }

            void CenterShape(PowerPoint.Shape pic, PowerPoint.Shape original)
                {
                pic.ScaleHeight(1f, Office.MsoTriState.msoTrue, Office.MsoScaleFrom.msoScaleFromMiddle);
                pic.Left = original.Left + original.Width / 2 - pic.Width / 2;
                pic.Top = original.Top + original.Height / 2 - pic.Height / 2;
                }
            }

        private void button17_Click(object sender, RibbonControlEventArgs e)
            {
            PowerPoint.Selection sel = app.ActiveWindow.Selection;
            PowerPoint.Slide slide = app.ActiveWindow.View.Slide;
            if (sel.Type == PowerPoint.PpSelectionType.ppSelectionNone)
                {
                Growl.WarningGlobal("请选中元素或幻灯片，且做好备份！");
                }
            else if (sel.Type == PowerPoint.PpSelectionType.ppSelectionShapes)
                {
                PowerPoint.ShapeRange range = sel.ShapeRange;
                foreach (PowerPoint.Shape item in range)
                    {
                    item.Copy();
                    PowerPoint.Shape pic = slide.Shapes.PasteSpecial(PowerPoint.PpPasteDataType.ppPasteJPG)[1];
                    pic.ScaleHeight(1f, Office.MsoTriState.msoTrue, Office.MsoScaleFrom.msoScaleFromMiddle);
                    pic.Left = item.Left + item.Width / 2 - pic.Width / 2;
                    pic.Top = item.Top + item.Height / 2 - pic.Height / 2;
                    item.Delete();
                    pic.Select();
                    }
                }
            else if (sel.Type == PowerPoint.PpSelectionType.ppSelectionSlides)
                {
                PowerPoint.SlideRange sliderange = sel.SlideRange;
                foreach (PowerPoint.Slide item in sliderange)
                    {
                    for (int i = item.Shapes.Count ; i >= 1 ; i--)
                        {
                        item.Shapes[i].Copy();
                        PowerPoint.Shape pic = item.Shapes.PasteSpecial(PowerPoint.PpPasteDataType.ppPasteJPG)[1];
                        pic.ScaleHeight(1f, Office.MsoTriState.msoTrue, Office.MsoScaleFrom.msoScaleFromMiddle);
                        pic.Left = item.Shapes[i].Left + item.Shapes[i].Width / 2 - pic.Width / 2;
                        pic.Top = item.Shapes[i].Top + item.Shapes[i].Height / 2 - pic.Height / 2;
                        if (!sel.HasChildShapeRange)
                            {
                            item.Shapes[i].Delete();
                            }
                        }
                    }
                Growl.SuccessGlobal("已将所选页面中的所有元素转为JPG图片");
                }
            }

        private void button18_Click(object sender, RibbonControlEventArgs e)
            {
            var sel = app.ActiveWindow.Selection;
            var slide = app.ActiveWindow.View.Slide;

            switch (sel.Type)
                {
                case PowerPoint.PpSelectionType.ppSelectionNone:
                    Growl.WarningGlobal("可选中形状和图片元素导出为JPG；选中多页幻灯片，只导出其中的图片元素");
                    break;

                case PowerPoint.PpSelectionType.ppSelectionShapes:
                    var range = sel.ShapeRange;
                    var name = Path.GetFileNameWithoutExtension(app.ActivePresentation.Name);
                    var cPath = Path.Combine(app.ActivePresentation.Path, $"{name} 的元素");

                    Directory.CreateDirectory(cPath);

                    var tasks = new List<Task>();
                    for (int i = 1 ; i <= range.Count ; i++)
                        {
                        var shape = range[i];
                        var dir = new DirectoryInfo(cPath);
                        var k = dir.GetFiles().Length + i;
                        var shname = $"{name}_{k}";

                        var task = Task.Run(() =>
                        {
                            shape.Export(Path.Combine(cPath, $"{shname}.jpg"), PowerPoint.PpShapeFormat.ppShapeFormatJPG);
                        });
                        tasks.Add(task);
                        }
                    Task.WaitAll(tasks.ToArray());

                    Process.Start("Explorer.exe", cPath);
                    break;

                case PowerPoint.PpSelectionType.ppSelectionSlides:
                    name = Path.GetFileNameWithoutExtension(app.ActivePresentation.Name);
                    cPath = Path.Combine(app.ActivePresentation.Path, $"{name} 的元素");

                    Directory.CreateDirectory(cPath);

                    tasks = new List<Task>();
                    foreach (PowerPoint.Slide item in sel.SlideRange)
                        {
                        for (int i = 1 ; i <= item.Shapes.Count ; i++)
                            {
                            var shape = item.Shapes[i];
                            if (shape.Type == Office.MsoShapeType.msoPicture)
                                {
                                var dir = new DirectoryInfo(cPath);
                                var k = dir.GetFiles().Length + i;
                                var shname = $"{name}_{k}";

                                var task = Task.Run(() =>
                                {
                                    shape.Export(Path.Combine(cPath, $"{shname}.jpg"), PowerPoint.PpShapeFormat.ppShapeFormatJPG);
                                });
                                tasks.Add(task);
                                }
                            }
                        }
                    Task.WaitAll(tasks.ToArray());

                    Process.Start("Explorer.exe", cPath);
                    break;
                }
            }

        private void button19_Click(object sender, RibbonControlEventArgs e)
            {
            PowerPoint.Selection sel = app.ActiveWindow.Selection;
            PowerPoint.Slide slide = app.ActiveWindow.View.Slide;
            string name = app.ActivePresentation.Name.Replace(".pptx", "").Replace(".ppt", "");
            string cPath = Path.Combine(app.ActivePresentation.Path, name + " 的元素");
            Directory.CreateDirectory(cPath);

            if (sel.Type == PowerPoint.PpSelectionType.ppSelectionNone)
                {
                Growl.WarningGlobal("可选中形状和图片元素导出为Png；选中多页幻灯片，只导出其中的图片元素");
                }
            else if (sel.Type == PowerPoint.PpSelectionType.ppSelectionShapes)
                {
                PowerPoint.ShapeRange range = sel.ShapeRange;

                List<Thread> threads = new List<Thread>();
                for (int i = 1 ; i <= range.Count ; i++)
                    {
                    PowerPoint.Shape shape = range[i];
                    string shname = name + "_" + i;
                    Thread thread = new Thread(() =>
                    {
                        shape.Export(Path.Combine(cPath, shname + ".png"), PowerPoint.PpShapeFormat.ppShapeFormatPNG);
                    });
                    threads.Add(thread);
                    thread.Start();
                    }

                foreach (Thread thread in threads)
                    {
                    thread.Join();
                    }

                System.Diagnostics.Process.Start("Explorer.exe", cPath);
                }
            else if (sel.Type == PowerPoint.PpSelectionType.ppSelectionSlides)
                {
                List<Thread> threads = new List<Thread>();
                foreach (PowerPoint.Slide item in sel.SlideRange)
                    {
                    foreach (PowerPoint.Shape shape in item.Shapes)
                        {
                        if (shape.Type == Office.MsoShapeType.msoPicture)
                            {
                            string shname = name + "_" + item.SlideIndex + "_" + shape.Id;
                            Thread thread = new Thread(() =>
                            {
                                shape.Export(Path.Combine(cPath, shname + ".png"), PowerPoint.PpShapeFormat.ppShapeFormatPNG);
                            });
                            threads.Add(thread);
                            thread.Start();
                            }
                        }
                    }

                foreach (Thread thread in threads)
                    {
                    thread.Join();
                    }

                System.Diagnostics.Process.Start("Explorer.exe", cPath);
                }
            }

        private void button20_Click(object sender, RibbonControlEventArgs e)
            {
            Selection selection = Globals.ThisAddIn.Application.ActiveWindow.Selection;
            if (selection.Type == PpSelectionType.ppSelectionNone)
                {
                Growl.WarningGlobal("可选中形状和图片元素导出为EMF；选中多页幻灯片，只导出其中的图片元素");
                return;
                }

            string presentationName = Path.GetFileNameWithoutExtension(Globals.ThisAddIn.Application.ActivePresentation.Name);
            string exportFolderPath = Globals.ThisAddIn.Application.ActivePresentation.Path + "\\" + presentationName + " 的元素\\";
            if (!Directory.Exists(exportFolderPath))
                {
                Directory.CreateDirectory(exportFolderPath);
                }

            switch (selection.Type)
                {
                case PpSelectionType.ppSelectionShapes:
                    PowerPoint.ShapeRange shapeRange = selection.ShapeRange;
                    int index = new DirectoryInfo(exportFolderPath).GetFiles().Length;
                    Parallel.ForEach(shapeRange.Cast<PowerPoint.Shape>(), shape =>
                    {
                        int i = Interlocked.Increment(ref index);
                        string shapeName = $"{presentationName}_{i}";
                        shape.Export(exportFolderPath + shapeName + ".emf", PpShapeFormat.ppShapeFormatEMF, 0, 0, PpExportMode.ppRelativeToSlide);
                    });
                    break;

                case PpSelectionType.ppSelectionSlides:
                    foreach (Slide slide in selection.SlideRange)
                        {
                        int j = 0;
                        for (int i = 1 ; i <= slide.Shapes.Count ; i++)
                            {
                            PowerPoint.Shape shape = slide.Shapes[i];
                            if (shape.Type == MsoShapeType.msoPicture)
                                {
                                try
                                    {
                                    j++;
                                    string shapeName = $"{presentationName}_{new DirectoryInfo(exportFolderPath).GetFiles().Length + j}";
                                    shape.Export(exportFolderPath + shapeName + ".emf", PpShapeFormat.ppShapeFormatEMF, 0, 0, PpExportMode.ppRelativeToSlide);
                                    }
                                catch
                                    {
                                    // Handle exceptions
                                    }
                                }
                            }
                        }
                    break;
                }

            // 打开导出文件所在的文件夹
            System.Diagnostics.Process.Start("explorer.exe", exportFolderPath);
            }

        /// <summary>
        /// 新增板式的函数
        /// </summary>
        /// <param name="type"></param>
        ///
        /// <param name="left"></param>
        /// <param name="top"></param>
        /// <param name="width"></param>
        /// <param name="height"></param>
        public void addPla(PpPlaceholderType type, float left, float top, float width, float height)
            {
            Slide slide = app.ActiveWindow.View.Slide;
            Master master = app.ActivePresentation.SlideMaster;

            // 获取当前幻灯片使用的自定义版式的 ID
            int layoutId = slide.CustomLayout.Index;

            // 在当前幻灯片的自定义版式中添加占位符
            PowerPoint.Shape shp = master.CustomLayouts[layoutId].Shapes.AddPlaceholder((PpPlaceholderType)type, left, top, width, height);

            // 给新添加的占位符添加标签
            shp.Tags.Add("母版", "占位符");

            // 设置新添加的占位符的字体大小
            float fsize = Settings.Default.Pla_size;
            shp.TextFrame.TextRange.Font.Size = fsize;

            // 刷新当前幻灯片的自定义版式
            DelePla(); // 删除所有非占位符的形状，以便重新排列占位符
            slide.CustomLayout = slide.CustomLayout;
            }

        /// <summary>
        /// 删除多余占位符的函数
        /// </summary>
        public void DelePla()
            {
            Slide slide = app.ActiveWindow.View.Slide;
            int count = slide.CustomLayout.Shapes.Count;
            for (int i = 0 ; i < count ; count--)
                {
                if (slide.CustomLayout.Shapes[count].Tags["母版"] != "占位符")
                    {
                    slide.CustomLayout.Shapes[count].Delete();
                    }
                }
            }

        private void splitButton3_Click(object sender, RibbonControlEventArgs e)
            {
            Selection sel = app.ActiveWindow.Selection;
            Slide slide = app.ActiveWindow.View.Slide;
            if ((Control.ModifierKeys & Keys.Control) == Keys.Control) //判断Ctrl键
                {
                Wpf_MasterSet wpf_MasterSet = new Wpf_MasterSet();
                wpf_MasterSet.ShowDialog();
                }
            else
                {
                if (sel.Type == PpSelectionType.ppSelectionShapes)
                    {
                    PowerPoint.ShapeRange range = sel.ShapeRange;
                    foreach (PowerPoint.Shape item in range)
                        {
                        PowerPoint.Shape shr = item;
                        float width = shr.Width;
                        float height = shr.Height;
                        float top = shr.Top;
                        float left = shr.Left;

                        int count = slide.CustomLayout.Index; //获取当前模板ID
                        if (shr.Type == MsoShapeType.msoTextBox) //文本
                            {
                            shr.PickUp(); //复制格式
                            if (Settings.Default.Pla_N3 == true)
                                {
                                PowerPoint.Shape shp = app.ActivePresentation.SlideMaster.CustomLayouts[count].Shapes.AddPlaceholder((PpPlaceholderType)2, left, top, width, height);
                                shp.Tags.Add("母版", "占位符");
                                shp.TextFrame.TextRange.Text = shr.TextFrame.TextRange.Text;
                                shp.TextFrame.TextRange.Font.Size = shr.TextFrame.TextRange.Font.Size;
                                shp.Apply(); //粘贴格式
                                shp.TextFrame.AutoSize = PpAutoSize.ppAutoSizeShapeToFitText;
                                DelePla();
                                slide.CustomLayout = slide.CustomLayout; //刷新当前板式
                                }
                            else
                                {
                                PowerPoint.Shape shp = app.ActivePresentation.SlideMaster.CustomLayouts[count].Shapes.AddPlaceholder((PpPlaceholderType)2, left, top, width, height);
                                shp.Tags.Add("母版", "占位符");
                                shp.Apply(); //粘贴格式
                                DelePla();
                                slide.CustomLayout = slide.CustomLayout; //刷新当前板式
                                }
                            }
                        else if (shr.Type == MsoShapeType.msoPicture) //图片
                            {
                            shr.PickUp(); //复制格式
                            if (Settings.Default.Pla_N4 == false)
                                {
                                PowerPoint.Shape shp = app.ActivePresentation.SlideMaster.CustomLayouts[count].Shapes.AddPlaceholder((PpPlaceholderType)18, left, top, width, height);
                                shp.Tags.Add("母版", "占位符");
                                shp.Apply(); //粘贴格式
                                DelePla();
                                slide.CustomLayout = slide.CustomLayout; //刷新当前板式
                                }
                            }
                        else if (shr.Type == MsoShapeType.msoSmartArt) //Smart图形
                            {
                            addPla(PpPlaceholderType.ppPlaceholderOrgChart, left, top, width, height);
                            }
                        else if (shr.Type == MsoShapeType.msoMedia) //媒体
                            {
                            addPla(PpPlaceholderType.ppPlaceholderMediaClip, left, top, width, height);
                            }
                        else if (shr.Type == MsoShapeType.msoChart) //图表
                            {
                            addPla(PpPlaceholderType.ppPlaceholderChart, left, top, width, height);
                            }
                        else if (shr.Type == MsoShapeType.msoTable) //表格
                            {
                            addPla(PpPlaceholderType.ppPlaceholderTable, left, top, width, height);
                            }
                        else if (shr.Type == MsoShapeType.msoAutoShape) //形状
                            {
                            if (Settings.Default.Pla_N2 == false)
                                {
                                addPla(PpPlaceholderType.ppPlaceholderObject, left, top, width, height);
                                }
                            else
                                {
                                return;
                                }
                            }
                        else
                            {
                            addPla(PpPlaceholderType.ppPlaceholderObject, left, top, width, height);
                            }
                        shr.ZOrder(MsoZOrderCmd.msoBringForward);
                        if (Settings.Default.Pla_N1 == false)
                            {
                            shr.Delete(); //删除源件
                            }
                        }
                    }
                else
                    {
                    Growl.Warning("请先选择内容对象！");
                    }
                }
            }

        private void button23_Click(object sender, RibbonControlEventArgs e)
            {
            Slides slides = Globals.ThisAddIn.Application.ActivePresentation.Slides;
            Slide slide = Globals.ThisAddIn.Application.ActiveWindow.View.Slide;
            if (Control.ModifierKeys.HasFlag(Keys.Control))
                {
                // 如果按下了 Ctrl 键，则清空当前幻灯片的所有形状和版式
                var app = Globals.ThisAddIn.Application;
                var layout = slide.CustomLayout;
                DeleteShapes(layout.Shapes);
                DeleteShapes(slide.Shapes);
                }
            else
                {
                // 如果没有按下 Ctrl 键，则在当前幻灯片后面添加一张空白幻灯片
                int count = slide.SlideIndex;
                CustomLayout customLayout = slides[slides.Count].CustomLayout;
                Slide newSlide = slides.AddSlide(count + 1, customLayout);
                newSlide.Layout = PpSlideLayout.ppLayoutBlank;
                }
            }

        /// <summary>
        /// 删除形状
        /// </summary>
        /// <param name="shapes"></param>
        private void DeleteShapes(PowerPoint.Shapes shapes)
            {
            foreach (PowerPoint.Shape shape in shapes.Cast<PowerPoint.Shape>().ToList())
                {
                shape.Delete();
                }
            }

        private void button24_Click(object sender, RibbonControlEventArgs e)
            {
            Presentation pre = app.ActivePresentation;
            Slide slide = app.ActiveWindow.View.Slide;

            for (int i = 1 ; i <= 36 ; i++)
                {
                int count = pre.Slides.Count + 1;
                pre.Slides.Add(count, (PpSlideLayout)i).Delete();//添加标题幻灯片
                }
            Growl.SuccessGlobal("版式补全成功");
            }

        private void button25_Click(object sender, RibbonControlEventArgs e)
            {
            PowerPoint.Master slideMaster = app.ActivePresentation.SlideMaster;
            PowerPoint.CustomLayouts customLayouts = slideMaster?.CustomLayouts;
            if (customLayouts != null)
                {
                int count = customLayouts.Count;
                int n = 0;
                for (int i = count ; i >= 1 ; i--)
                    {
                    PowerPoint.CustomLayout layout = slideMaster.CustomLayouts[i];
                    if (layout == null)
                        {
                        continue;
                        }
                    bool isUsed = false;
                    // 判断版式是否正在被使用
                    foreach (PowerPoint.Slide slide in app.ActivePresentation.Slides)
                        {
                        if (string.Equals(slide.CustomLayout.Name, layout.Name))
                            {
                            isUsed = true;
                            break;
                            }
                        }
                    if (!isUsed)
                        {
                        try
                            {
                            layout.Delete();
                            n++;
                            }
                        catch (Exception ex)
                            {
                            // 记录异常信息，这里可以选择将其输出到日志文件中
                            Console.WriteLine(ex.Message);
                            }
                        }
                    }
                Growl.Success("已删除 " + n + " 张未使用版式");
                }
            else
                {
                Growl.Warning("没有找到版式", "温馨提示");
                }
            }

        private void button21_Click(object sender, RibbonControlEventArgs e)
            {
            try
                {
                Slide slide = app.ActiveWindow.View.Slide;
                Presentation pre = app.ActivePresentation;
                int count = pre.SlideMaster.CustomLayouts.Count + 1;
                CustomLayout cus = pre.SlideMaster.CustomLayouts.Add(count);
                cus.Name = "PresPio-" + count;
                pre.SlideMaster.CustomLayouts[count].Shapes.Placeholders[1].Delete();
                slide.CustomLayout = pre.SlideMaster.CustomLayouts[count];//刷新当前板式
                }
            catch
                {
                return;
                }
            }

        private void button22_Click(object sender, RibbonControlEventArgs e)
            {
            try
                {
                PowerPoint.Selection sel = app.ActiveWindow.Selection;
                if (sel.Type == PpSelectionType.ppSelectionShapes || sel.Type == PpSelectionType.ppSelectionText)
                    {
                    PowerPoint.Slide slide = app.ActiveWindow.View.Slide;
                    PowerPoint.ShapeRange range = sel.ShapeRange;
                    if (sel.HasChildShapeRange)
                        {
                        range = sel.ChildShapeRange;
                        }
                    range.Copy();
                    slide.CustomLayout.Shapes.Paste();
                    }
                else
                    {
                    Growl.Warning("请选中形状");
                    }
                }
            catch
                {
                return;
                }
            }

        private void button26_Click(object sender, RibbonControlEventArgs e)
            {
            try
                {
                Slide slide = app.ActiveWindow.View.Slide;
                slide.CustomLayout = slide.CustomLayout;//刷新当前板式
                }
            catch
                {
                return;
                }
            }

        private void button27_Click(object sender, RibbonControlEventArgs e)
            {
            }

        private void button29_Click(object sender, RibbonControlEventArgs e)
            {
            saveFileDialog1.Filter = "PPT文件（*.pot)|*.pot";
            //设置默认文件类型显示顺序（可以不设置）
            saveFileDialog1.FilterIndex = 2;
            //保存对话框是否记忆上次打开的目录
            saveFileDialog1.RestoreDirectory = true;
            DialogResult dr = saveFileDialog1.ShowDialog();
            string fileName = saveFileDialog1.FileName;
            if (dr == System.Windows.Forms.DialogResult.OK && !string.IsNullOrEmpty(fileName))
                {
                app.Presentations[1].SaveAs(fileName, PpSaveAsFileType.ppSaveAsTemplate, MsoTriState.msoTriStateMixed);//导出PDF
                }
            else
                {
                Growl.Warning("您未选择文件夹，导出失败！");
                }
            }

        private void button28_Click(object sender, RibbonControlEventArgs e)
            {
            app.CommandBars.ExecuteMso("FileSaveAsPdfOrXps");//导出PDF
            }

        private void button30_Click(object sender, RibbonControlEventArgs e)
            {
            saveFileDialog1.Filter = "图片文件（*.bmp)|*.bmp";
            //设置默认文件类型显示顺序（可以不设置）
            saveFileDialog1.FilterIndex = 2;
            //保存对话框是否记忆上次打开的目录
            saveFileDialog1.RestoreDirectory = true;
            DialogResult dr = saveFileDialog1.ShowDialog();
            string fileName = saveFileDialog1.FileName;
            if (dr == System.Windows.Forms.DialogResult.OK && !string.IsNullOrEmpty(fileName))
                {
                app.Presentations[1].SaveAs(fileName, PpSaveAsFileType.ppSaveAsBMP, MsoTriState.msoTriStateMixed);//导出PDF
                }
            else
                {
                Growl.WarningGlobal("您未选择文件夹，导出失败！");
                }
            }

        private void button34_Click(object sender, RibbonControlEventArgs e)
            {
            app.CommandBars.ExecuteMso("FileSaveAsOtherFormats");
            }

        private void button31_Click(object sender, RibbonControlEventArgs e)
            {
            app.CommandBars.ExecuteMso("FileDocumentInspect");
            }

        private void button32_Click(object sender, RibbonControlEventArgs e)
            {
            app.CommandBars.ExecuteMso("FileCompatibilityCheckerPowerPoint");
            }

        private void button33_Click(object sender, RibbonControlEventArgs e)
            {
            app.CommandBars.ExecuteMso("CreateHandoutsInWord");
            }

        //新增触发器动画函数
        public void addTips(PowerPoint.Shape shp1, PowerPoint.Shape shp2, MsoAnimEffect type)
            {
            //shp1为触发器
            //shp2为动画窗格
            //动画类型 https://docs.microsoft.com/zh-cn/office/vba/api/powerpoint.msoanimeffect
            Slide slide = app.ActiveWindow.View.Slide;
            slide.TimeLine.InteractiveSequences.Add();
            Effect eff1 = slide.TimeLine.InteractiveSequences[1].AddTriggerEffect(shp2, type, MsoAnimTriggerType.msoAnimTriggerOnShapeClick, shp1);
            Effect eff2 = slide.TimeLine.InteractiveSequences[1].AddTriggerEffect(shp2, type, MsoAnimTriggerType.msoAnimTriggerOnShapeClick, shp1);
            eff2.Exit = MsoTriState.msoTrue;//退出效果
            }

        private void button35_Click(object sender, RibbonControlEventArgs e)
            {
            // 获取当前视图中的幻灯片
            Slide slide = app.ActiveWindow.View.Slide;
            // 创建交互序列并添加到幻灯片上
            TimeLine timeLine = slide.TimeLine;
            var interactiveSequence = timeLine.InteractiveSequences.Add();

            // 创建触发器形状
            PowerPoint.Shape shp1 = slide.Shapes.AddShape(MsoAutoShapeType.msoShapeOval, 45, 118, 30, 30);
            shp1.Name = "触发器";
            shp1.Fill.ForeColor.RGB = 11053224;  // 浅灰色10进制
            shp1.TextFrame2.TextRange.Text = "?";
            shp1.Line.Visible = MsoTriState.msoFalse;

            // 创建提示标签形状
            PowerPoint.Shape shp2 = slide.Shapes.AddShape(MsoAutoShapeType.msoShapeRoundedRectangularCallout, 0, 0, 200, 100);
            shp2.Name = "提示标签";
            shp2.Fill.ForeColor.RGB = 10134027;  // 填充颜色
            shp2.Line.ForeColor.RGB = 11053224;
            shp2.TextFrame2.TextRange.Text = "请输入备注内容";
            shp2.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = 16777215;
            shp2.TextFrame2.TextRange.Font.Size = 12f;

            // 将触发器形状和提示标签形状添加到交互序列并设置触发动画类型
            addTips(shp1, shp2, MsoAnimEffect.msoAnimEffectAppear);
            }

        private void button36_Click(object sender, RibbonControlEventArgs e)
            {
            Slide slide = app.ActiveWindow.View.Slide;
            Selection sel = app.ActiveWindow.Selection;
            if (sel.Type == PpSelectionType.ppSelectionShapes)
                {
                try
                    {
                    foreach (object shp in sel.ShapeRange)
                        {
                        PowerPoint.Shape shp1 = shp as PowerPoint.Shape;
                        float width = shp1.Width;
                        float height = shp1.Height;
                        float top = shp1.Top - height;
                        float left = shp1.Left;
                        PowerPoint.Shape shp2 = slide.Shapes.AddShape(MsoAutoShapeType.msoShapeRoundedRectangularCallout, left, top, width, height);
                        addTips(shp1, shp2, MsoAnimEffect.msoAnimEffectAppear);
                        }
                    }
                catch
                    {
                    }
                }
            }

        private void button37_Click(object sender, RibbonControlEventArgs e)
            {
            Slide slide = app.ActiveWindow.View.Slide;
            Selection sel = app.ActiveWindow.Selection;
            if (sel.Type == PpSelectionType.ppSelectionShapes)
                {
                try
                    {
                    foreach (object shp in sel.ShapeRange)
                        {
                        PowerPoint.Shape shp2 = shp as PowerPoint.Shape;
                        float width = shp2.Width;
                        float height = shp2.Height;
                        float top = shp2.Top;
                        float left = shp2.Left;
                        PowerPoint.Shape shp1 = slide.Shapes.AddShape(MsoAutoShapeType.msoShapeOval, left - 30, top, 30, 30);
                        shp1.Name = "触发器";//添加触发器
                        shp1.Fill.ForeColor.RGB = 11053224;//浅灰色10进制
                        shp1.TextFrame2.TextRange.Text = "?";
                        shp1.Line.Visible = MsoTriState.msoFalse;
                        addTips(shp1, shp2, MsoAnimEffect.msoAnimEffectAppear);
                        }
                    }
                catch
                    {
                    }
                }
            }

        private void button38_Click(object sender, RibbonControlEventArgs e)
            {
            Selection selection = this.app.ActiveWindow.Selection;
            if (selection.Type != PpSelectionType.ppSelectionShapes)
                {
                Growl.Warning("请选中一个对象");
                return;
                }

            PowerPoint.Shape shape = selection.ShapeRange[1];
            float width = shape.Width;
            float height = shape.Height;
            float aspectRatio = width / height;
            float inverseAspectRatio = height / width;
            string warningText = "ppt不支持该比例";

            // Ensure minimum dimensions
            if (width < 72f || height < 72f)
                {
                if (width < 72f) width = 72f;
                if (height < 72f) height = 72f;
                }

            // Cap dimensions
            if (width > 4032f || height > 4032f)
                {
                if (width > 4032f) width = 4032f;
                if (height > 4032f) height = 4032f;
                }

            // Adjust dimensions based on aspect ratio
            if (width < 72f || height < 72f || width > 4032f || height > 4032f)
                {
                Growl.Success(warningText);
                return;
                }

            // Convert to centimeters if Control key is pressed
            if ((System.Windows.Forms.Control.ModifierKeys & Keys.Control) == Keys.Control)
                {
                width = (float)Math.Round(width * 2.54f / 72f, 2);
                height = (float)Math.Round(height * 2.54f / 72f, 2);
                System.Windows.Clipboard.SetText($"宽：{width}， 高：{height}");
                Growl.Success("尺寸值(单位：厘米)已复制到剪贴板");
                return;
                }

            // Round and set dimensions
            width = (float)Math.Round(width, 2);
            height = (float)Math.Round(height, 2);
            this.app.ActivePresentation.PageSetup.SlideWidth = width;
            this.app.ActivePresentation.PageSetup.SlideHeight = height;
            shape.Width = width;
            shape.Height = height;
            shape.Left = 0f;
            shape.Top = 0f;
            }

        private void button39_Click(object sender, RibbonControlEventArgs e)
            {
            Presentation pre = app.ActivePresentation;
            float w = pre.PageSetup.SlideWidth;
            float h = pre.PageSetup.SlideHeight;
            Selection sel = app.ActiveWindow.Selection;
            if (sel.Type == PpSelectionType.ppSelectionShapes)
                {
                sel.ShapeRange.Height = h;
                sel.ShapeRange.Width = w;
                sel.ShapeRange.Top = 0;
                sel.ShapeRange.Left = 0;
                }
            else
                {
                Growl.Warning("请选择形状或图片再操作！");
                }
            }

        private void button40_Click(object sender, RibbonControlEventArgs e)
            {
            Presentation pre = app.ActivePresentation;
            try
                {
                pre.PageSetup.SlideWidth = 900f;
                pre.PageSetup.SlideHeight = 383f;
                }
            catch
                {
                return;
                }
            }

        private void button41_Click(object sender, RibbonControlEventArgs e)
            {
            Presentation pre = app.ActivePresentation;
            try
                {
                pre.PageSetup.SlideWidth = 200 * 2.834f;
                pre.PageSetup.SlideHeight = 200 * 2.834f;
                }
            catch
                {
                return;
                }
            }

        private void button42_Click(object sender, RibbonControlEventArgs e)
            {
            // 获取当前活动的演示文稿和幻灯片
            Presentation pre = app.ActivePresentation;
            Slide slide = app.ActiveWindow.View.Slide;

            // 设置幻灯片尺寸
            try
                {
                pre.PageSetup.SlideWidth = 340 * 2.834f;
                pre.PageSetup.SlideHeight = 102 * 2.834f;

                // 添加封面间隔线
                PowerPoint.Shape shp = slide.Shapes.AddLine(670, 0, 670, 288);
                shp.Name = "封面间隔线";
                shp.Line.ForeColor.RGB = 5526612;
                shp.Line.Weight = 2;

                // 添加首页封面文本框
                PowerPoint.Shape shp1 = slide.Shapes.AddTextbox(MsoTextOrientation.msoTextOrientationHorizontal, 0, 300, 280, 80);
                shp1.TextFrame.TextRange.Text = "首页封面";
                shp1.TextFrame.TextRange.Font.Size = 20;

                // 添加小图封面文本框
                PowerPoint.Shape shp2 = slide.Shapes.AddTextbox(MsoTextOrientation.msoTextOrientationHorizontal, 680, 300, 280, 80);
                shp2.TextFrame.TextRange.Text = "小图封面";
                shp2.TextFrame.TextRange.Font.Size = 20;
                }
            catch (Exception ex)
                {
                // 捕获并处理异常
                MessageBox.Show($"An error occurred: {ex.Message}", "Error", (MessageBoxButtons)MessageBoxButton.OK, (MessageBoxIcon)MessageBoxImage.Error);
                return;
                }
            }

        private void button43_Click(object sender, RibbonControlEventArgs e)
            {
            Presentation pre = app.ActivePresentation;
            try
                {
                pre.PageSetup.SlideWidth = 92 * 2.834f;
                pre.PageSetup.SlideHeight = 56 * 2.834f;
                }
            catch
                {
                return;
                }
            }

        private void button44_Click(object sender, RibbonControlEventArgs e)
            {
            Presentation pre = app.ActivePresentation;
            try
                {
                pre.PageSetup.SlideWidth = 800f;
                pre.PageSetup.SlideHeight = 800f;
                }
            catch
                {
                return;
                }
            }

        private void button45_Click(object sender, RibbonControlEventArgs e)
            {
            Presentation pre = app.ActivePresentation;
            try
                {
                pre.PageSetup.SlideWidth = 1920 / 2f;
                pre.PageSetup.SlideHeight = 1080 / 2f;
                }
            catch
                {
                return;
                }
            }

        private void button46_Click(object sender, RibbonControlEventArgs e)
            {
            Presentation pre = app.ActivePresentation;
            try
                {
                pre.PageSetup.SlideWidth = 750 * 2.834f;
                pre.PageSetup.SlideHeight = 280 * 2.834f;
                }
            catch
                {
                return;
                }
            }

        private void button47_Click(object sender, RibbonControlEventArgs e)
            {
            PowerPoint.Selection sel = app.ActiveWindow.Selection;
            if (sel.Type == PowerPoint.PpSelectionType.ppSelectionNone || sel.ShapeRange.Count <= 1)
                {
                Growl.Warning("请先选中至少两个元素，其中一个是形状，另一个是图片");
                }
            else
                {
                PowerPoint.Slide slide = app.ActiveWindow.View.Slide;
                string apath = app.ActivePresentation.Path;
                PowerPoint.ShapeRange range = sel.ShapeRange;
                List<int> list1 = new List<int>();
                List<int> list2 = new List<int>();
                for (int i = 1 ; i <= range.Count ; i++)
                    {
                    PowerPoint.Shape shape = range[i];
                    if (shape.Type == Office.MsoShapeType.msoAutoShape || shape.Type == Office.MsoShapeType.msoFreeform)
                        {
                        list1.Add(i);
                        }
                    else if (shape.Type == Office.MsoShapeType.msoPicture)
                        {
                        list2.Add(i);
                        }
                    }
                if (list2.Count == 0 && list1.Count != 0)
                    {
                    Growl.Warning("所选元素中没有图片或不能有组合");
                    }
                else if (list1.Count == 0)
                    {
                    Growl.Warning("所选元素中没有形状或不能有组合");
                    }
                else
                    {
                    float mw = 0;
                    for (int i = 0 ; i < list1.Count() ; i++)
                        {
                        int n = i % list2.Count();
                        PowerPoint.Shape pics = range[list2[n]];
                        float rt = pics.Rotation;
                        if (rt == 0)
                            {
                            pics.Export(apath + @"xshape.png", PowerPoint.PpShapeFormat.ppShapeFormatPNG);
                            PowerPoint.Shape shapes = range[list1[i]];
                            shapes.Fill.Visible = Office.MsoTriState.msoTrue;
                            shapes.Fill.UserPicture(apath + @"xshape.png");
                            if (!checkBox1.Checked)
                                {
                                mw = shapes.Left + shapes.Width / 2;
                                shapes.Width = pics.Width / pics.Height * shapes.Height;
                                shapes.Left = mw - shapes.Width / 2;
                                }
                            System.IO.File.Delete(apath + @"xshape.png");
                            }
                        else
                            {
                            pics.Copy();
                            PowerPoint.Shape npics = slide.Shapes.PasteSpecial(PowerPoint.PpPasteDataType.ppPastePNG)[1];
                            npics.ScaleHeight(1f, Office.MsoTriState.msoFalse, Office.MsoScaleFrom.msoScaleFromTopLeft);
                            npics.Export(apath + @"xshape.png", PowerPoint.PpShapeFormat.ppShapeFormatPNG);
                            PowerPoint.Shape shapes = range[list1[i]];
                            shapes.Fill.Visible = Office.MsoTriState.msoTrue;
                            shapes.Fill.UserPicture(apath + @"xshape.png");
                            if (!checkBox1.Checked)
                                {
                                mw = shapes.Left + shapes.Width / 2;
                                shapes.Width = npics.Width / npics.Height * shapes.Height;
                                shapes.Left = mw - shapes.Width / 2;
                                }
                            System.IO.File.Delete(apath + @"xshape.png");
                            npics.Delete();
                            }
                        }
                    }
                }
            }

        private void button48_Click(object sender, RibbonControlEventArgs e)
            {
            PowerPoint.Selection sel = app.ActiveWindow.Selection;
            if (sel.Type == PowerPoint.PpSelectionType.ppSelectionShapes)
                {
                System.Windows.Forms.OpenFileDialog openFileDialog1 = new System.Windows.Forms.OpenFileDialog();
                openFileDialog1.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
                openFileDialog1.Filter = "JPG图片|*.jpg|JPEG图片|*.jpeg|PNG图片|*.png|BMP图片|*.bmp|GIF图片|*.gif|EMF图片|*.emf|WMF图片|*.wmf";
                openFileDialog1.FilterIndex = 1;
                openFileDialog1.RestoreDirectory = true;
                openFileDialog1.AddExtension = true;
                openFileDialog1.Multiselect = true;
                if (openFileDialog1.ShowDialog() == DialogResult.OK)
                    {
                    string[] files = openFileDialog1.FileNames.ToArray();
                    PowerPoint.Slide slide = app.ActiveWindow.View.Slide;
                    PowerPoint.ShapeRange range = sel.ShapeRange;
                    if (sel.HasChildShapeRange)
                        {
                        range = sel.ChildShapeRange;
                        }
                    int count = Math.Min(range.Count, files.Count());
                    string[] oshapes = new string[count];

                    Bitmap obmp = null;
                    for (int i = 1 ; i <= count ; i++)
                        {
                        PowerPoint.Shape opic = sel.ShapeRange[i];
                        obmp = new Bitmap(files[i - 1]);
                        PowerPoint.Shape npic = null;
                        if (checkBox1.Checked)
                            {
                            npic = slide.Shapes.AddPicture(files[i - 1], Office.MsoTriState.msoFalse, Office.MsoTriState.msoTrue, opic.Left, opic.Top, opic.Width, opic.Height);
                            }
                        else
                            {
                            float whn = (float)obmp.Width / (float)obmp.Height * opic.Height;
                            npic = slide.Shapes.AddPicture(files[i - 1], Office.MsoTriState.msoFalse, Office.MsoTriState.msoTrue, opic.Left + opic.Width / 2 - whn / 2, opic.Top, whn, opic.Height);
                            }
                        if (checkBox1.Checked)
                            {
                            try
                                {
                                opic.PickUp();
                                npic.Apply();
                                }
                            catch { }
                            try
                                {
                                opic.PickupAnimation();
                                npic.ApplyAnimation();
                                }
                            catch { }
                            }
                        npic.Rotation = opic.Rotation;
                        oshapes[i - 1] = opic.Name;
                        }
                    obmp.Dispose();
                    slide.Shapes.Range(oshapes).Delete();
                    }
                }
            else
                {
                Growl.Warning("请先选中要替换的图片");
                }
            }

        private void button49_Click(object sender, RibbonControlEventArgs e)
            {
            PowerPoint.Selection sel = app.ActiveWindow.Selection;
            if (sel.Type != PowerPoint.PpSelectionType.ppSelectionShapes || sel.ShapeRange.Count < 2)
                {
                Growl.Warning("请先选中至少两个元素，先选中原形状，最后选中新形状");
                }
            else
                {
                PowerPoint.Slide slide = app.ActiveWindow.View.Slide;
                PowerPoint.ShapeRange range = sel.ShapeRange;
                PowerPoint.Shape nshape1 = range[range.Count];
                if (nshape1.LockAspectRatio == Office.MsoTriState.msoFalse)
                    {
                    nshape1.LockAspectRatio = Office.MsoTriState.msoTrue;
                    }
                List<string> name = new List<string>();
                for (int i = 1 ; i < range.Count ; i++)
                    {
                    range[i].Name = range[i].Name + "_" + i;
                    name.Add(range[i].Name);
                    }
                for (int i = 1 ; i < range.Count ; i++)
                    {
                    PowerPoint.Shape nshape = nshape1.Duplicate()[1];
                    PowerPoint.Shape oshape = slide.Shapes[name[i - 1]];
                    if (checkBox1.Checked)
                        {
                        if (nshape.LockAspectRatio == Office.MsoTriState.msoTrue)
                            {
                            nshape.LockAspectRatio = Office.MsoTriState.msoFalse;
                            }
                        nshape.Width = oshape.Width;
                        nshape.Height = oshape.Height;
                        nshape.Top = oshape.Top;
                        nshape.Left = oshape.Left;
                        }
                    else
                        {
                        nshape.Height = oshape.Height;
                        nshape.Left = oshape.Left + oshape.Width / 2 - nshape.Width / 2;
                        nshape.Top = oshape.Top;
                        }
                    nshape.Rotation = oshape.Rotation;
                    if (checkBox1.Checked)
                        {
                        try
                            {
                            oshape.PickUp();
                            nshape.Apply();
                            }
                        catch { }
                        try
                            {
                            oshape.PickupAnimation();
                            nshape.ApplyAnimation();
                            }
                        catch { }
                        }
                    }
                slide.Shapes.Range(name.ToArray()).Delete();
                }
            }

        private void checkBox1_Click(object sender, RibbonControlEventArgs e)
            {
            Properties.Settings.Default.Replace = checkBox1.Checked;
            Properties.Settings.Default.Save();
            }

        private void toggleButton1_Click(object sender, RibbonControlEventArgs e)
            {
            }

        private void toggleButton3_Click(object sender, RibbonControlEventArgs e)
            {
            }

        //删除配色色块
        /// <summary>
        /// 输入图形的名称 string Name
        /// </summary>
        /// <param name="Name">名称</param>
        public void DelShpe(string Name)
            {
            Slide slide = app.ActiveWindow.View.Slide;
            int Count = slide.Shapes.Count;
            for (int i = 0 ; i < Count ; Count--)
                {
                PowerPoint.Shape shape = slide.Shapes[Count];
                if (shape.Tags["配色"] == Name)
                    {
                    shape.Delete();
                    }
                }
            }

        private void button54_Click(object sender, RibbonControlEventArgs e)
            {
            Selection sel = app.ActiveWindow.Selection;
            DelShpe("色块");
            PowerPoint.ShapeRange shp0 = sel.ShapeRange;

            shp0.Copy(); // 复制
            float toTop = shp0.Top;
            int oldColor = sel.ShapeRange.Fill.ForeColor.RGB;
            MyFunction F = new MyFunction();
            Color newColor = F.Int2RGB(oldColor);
            var shapes = new PowerPoint.ShapeRange[6];
            var random = new Random();

            for (int i = 1 ; i < shapes.Length ; i++)
                {
                shapes[i] = shp0.Duplicate();
                shapes[i].Tags.Add("配色", "色块");
                shapes[i].Top = toTop;
                shapes[i].Left = shp0.Left + shp0.Width * i;

                int r = random.Next(newColor.R);
                int g = random.Next(newColor.G);
                int b = random.Next(newColor.B);
                b = (r + g > 400) ? r + g - 400 : b;
                b = (b > 255) ? 255 : b;
                shapes[i].Fill.ForeColor.RGB = F.RGB2Int(r, g, b);
                shapes[i].TextFrame.TextRange.Text = F.RGB2Int(r, g, b).ToString();
                }
            }

        #region

        private void button55_Click(object sender, RibbonControlEventArgs e)
            {
            }

        #endregion

        private void button52_Click(object sender, RibbonControlEventArgs e)
            {
            }

        private void button51_Click(object sender, RibbonControlEventArgs e)
            {
            }

        private void button60_Click(object sender, RibbonControlEventArgs e)
            {
            try
                {
                Selection sel = app.ActiveWindow.Selection;
                Slide slide = app.ActiveWindow.View.Slide;

                if (sel.Type == PpSelectionType.ppSelectionShapes)
                    {
                    PowerPoint.ShapeRange shr = sel.ShapeRange;
                    string oPath_Dir = Path.GetTempPath();
                    int count = shr.Count;

                    for (int i = count ; i >= 1 ; i--)
                        {
                        string oPath_Full = Path.Combine(oPath_Dir, "temp" + i + ".png");
                        shr[i].Export(oPath_Full, PpShapeFormat.ppShapeFormatPNG, 0, 0);

                        PowerPoint.Shape shrNew = slide.Shapes.AddShape(MsoAutoShapeType.msoShapeRectangle, shr[i].Left, shr[i].Top, shr[i].Width, shr[i].Height);
                        shrNew.Fill.Visible = MsoTriState.msoTrue;
                        shrNew.Fill.Solid();
                        shrNew.Line.Visible = MsoTriState.msoFalse;
                        shrNew.Fill.UserPicture(oPath_Full);

                        shr[i].Delete();
                        File.Delete(oPath_Full); // 删除临时文件
                        }
                    }
                Growl.SuccessGlobal("转换成功！");
                }
            catch (Exception ex)
                {
                Console.WriteLine("Error processing shapes: " + ex.Message);
                }
            }

        private void button57_Click(object sender, RibbonControlEventArgs e)
            {
            YouDao youdao = new YouDao();
            Selection sel = app.ActiveWindow.Selection;
            PowerPoint.Shape oShp;
            string okeyWord;
            if (sel.Type == PpSelectionType.ppSelectionNone)
                {
                return;
                }
            else if (sel.Type == PpSelectionType.ppSelectionShapes)
                {
                okeyWord = sel.ShapeRange.TextFrame.TextRange.Text;
                oShp = sel.ShapeRange[1];
                PowerPoint.ShapeRange shr = oShp.Duplicate();
                shr.Top = oShp.Top + oShp.Height;
                shr.Left = oShp.Left;
                shr.TextFrame.TextRange.Font.Size = oShp.TextFrame.TextRange.Font.Size / 5 * 3;
                shr.TextFrame.TextRange.Text = youdao.YouDaos(okeyWord, "auto", "en");
                shr.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = 12632256;
                shr.TextFrame.TextRange.ChangeCase(PpChangeCase.ppCaseUpper);//英文大写
                shr.Name = "英文修饰";
                }
            else if (sel.Type == PpSelectionType.ppSelectionShapes)
                {
                foreach (PowerPoint.ShapeRange shr in sel.ShapeRange)
                    {
                    if (shr.HasTextFrame == MsoTriState.msoCTrue)
                        {
                        if (shr.TextFrame.HasText == MsoTriState.msoCTrue)
                            {
                            okeyWord = sel.ShapeRange.TextFrame.TextRange.Text;
                            PowerPoint.ShapeRange shr1 = shr.Duplicate();
                            shr1.Top = shr.Top + shr.Height;
                            shr1.Left = shr.Left;
                            shr1.TextFrame.TextRange.Font.Size = shr.TextFrame.TextRange.Font.Size / 5 * 3;
                            shr1.TextFrame.TextRange.Text = youdao.YouDaos(okeyWord, "auto", "en");
                            shr.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = 12632256;
                            shr.TextFrame.TextRange.ChangeCase(PpChangeCase.ppCaseUpper);//英文大写
                            }
                        }
                    }
                }
            }

        private void button58_Click(object sender, RibbonControlEventArgs e)
            {
            Selection sel = app.ActiveWindow.Selection;
            PowerPoint.Shape oShp;
            string okeyWord;
            if (sel.Type == PpSelectionType.ppSelectionNone)
                {
                return;
                }
            else if (sel.Type == PpSelectionType.ppSelectionShapes)
                {
                okeyWord = sel.ShapeRange.TextFrame.TextRange.Text;
                string Text3 = NPinyin.Pinyin.GetPinyin(okeyWord);
                oShp = sel.ShapeRange[1];
                PowerPoint.ShapeRange shr = oShp.Duplicate();
                shr.Top = oShp.Top - oShp.Height / 2;
                shr.Left = oShp.Left;
                shr.TextFrame.TextRange.Font.Size = oShp.TextFrame.TextRange.Font.Size / 5 * 3;
                shr.TextFrame.TextRange.Text = Text3;
                shr.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = 12632256;
                shr.TextFrame.TextRange.ChangeCase(PpChangeCase.ppCaseUpper);//英文大写
                shr.Name = "拼音修饰";
                }
            else if (sel.Type == PpSelectionType.ppSelectionShapes)
                {
                foreach (PowerPoint.ShapeRange shr in sel.ShapeRange)
                    {
                    if (shr.HasTextFrame == MsoTriState.msoCTrue)
                        {
                        if (shr.TextFrame.HasText == MsoTriState.msoCTrue)
                            {
                            okeyWord = sel.ShapeRange.TextFrame.TextRange.Text;
                            string Text3 = NPinyin.Pinyin.GetPinyin(okeyWord);
                            PowerPoint.ShapeRange shr1 = shr.Duplicate();
                            shr1.Top = shr.Top - shr.Height / 2;
                            shr1.Left = shr.Left;
                            shr1.TextFrame.TextRange.Font.Size = shr.TextFrame.TextRange.Font.Size / 5 * 3;
                            shr1.TextFrame.TextRange.Text = Text3;
                            shr.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = 12632256;
                            shr.TextFrame.TextRange.ChangeCase(PpChangeCase.ppCaseUpper);//英文大写
                            shr.Name = "拼音修饰";
                            }
                        }
                    }
                }
            }

        private void button69_Click(object sender, RibbonControlEventArgs e)
            {
            // 创建 Wpf_superGuide 对象实例
            Wpf_superGuide wpf_SuperGuide = new Wpf_superGuide();
            // 显示窗口
            wpf_SuperGuide.Show();
            }

        private void button70_Click(object sender, RibbonControlEventArgs e)
            {
            PowerPoint.Selection sel = app.ActiveWindow.Selection;
            PowerPoint.Slide myDocument = app.ActiveWindow.View.Slide;

            if (sel.Type == PpSelectionType.ppSelectionSlides)
                {
                foreach (PowerPoint.Slide item in sel.SlideRange)
                    {
                    if (item.SlideShowTransition.Hidden == MsoTriState.msoTrue)
                        {
                        item.SlideShowTransition.Hidden = MsoTriState.msoFalse;
                        }

                    foreach (PowerPoint.Shape shp in item.Shapes)
                        {
                        shp.Visible = MsoTriState.msoTrue;
                        }
                    }
                }
            else
                {
                for (int i = myDocument.Shapes.Count ; i > 0 ; i--)
                    {
                    PowerPoint.Shape shape = myDocument.Shapes[i];

                    if (shape.Visible == MsoTriState.msoFalse)
                        {
                        shape.Visible = MsoTriState.msoTrue;
                        }
                    }
                }
            }

        private void splitButton5_Click(object sender, RibbonControlEventArgs e)
            {
            PowerPoint.Selection sel = app.ActiveWindow.Selection;

            if (sel.Type == PpSelectionType.ppSelectionSlides)
                {
                sel.SlideRange.SlideShowTransition.Hidden = MsoTriState.msoTrue;
                }
            else if (sel.Type == PowerPoint.PpSelectionType.ppSelectionShapes)
                {
                PowerPoint.ShapeRange range = sel.ShapeRange;
                if (sel.HasChildShapeRange)
                    {
                    range = sel.ChildShapeRange;
                    }
                foreach (PowerPoint.Shape shape in range)
                    {
                    if (shape.Visible == Office.MsoTriState.msoTrue)
                        {
                        shape.Visible = Office.MsoTriState.msoFalse;
                        }
                    }
                }
            else
                {
                Growl.WarningGlobal("请先选择幻灯片或元素");
                }
            }

        private void button53_Click(object sender, RibbonControlEventArgs e)
            {
            }

        private void comboBox1_TextChanged(object sender, RibbonControlEventArgs e)
            {
            }

        private void button71_Click(object sender, RibbonControlEventArgs e)
            {
            //Form_iTheme iTheme = null;
            //if (iTheme == null || iTheme.IsDisposed)
            //    {
            //    iTheme = new Form_iTheme();
            //    IntPtr handle = Process.GetCurrentProcess().MainWindowHandle;
            //    NativeWindow win = NativeWindow.FromHandle(handle);
            //    iTheme.ShowDialog();

            //    }
            }

        private void splitButton4_Click(object sender, RibbonControlEventArgs e)
            {
            Wpf_crossPage wpf_CrossPage = new Wpf_crossPage();
            wpf_CrossPage.Show();
            }

        private void button72_Click(object sender, RibbonControlEventArgs e)
            {
            Selection sel = app.ActiveWindow.Selection;
            Slide slide = app.ActiveWindow.View.Slide;
            Presentation pre = app.ActivePresentation;
            float width = pre.PageSetup.SlideWidth;//获取幻灯片长宽
            float height = pre.PageSetup.SlideHeight;

            if (sel.Type == PpSelectionType.ppSelectionSlides)
                {
                string sName = System.IO.Path.GetFileNameWithoutExtension(pre.Name);//获取无后缀名称
                int count = sel.SlideRange.Count;
                Stopwatch sw = new Stopwatch();
                sw.Start();
                for (int i = 1 ; i <= count ; i++)
                    {
                    Slide sl = sel.SlideRange[i];//获取每一页PPT
                    string pathName = Path.Combine(Path.GetTempPath(), $"{sName}-{i}.png");
                    sl.Export(pathName, "PNG", (int)width, (int)height);
                    sl.Shapes.AddPicture(pathName, MsoTriState.msoTrue, MsoTriState.msoTrue, 0, 0, width, height).Name = "覆盖封面";
                    File.Delete(pathName);//删除临时文件
                    }
                sw.Stop();
                MessageBox.Show($"共处理{count}页，总耗时{sw.ElapsedMilliseconds}毫秒。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            else
                {
                MessageBox.Show("请选择幻灯片！", "温馨提示");
                }
            }

        private void button73_Click(object sender, RibbonControlEventArgs e)
            {
            Presentation pre = app.ActivePresentation;
            Selection sel = app.ActiveWindow.Selection;
            if (sel.Type == PpSelectionType.ppSelectionShapes)
                {
                if (sel.ShapeRange.HasTextFrame == MsoTriState.msoTrue)
                    {
                    PowerPoint.ShapeRange shr = sel.ShapeRange;
                    int count = shr.Count;
                    for (int i = 1 ; i <= count ; count--)
                        {
                        PowerPoint.Shape shape = shr[count];
                        float sTop = shape.Top;
                        float sLeft = shape.Left;
                        PowerPoint.ShapeRange shr1 = shape.Duplicate();
                        shr1.Fill.Visible = MsoTriState.msoFalse;
                        shr1.Line.Visible = MsoTriState.msoFalse;
                        shr1.Top = shape.Top;
                        shr1.Left = shape.Left;
                        shape.TextFrame.TextRange.Text = "";
                        }
                    }
                else
                    {
                    Growl.Warning("请选择带有文本的形状", "温馨提示");
                    }
                }
            }

        private void button74_Click(object sender, RibbonControlEventArgs e)
            {
            PowerPoint.Selection sel = app.ActiveWindow.Selection;
            string apath = Properties.Settings.Default.Temp_Path;
            if (sel.Type != PowerPoint.PpSelectionType.ppSelectionShapes || sel.ShapeRange.Count < 2)
                {
                Growl.Warning("至少选中2个图形，并将矢量形状置于要裁剪的图片之上，形状不要旋转", "温馨提示");
                }
            else
                {
                PowerPoint.Slide slide = app.ActiveWindow.View.Slide;//定义slide
                PowerPoint.ShapeRange range = sel.ShapeRange; //定义shaperange
                int count = range.Count;
                List<string> oname = new List<string>();
                List<string> name = new List<string>();
                for (int i = 2 ; i <= range.Count ; i++)
                    {
                    oname.Add(range[i].Name);
                    range[i].Name = range[i].Name + "_" + i;
                    name.Add(range[i].Name);
                    }
                PowerPoint.Shape pic = range[1];
                pic.Export(apath + @"xshape.png", PowerPoint.PpShapeFormat.ppShapeFormatPNG);
                for (int i = 2 ; i <= count ; i++)
                    {
                    PowerPoint.Shape shape = slide.Shapes[name[i - 2]];
                    if (shape.Rotation == 0)
                        {
                        shape.Fill.Visible = MsoTriState.msoTrue;
                        shape.Fill.UserPicture(apath + @"xshape.png");
                        shape.Fill.TextureTile = MsoTriState.msoFalse;
                        shape.Fill.RotateWithObject = MsoTriState.msoTrue;
                        shape.PictureFormat.Crop.PictureOffsetX = pic.Left + pic.Width / 2 - shape.Left - shape.PictureFormat.Crop.PictureWidth / 2;
                        shape.PictureFormat.Crop.PictureOffsetY = pic.Top + pic.Height / 2 - shape.Top - shape.PictureFormat.Crop.PictureHeight / 2;
                        shape.PictureFormat.Crop.PictureHeight = pic.Height;
                        shape.PictureFormat.Crop.PictureWidth = pic.Width;
                        }
                    else
                        {
                        shape.Copy();
                        PowerPoint.Shape nshape = slide.Shapes.PasteSpecial(PowerPoint.PpPasteDataType.ppPastePNG)[1];
                        nshape.Left = shape.Left + shape.Width / 2 - nshape.Width / 2;
                        nshape.Top = shape.Top + shape.Height / 2 - nshape.Height / 2;
                        shape.Fill.Visible = Office.MsoTriState.msoTrue;
                        shape.Fill.UserPicture(apath + @"xshape.png");
                        shape.Fill.TextureTile = Office.MsoTriState.msoFalse;
                        shape.Fill.RotateWithObject = Office.MsoTriState.msoFalse;
                        shape.PictureFormat.Crop.PictureOffsetX = pic.Left + pic.Width / 2 - nshape.Left - shape.PictureFormat.Crop.PictureWidth / 2;
                        shape.PictureFormat.Crop.PictureOffsetY = pic.Top + pic.Height / 2 - nshape.Top - shape.PictureFormat.Crop.PictureHeight / 2;
                        shape.PictureFormat.Crop.PictureHeight = pic.Height;
                        shape.PictureFormat.Crop.PictureWidth = pic.Width;
                        nshape.Delete();
                        }
                    }
                System.IO.File.Delete(apath + @"xshape.png");
                for (int i = 0 ; i < oname.Count() ; i++)
                    {
                    slide.Shapes[name[i]].Name = oname[i];
                    }
                }
            }

        private void button90_Click(object sender, RibbonControlEventArgs e)
            {
            }

        private void splitButton6_Click(object sender, RibbonControlEventArgs e)
            {
            MsoAutoShapeType msoAutoShapeType = Properties.Settings.Default.Shape_Style;
            Color shpColor = Properties.Settings.Default.Shape_Color;
            Selection sel = app.ActiveWindow.Selection;
            Slide slide = app.ActiveWindow.View.Slide;
            float transparency = Properties.Settings.Default.Shape_Tra / 100;

            // 判断是否按下 Ctrl 键
            bool isCtrlPressed = (System.Windows.Forms.Control.ModifierKeys & Keys.Control) == Keys.Control;

            // 获取幻灯片宽度和高度
            float slideWidth = app.ActivePresentation.PageSetup.SlideWidth;
            float slideHeight = app.ActivePresentation.PageSetup.SlideHeight;

            // 添加遮罩或衬底的函数
            void AddShapeToSlide(string prefix)
                {
                if (sel.Type != PpSelectionType.ppSelectionNone)
                    {
                    int count = sel.ShapeRange.Count;
                    for (int i = 0 ; i < count ; i++)
                        {
                        PowerPoint.Shape shp = sel.ShapeRange[count - i]; // 从后向前选择
                        float left = shp.Left;
                        float top = shp.Top;
                        float width = shp.Width;
                        float height = shp.Height;

                        PowerPoint.Shape newShp = slide.Shapes.AddShape(msoAutoShapeType, left, top, width, height);
                        newShp.Name = $"{prefix}_{i}";
                        newShp.Line.Visible = MsoTriState.msoFalse;

                        if (isCtrlPressed)
                            {
                            MyFunction F = new MyFunction();
                            newShp.Fill.ForeColor.RGB = F.RGB2Int(shpColor.R, shpColor.G, shpColor.B);
                            newShp.Fill.Transparency = transparency;
                            newShp.ZOrder(MsoZOrderCmd.msoBringForward);
                            }
                        else
                            {
                            newShp.ZOrder(MsoZOrderCmd.msoSendToBack);
                            }
                        }
                    }
                else
                    {
                    PowerPoint.Shape shape = slide.Shapes.AddShape(msoAutoShapeType, 0, 0, slideWidth, slideHeight);
                    shape.Line.Visible = MsoTriState.msoFalse;
                    shape.Line.Weight = 0;
                    shape.Select();

                    if (isCtrlPressed)
                        {
                        MyFunction F = new MyFunction();
                        shape.Fill.ForeColor.RGB = F.RGB2Int(shpColor.R, shpColor.G, shpColor.B);
                        shape.Fill.Transparency = transparency;
                        }
                    }
                }

            // 添加遮罩或衬底
            AddShapeToSlide(isCtrlPressed ? "遮罩" : "衬底");
            }

        /// <summary>
        /// 添加形状
        /// </summary>
        /// <param name="shapeRange"></param>
        private void AddRectangleShape(PowerPoint.ShapeRange shapeRange)
            {
            PowerPoint.Slide slide = app.ActiveWindow.View.Slide;
            float w = app.ActivePresentation.PageSetup.SlideWidth;
            float h = app.ActivePresentation.PageSetup.SlideHeight;
            PowerPoint.Shape shape = slide.Shapes.AddShape(Office.MsoAutoShapeType.msoShapeRectangle, 0, 0, w, h);
            if (shape.Line.Visible == Office.MsoTriState.msoTrue)
                {
                shape.Line.Visible = Office.MsoTriState.msoFalse;
                }
            shape.Line.Weight = 0;
            shape.Select();
            }

        private void button75_Click(object sender, RibbonControlEventArgs e)
            {
            PowerPoint.Application app = Globals.ThisAddIn.Application;
            PowerPoint.Selection sel = app.ActiveWindow.Selection;
            PowerPoint.Slide slide = sel.SlideRange[1];

            float slideWidth = slide.Master.Width; // 获取画布宽度
            float slideHeight = slide.Master.Height; // 获取画布高度

            // 在当前页插入矩形形状
            PowerPoint.Shape rectangle = slide.Shapes.AddShape(
                Microsoft.Office.Core.MsoAutoShapeType.msoShapeRectangle,
                0, 0, slideWidth, slideHeight);

            // 可以设置矩形的属性，如填充颜色、边框颜色等
            //rectangle.Fill.ForeColor.RGB = System.Drawing.Color.FromArgb(255, 255, 255).ToArgb();
            //rectangle.Line.ForeColor.RGB = System.Drawing.Color.FromArgb(0, 0, 0).ToArgb();
            }

        private void button81_Click(object sender, RibbonControlEventArgs e)
            {
            DeleCG();
            Slide slide = app.ActiveWindow.View.Slide;
            float PageWidth = app.ActivePresentation.PageSetup.SlideWidth;
            float PageHeigh = app.ActivePresentation.PageSetup.SlideHeight;
            PowerPoint.Shape shp1 = slide.Shapes.AddShape(Office.MsoAutoShapeType.msoShapeRoundedRectangle, 0, PageHeigh * 2, 30, 30);
            PowerPoint.Shape shp2 = slide.Shapes.AddShape(Office.MsoAutoShapeType.msoShapeRoundedRectangle, 0, -PageHeigh, 30, 30);
            shp1.Line.Visible = MsoTriState.msoFalse;
            shp2.Line.Visible = MsoTriState.msoFalse;
            shp1.Tags.Add("撑高", "Targeting");
            shp2.Tags.Add("撑高", "Targeting");
            Growl.SuccessGlobal("定位设置成功！");
            }

        //删除页面撑高色块
        public void DeleCG()
            {
            Slide slide = app.ActiveWindow.View.Slide;
            int num = slide.Shapes.Count;
            for (int i = 0 ; i < num ; num--)
                {
                PowerPoint.Shape shp = slide.Shapes[num];
                if (shp.Tags["撑高"] == "Targeting")
                    {
                    shp.Delete();
                    }
                else
                    {
                    return;
                    }
                }
            }

        private void button5_Click_1(object sender, RibbonControlEventArgs e)
            {
            DeleCG();
            Growl.SuccessGlobal("取消定位成功！");
            }

        private void button92_Click(object sender, RibbonControlEventArgs e)
            {
            PresPio.Public_Wpf.Wpf_UnitName wpf_UnitName = new PresPio.Public_Wpf.Wpf_UnitName();
            wpf_UnitName.Show();
            }

        //底层函数
        public class MyFunction
            {
            //颜色转换函数
            public void RGB2HSL()
                {
                }

            public int RGB2Int(int R, int G, int B)
                {
                int PPTRGB = R + G * 256 + B * 256 * 256;
                return PPTRGB;
                }

            public Color Int2RGB(int color)
                {
                int B = color / (256 * 256);
                int G = (color - B * 256 * 256) / 256;
                int R = color - B * 256 * 256 - G * 256;
                return Color.FromArgb(R, G, B);
                }

            public int Rgb2Hsl(int r, int g, int b)
                {
                float num = 0f;
                float num2 = (float)Math.Max(Math.Max(r, g), b);
                float num3 = (float)Math.Min(Math.Min(r, g), b);
                if (num2 == num3)
                    {
                    num = 0f;
                    }
                else
                    {
                    if (num2 == (float)r)
                        {
                        if (g >= b)
                            {
                            num = (float)(42 * (g - b)) / (num2 - num3) + 0f;
                            }
                        else
                            {
                            num = (float)(42 * (g - b)) / (num2 - num3) + 255f;
                            }
                        }
                    if ((num2 == (float)g) & (num2 != (float)r))
                        {
                        num = (float)(42 * (b - r)) / (num2 - num3) + 85f;
                        }
                    if (num2 == (float)b && num2 != (float)g)
                        {
                        num = (float)(42 * (r - g)) / (num2 - num3) + 170f;
                        }
                    }
                if (num >= (float)((int)num) + 0.5f)
                    {
                    num = (float)((int)num + 1);
                    }
                float num4 = (num2 + num3) / 2f;
                if (num2 + num3 == 255f)
                    {
                    num4 = 128f;
                    }
                if (num4 >= (float)((int)num4) + 0.5f)
                    {
                    num4 = (float)((int)num4 + 1);
                    }
                float num5;
                if (num4 == 0f || num2 == num3)
                    {
                    num5 = 0f;
                    }
                else if (num4 <= 127f)
                    {
                    num5 = 255f * (num2 - num3) / (num2 + num3);
                    }
                else
                    {
                    num5 = 255f * (num2 - num3) / (510f - (num2 + num3));
                    }
                if (num5 >= (float)((int)num5) + 0.5f)
                    {
                    num5 = (float)((int)num5 + 1);
                    }
                return (int)num + (int)num5 * 256 + (int)num4 * 256 * 256;
                }

            public int Hsl2Rgb(int h, int s, int l)
                {
                float num = (float)h / 255f;
                float num2 = (float)s / 255f;
                float num3 = (float)l / 255f;
                float num4;
                float num5;
                float num6;
                if (num2 == 0f)
                    {
                    num4 = num3;
                    num5 = num3;
                    num6 = num3;
                    }
                else
                    {
                    float num7;
                    if (num3 < 0.5f)
                        {
                        num7 = num3 * (1f + num2);
                        }
                    else
                        {
                        num7 = num3 + num2 - num3 * num2;
                        }
                    float num8 = 2f * num3 - num7;
                    float num9 = num + 0.33333334f;
                    float num10 = num;
                    float num11 = num - 0.33333334f;
                    if (num9 < 0f)
                        {
                        num9 += 1f;
                        }
                    else if (num9 > 1f)
                        {
                        num9 -= 1f;
                        }
                    else
                        {
                        num9 += 0f;
                        }
                    if (num10 < 0f)
                        {
                        num10 += 1f;
                        }
                    else if (num10 > 1f)
                        {
                        num10 -= 1f;
                        }
                    else
                        {
                        num10 += 0f;
                        }
                    if (num11 < 0f)
                        {
                        num11 += 1f;
                        }
                    else if (num11 > 1f)
                        {
                        num11 -= 1f;
                        }
                    else
                        {
                        num11 += 0f;
                        }
                    if (num9 < 0.16666667f)
                        {
                        num4 = num8 + (num7 - num8) * 6f * num9;
                        }
                    else if (num9 < 0.5f)
                        {
                        num4 = num7;
                        }
                    else if (num9 < 0.6666667f)
                        {
                        num4 = num8 + (num7 - num8) * 6f * (0.6666667f - num9);
                        }
                    else
                        {
                        num4 = num8;
                        }
                    if (num10 < 0.16666667f)
                        {
                        num5 = num8 + (num7 - num8) * 6f * num10;
                        }
                    else if (num10 < 0.5f)
                        {
                        num5 = num7;
                        }
                    else if (num10 < 0.6666667f)
                        {
                        num5 = num8 + (num7 - num8) * 6f * (0.6666667f - num10);
                        }
                    else
                        {
                        num5 = num8;
                        }
                    if (num11 < 0.16666667f)
                        {
                        num6 = num8 + (num7 - num8) * 6f * num11;
                        }
                    else if (num11 < 0.5f)
                        {
                        num6 = num7;
                        }
                    else if (num11 < 0.6666667f)
                        {
                        num6 = num8 + (num7 - num8) * 6f * (0.6666667f - num11);
                        }
                    else
                        {
                        num6 = num8;
                        }
                    }
                int num12 = (int)(num4 * 255f);
                int num13 = (int)(num5 * 255f);
                int num14 = (int)(num6 * 255f);
                return num12 + num13 * 256 + num14 * 256 * 256;
                }
            }

        private void button80_Click(object sender, RibbonControlEventArgs e)
            {
            Selection sel = this.app.ActiveWindow.Selection;
            Slide slide = app.ActiveWindow.View.Slide;

            if (sel.Type != PpSelectionType.ppSelectionShapes)
                {
                Growl.WarningGlobal("请选择组合后再试！");
                return;
                }
            List<string> shps = new List<string>();
            if (sel.HasChildShapeRange)
                {
                List<PowerPoint.Shape> gshape = new List<PowerPoint.Shape> { sel.ChildShapeRange[1].ParentGroup };
                sel.Unselect();
                for (int i = 1 ; i <= gshape[0].GroupItems.Count ; i++)
                    {
                    if (gshape[0].GroupItems[i].Visible == MsoTriState.msoTrue && gshape[0].GroupItems[i].Name != sel.ShapeRange[0].Name)
                        {
                        gshape[0].GroupItems[i].Select(MsoTriState.msoFalse);
                        //shps.Add(gshape[0].GroupItems[i].Name);
                        }
                    }
                //slide.Shapes.Range(shps.ToArray()).Select(MsoTriState.msoTrue);
                Globals.ThisAddIn.Application.CommandBars.ExecuteMso("ObjectsGroup");//组合
                return;
                }
            List<string> shps2 = new List<string>();
            List<PowerPoint.Shape> gshape2 = new List<PowerPoint.Shape> { sel.ShapeRange[1] };
            sel.Unselect();
            for (int j = 1 ; j <= gshape2[0].GroupItems.Count ; j++)
                {
                if (gshape2[0].GroupItems[j].Visible == MsoTriState.msoTrue && gshape2[0].GroupItems[j].Name != sel.ShapeRange[0].Name)
                    {
                    gshape2[0].GroupItems[j].Select(MsoTriState.msoFalse);
                    //shps2.Add(gshape2[0].GroupItems[j].Name);
                    }
                }
            //slide.Shapes.Range(shps2.ToArray()).Select(MsoTriState.msoTrue);
            Globals.ThisAddIn.Application.CommandBars.ExecuteMso("ObjectsGroup");//组合
            }

        private void button79_Click(object sender, RibbonControlEventArgs e)
            {
            var sel = app.ActiveWindow.Selection;

            if ((System.Windows.Forms.Control.ModifierKeys & Keys.Control) == Keys.Control) //判断Ctrl键
                {
                var range = sel.Type == PowerPoint.PpSelectionType.ppSelectionShapes && sel.HasChildShapeRange
                    ? sel.ChildShapeRange
                    : sel.ShapeRange;

                foreach (PowerPoint.Shape shape in range)
                    {
                    if (shape.Type == Office.MsoShapeType.msoGroup)
                        {
                        shape.Ungroup();
                        }
                    }
                }
            else
                {
                try
                    {
                    sel.ShapeRange.Ungroup();
                    }
                catch
                    {
                    // Handle the exception, e.g. log it or show a message to the user.
                    return;
                    }
                }
            }

        private void button94_Click(object sender, RibbonControlEventArgs e)
            {
            if (Control.ModifierKeys.HasFlag(Keys.Control))
                {
                // 获取当前活动窗口的幻灯片
                PowerPoint.Slide slide = app.ActiveWindow.View.Slide;
                // 设置文本框的位置和大小
                float textBoxLeft = (slide.Master.Width - slide.Master.Width / 3) / 2; // 居中
                float textBoxTop = (slide.Master.Height - slide.Master.Width / 3) / 2; // 居中
                float textBoxWidth = slide.Master.Width / 3; // 幻灯片长度的三分之一
                float textBoxHeight = slide.Master.Width / 3; // 幻灯片长度的三分之一

                // 添加文本框
                PowerPoint.Shape textBox = slide.Shapes.AddTextbox(MsoTextOrientation.msoTextOrientationHorizontal, textBoxLeft, textBoxTop, textBoxWidth, textBoxHeight);

                // 设置文本框的文本内容和格式
                textBox.TextFrame.TextRange.Text = "样稿";

                // 设置字体大小
                float fontSize = textBoxHeight / 3;
                textBox.TextFrame.TextRange.Font.Size = fontSize;

                // 设置文本框中的文本居中
                textBox.TextFrame.TextRange.ParagraphFormat.Alignment = PowerPoint.PpParagraphAlignment.ppAlignCenter;
                textBox.TextFrame.VerticalAnchor = MsoVerticalAnchor.msoAnchorMiddle;

                // 设置文本框的字体颜色为灰色
                textBox.TextFrame.TextRange.Font.Color.RGB = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Gray);
                }
            else
                {
                Wpf_waterMark wpf_WaterMark = new Wpf_waterMark();
                wpf_WaterMark.ShowDialog();
                }
            }

        private void button84_Click(object sender, RibbonControlEventArgs e)
            {
            Selection sel = app.ActiveWindow.Selection;
            Presentation pre = app.ActiveWindow.Presentation;
            if (sel.Type == PpSelectionType.ppSelectionShapes && sel.ShapeRange.Count == 1)
                {
                PowerPoint.ShapeRange shr = sel.ShapeRange;
                float Top = shr.Top;
                float Left = shr.Left;
                float Width = shr.Width;
                float Height = shr.Height;
                float Buttom = Top + Height;
                float Right = Left + Width;
                pre.Guides.Add(PpGuideOrientation.ppHorizontalGuide, Top);
                pre.Guides.Add(PpGuideOrientation.ppHorizontalGuide, Buttom);
                pre.Guides.Add(PpGuideOrientation.ppVerticalGuide, Left);
                pre.Guides.Add(PpGuideOrientation.ppVerticalGuide, Right);
                }
            else
                {
                Growl.WarningGlobal("请选择单一对象后操作！");
                }
            }

        private void button93_Click(object sender, RibbonControlEventArgs e)
            {
            Slide slide = app.ActiveWindow.View.Slide;
            int i = app.ActivePresentation.Guides.Count;
            if (i > 0)
                {
                for (int j = i ; j > 0 ; j--)
                    {
                    Guide guide1 = app.ActivePresentation.Guides[j];
                    guide1.Delete();
                    }
                }
            int m = app.ActivePresentation.SlideMaster.Guides.Count;
            if (m > 0)
                {
                for (int j = m ; j > 0 ; j--)
                    {
                    Guide guide2 = app.ActivePresentation.SlideMaster.Guides[j];
                    guide2.Delete();
                    }
                }
            int num = app.ActiveWindow.View.Slide.CustomLayout.Index;//获取当前模板ID
            int n = app.ActivePresentation.SlideMaster.Guides.Count;
            if (n > 0)
                {
                for (int j = n ; j > 0 ; j--)
                    {
                    Guide guide3 = app.ActivePresentation.SlideMaster.CustomLayouts[num].Guides[j];
                    guide3.Delete();
                    }
                }
            }

        private void button95_Click(object sender, RibbonControlEventArgs e)
            {
            }

        private void group4_DialogLauncherClick(object sender, RibbonControlEventArgs e)
            {
            app.CommandBars.ExecuteMso("SelectionPane");//调用选择窗格
            }

        private void button87_Click(object sender, RibbonControlEventArgs e)
            {
            Slide slide = app.ActiveWindow.View.Slide;
            int count = slide.Shapes.Count;
            for (int i = 0 ; i < count ; count--)
                {
                slide.Shapes[count].Select();
                app.CommandBars.ExecuteMso("LockObject");
                }

            Growl.SuccessGlobal("当页对象已锁定！");
            }

        public void LockObj(Slide slide)
            {
            //名称、选择、类型
            int i = slide.Shapes.Count;

            for (int s = 1 ; s <= i ; i--)
                {
                slide.Shapes[i].Select(MsoTriState.msoTrue);
                app.CommandBars.ExecuteMso("LockObject");
                slide.Shapes[i].Select(MsoTriState.msoFalse);
                }

            Growl.SuccessGlobal("全文对象已锁定！");
            }

        private void button88_Click(object sender, RibbonControlEventArgs e)
            {
            Slide slide = app.ActiveWindow.View.Slide;
            int count = slide.Shapes.Count;
            for (int i = 0 ; i < count ; count--)
                {
                slide.Shapes[count].Select();
                app.CommandBars.ExecuteMso("UnlockObject");
                }

            Growl.SuccessGlobal("当页对象已解锁！");
            }

        private void button96_Click(object sender, RibbonControlEventArgs e)
            {
            Selection sel = app.ActiveWindow.Selection;
            if ((System.Windows.Forms.Control.ModifierKeys & Keys.Control) == Keys.Control)  //判断Ctrl键
                {
                if (sel.Type == PpSelectionType.ppSelectionShapes)
                    {
                    foreach (PowerPoint.Shape shp in sel.ShapeRange)
                        {
                        int count = shp.Nodes.Count;
                        for (int i = 0 ; i < count ; count--)
                            {
                            shp.Nodes.SetEditingType(count, MsoEditingType.msoEditingSmooth);//平滑顶点
                            shp.Nodes.SetEditingType(count, MsoEditingType.msoEditingAuto);//平滑顶点
                            }
                        }
                    }
                }
            else
                {
                if (sel.Type == PpSelectionType.ppSelectionShapes)
                    {
                    Globals.ThisAddIn.Application.CommandBars.ExecuteMso("ObjectEditPoints");
                    }
                else
                    {
                    Growl.WarningGlobal("请选择形状后再试！");
                    }
                }
            }

        private void button97_Click(object sender, RibbonControlEventArgs e)
            {
            Selection sel = app.ActiveWindow.Selection;

            if (sel.Type == PpSelectionType.ppSelectionShapes)
                {
                PowerPoint.ShapeRange shr = sel.ShapeRange;
                if (shr.Type == MsoShapeType.msoPicture)
                    {
                    Globals.ThisAddIn.Application.CommandBars.ExecuteMso("PicturesCompress");
                    Growl.WarningGlobal("压缩窗体已开启！");
                    }
                else
                    {
                    Growl.WarningGlobal("请选择图片后再试！");
                    }
                }
            else
                {
                Growl.WarningGlobal("请选择图片后再试！");
                }
            }

        private void splitButton7_Click(object sender, RibbonControlEventArgs e)
            {
            PresPio.Public_Wpf.Wpf_PhotoGallery wpf_PhotoGallery = new PresPio.Public_Wpf.Wpf_PhotoGallery();
            wpf_PhotoGallery.Show();
            }

        private void button12_Click_1(object sender, RibbonControlEventArgs e)
            {
            Selection sel = app.ActiveWindow.Selection;

            if (sel.Type == PpSelectionType.ppSelectionShapes)
                {
                PowerPoint.ShapeRange shr = sel.ShapeRange;
                if (shr.Type == MsoShapeType.msoPicture)
                    {
                    Globals.ThisAddIn.Application.CommandBars.ExecuteMso("PictureResetAndSize");
                    }
                else
                    {
                    Growl.WarningGlobal("请选择图片后再试！");
                    }
                }
            else
                {
                Growl.WarningGlobal("请选择图片后再试！");
                }
            }

        private void gallery1_ButtonClick(object sender, RibbonControlEventArgs e)
            {
            Selection sel = app.ActiveWindow.Selection;

            if (sel.Type == PpSelectionType.ppSelectionShapes)
                {
                PowerPoint.ShapeRange shr = sel.ShapeRange;
                if (shr.Type == MsoShapeType.msoPicture)
                    {
                    app.CommandBars.ExecuteMso("PictureCrop");
                    }
                else
                    {
                    Growl.WarningGlobal("请选择图片后再试！");
                    }
                }
            else
                {
                Growl.WarningGlobal("请选择图片后再试！");
                }
            }

        private void button99_Click(object sender, RibbonControlEventArgs e)
            {
            Selection sel = app.ActiveWindow.Selection;
            if (sel.Type != PpSelectionType.ppSelectionNone)
                {
                PowerPoint.ShapeRange shr = sel.ShapeRange;
                app.CommandBars.ExecuteMso("ObjectSaveAsPicture");
                }
            else
                {
                Growl.WarningGlobal("请选择内容后再试！");
                }
            }

        private void button78_Click(object sender, RibbonControlEventArgs e)
            {
            Globals.ThisAddIn.Application.CommandBars.ExecuteMso("ObjectsGroup");//组合
            }

        private void button95_Click_1(object sender, RibbonControlEventArgs e)
            {
            app.CommandBars.ExecuteMso("AddHorizontalGuide");
            }

        private void button100_Click(object sender, RibbonControlEventArgs e)
            {
            app.CommandBars.ExecuteMso("AddVerticalGuide");
            }

        private void button101_Click(object sender, RibbonControlEventArgs e)
            {
            app.CommandBars.ExecuteMso("GridSettings");
            }

        private void button102_Click(object sender, RibbonControlEventArgs e)
            {
            app.CommandBars.ExecuteMso("SectionAdd");
            }

        private void button103_Click(object sender, RibbonControlEventArgs e)
            {
            Growl.WarningGlobal("此功能要在存在节的文档中使用哦！");
            app.CommandBars.ExecuteMso("SectionCollapseAll");
            }

        private void button104_Click(object sender, RibbonControlEventArgs e)
            {
            app.CommandBars.ExecuteMso("SectionExpandAll");
            Growl.SuccessGlobal("所有分节已展开！");
            }

        private void button105_Click(object sender, RibbonControlEventArgs e)
            {
            app.CommandBars.ExecuteMso("SectionMergeWithPrevious");
            Growl.SuccessGlobal("当前分节已移除！");
            }

        private void button106_Click(object sender, RibbonControlEventArgs e)
            {
            app.CommandBars.ExecuteMso("SectionRemoveAll");
            Growl.SuccessGlobal("所有分节已移除！");
            }

        private void button107_Click(object sender, RibbonControlEventArgs e)
            {
            app.CommandBars.ExecuteMso("LockObject");

            Growl.SuccessGlobal("对象已锁定！");
            }

        private void button108_Click(object sender, RibbonControlEventArgs e)
            {
            app.CommandBars.ExecuteMso("UnlockObject");

            Growl.SuccessGlobal("对象已解锁！");
            }

        private void button76_Click(object sender, RibbonControlEventArgs e)
            {
            Selection sel = app.ActiveWindow.Selection;
            if (sel.Type == PpSelectionType.ppSelectionShapes && sel.ShapeRange.Count > 0)
                {
                Slide slide = app.ActiveWindow.View.Slide;
                List<PowerPoint.Shape> oshapes = new List<PowerPoint.Shape>();

                foreach (PowerPoint.Shape shape in sel.ShapeRange)
                    {
                    oshapes.Add(shape);
                    }

                foreach (var oshape in oshapes)
                    {
                    if (oshape.Type == MsoShapeType.msoTextBox)
                        {
                        float Width = oshape.Width;  // 原图宽度
                        float Height = oshape.Height;
                        float Top = oshape.Top;      // 原图位置
                        float Left = oshape.Left;

                        // 创建新的矩形形状
                        PowerPoint.Shape newShape = slide.Shapes.AddShape(MsoAutoShapeType.msoShapeRectangle, Left, Top, Width, Height);

                        // 选择要操作的文本框和新创建的矩形
                        oshape.Select(MsoTriState.msoTrue);
                        newShape.Select(MsoTriState.msoFalse);

                        // 执行减去命令
                        string strCmd = "ShapesIntersect";
                        Globals.ThisAddIn.Application.CommandBars.ExecuteMso(strCmd); // 执行减去

                        // 检查新创建的矩形形状是否还存在并删除
                        try
                            {
                            newShape.Delete();
                            }
                        catch (Exception ex)
                            {
                            // 处理形状删除失败的情况
                            Debug.WriteLine("删除新创建的矩形形状时出错: " + ex.Message);
                            }
                        }
                    }
                }
            else
                {
                Growl.Warning("请选择一个或多个文本框内容后再操作！", "温馨提示");
                }
            }

        private void button109_Click(object sender, RibbonControlEventArgs e)
            {
            PowerPoint.Selection sel = app.ActiveWindow.Selection;
            if (sel.Type == PpSelectionType.ppSelectionSlides)
                {
                sel.SlideRange.SlideShowTransition.Hidden = MsoTriState.msoTrue;
                }
            else if (sel.Type == PowerPoint.PpSelectionType.ppSelectionShapes)
                {
                PowerPoint.ShapeRange range = sel.ShapeRange;
                if (sel.HasChildShapeRange)
                    {
                    range = sel.ChildShapeRange;
                    }
                for (int i = 1 ; i <= range.Count ; i++)
                    {
                    PowerPoint.Shape shape = range[i];
                    if (shape.Visible == Office.MsoTriState.msoTrue)
                        {
                        shape.Visible = Office.MsoTriState.msoFalse;
                        }
                    }
                }
            else
                {
                Growl.WarningGlobal("请选择对象后再试！");
                }
            }

        private void button110_Click(object sender, RibbonControlEventArgs e)
            {
            Selection sel = app.ActiveWindow.Selection;
            Slide myDocument = app.ActiveWindow.View.Slide;
            Presentations pre = app.Presentations;
            int count = myDocument.Shapes.Count;
            List<PowerPoint.Shape> shps = new List<PowerPoint.Shape>();
            for (int i = 0 ; i <= count ; count--)
                {
                if (myDocument.Shapes[count].Visible == MsoTriState.msoFalse)
                    {
                    myDocument.Shapes[count].Visible = MsoTriState.msoTrue;
                    break;
                    }
                }
            }

        private void button111_Click(object sender, RibbonControlEventArgs e)
            {
            Globals.ThisAddIn.Application.CommandBars.ExecuteMso("ObjectsGroup");//组合
            }

        private void button112_Click(object sender, RibbonControlEventArgs e)
            {
            Globals.ThisAddIn.Application.CommandBars.ExecuteMso("ObjectsUngroup");//组合
            }

        private void button85_Click(object sender, RibbonControlEventArgs e)
            {
            PowerPoint.Selection sel = app.ActiveWindow.Selection;
            if (sel.Type == PowerPoint.PpSelectionType.ppSelectionShapes)
                {
                PowerPoint.Slide slide = app.ActiveWindow.View.Slide;
                PowerPoint.ShapeRange range = sel.ShapeRange;
                List<string> oname = new List<string>();
                List<string> aname = new List<string>();
                if (sel.HasChildShapeRange)
                    {
                    range = sel.ChildShapeRange;
                    foreach (PowerPoint.Shape shape in range)
                        {
                        oname.Add(shape.Name);
                        }
                    foreach (PowerPoint.Shape shape in range[1].ParentGroup.GroupItems)
                        {
                        aname.Add(shape.Name);
                        }
                    for (int i = 0 ; i < oname.Count() ; i++)
                        {
                        if (aname.Contains(oname[i]))
                            {
                            aname.Remove(oname[i]);
                            }
                        }
                    }
                else
                    {
                    foreach (PowerPoint.Shape shape in range)
                        {
                        oname.Add(shape.Name);
                        }
                    foreach (PowerPoint.Shape shape in slide.Shapes)
                        {
                        aname.Add(shape.Name);
                        }
                    for (int i = 0 ; i < oname.Count() ; i++)
                        {
                        if (aname.Contains(oname[i]))
                            {
                            aname.Remove(oname[i]);
                            }
                        }
                    }
                if (aname.Count() != 0)
                    {
                    slide.Shapes.Range(aname.ToArray()).Select();
                    }
                }
            else if (sel.Type == PowerPoint.PpSelectionType.ppSelectionSlides)
                {
                PowerPoint.Slides slides = app.ActivePresentation.Slides;
                PowerPoint.SlideRange srange = sel.SlideRange;
                List<int> oslides = new List<int>();
                List<int> aslides = new List<int>();
                foreach (PowerPoint.Slide slide in srange)
                    {
                    oslides.Add(slide.SlideIndex);
                    }
                foreach (PowerPoint.Slide slide in slides)
                    {
                    aslides.Add(slide.SlideIndex);
                    }
                for (int i = 0 ; i < oslides.Count() ; i++)
                    {
                    if (aslides.Contains(oslides[i]))
                        {
                        aslides.Remove(oslides[i]);
                        }
                    }
                if (aslides.Count() != 0)
                    {
                    slides.Range(aslides.ToArray()).Select();
                    }
                }
            else
                {
                Growl.WarningGlobal("请先选择形状！");
                }
            }

        private void button115_Click(object sender, RibbonControlEventArgs e)
            {
            app = Globals.ThisAddIn.Application;
            Slide slide = app.ActiveWindow.View.Slide;
            Selection sel = app.ActiveWindow.Selection;
            if (sel.Type != PpSelectionType.ppSelectionShapes)
                {
                return;
                }
            else
                {
                PowerPoint.Shape shp = sel.ShapeRange[1];
                foreach (PowerPoint.Shape item in slide.Shapes)
                    {
                    if (item.Type == shp.Type && item.Visible == MsoTriState.msoTrue)
                        {
                        item.Select(MsoTriState.msoFalse);
                        }
                    else
                        {
                        return;
                        }
                    }
                }
            }

        private void button116_Click(object sender, RibbonControlEventArgs e)
            {
            Selection sel = this.app.ActiveWindow.Selection;

            if (sel.Type != PpSelectionType.ppSelectionShapes)
                {
                Growl.WarningGlobal("请选择组合后再试！");
                return;
                }
            if (sel.HasChildShapeRange)
                {
                List<PowerPoint.Shape> gshape = new List<PowerPoint.Shape> { sel.ChildShapeRange[1].ParentGroup };
                sel.Unselect();
                for (int i = 1 ; i <= gshape[0].GroupItems.Count ; i++)
                    {
                    if (gshape[0].GroupItems[i].Visible == MsoTriState.msoTrue)
                        {
                        gshape[0].GroupItems[i].Select(MsoTriState.msoFalse);
                        }
                    }
                return;
                }
            List<PowerPoint.Shape> gshape2 = new List<PowerPoint.Shape> { sel.ShapeRange[1] };
            sel.Unselect();
            for (int j = 1 ; j <= gshape2[0].GroupItems.Count ; j++)
                {
                if (gshape2[0].GroupItems[j].Visible == MsoTriState.msoTrue)
                    {
                    gshape2[0].GroupItems[j].Select(MsoTriState.msoFalse);
                    }
                }
            }

        private void button117_Click(object sender, RibbonControlEventArgs e)
            {
            var sel = app.ActiveWindow.Selection;
            var slide = app.ActiveWindow.View.Slide;
            var shps = new List<string>();

            MyFunction F = new MyFunction();

            if (sel.Type != PpSelectionType.ppSelectionShapes)
                {
                Growl.WarningGlobal("请选择内容后再试！");
                }
            else
                {
                foreach (PowerPoint.Shape shp in sel.ShapeRange)
                    {
                    float oldTop = shp.Top;
                    float oldLeft = shp.Left;
                    float oldWidth = shp.Width;
                    float oldHeight = shp.Height;
                    float ceNTop = oldTop + oldHeight / 2;
                    float ceNLeft = oldLeft + oldWidth / 2;

                    var (newWidth, newHeight) = oldHeight >= oldWidth
                        ? (oldHeight + oldHeight, oldHeight + oldHeight)
                        : (oldWidth + oldWidth, oldWidth + oldWidth);

                    var newTop = ceNTop - newHeight / 2;
                    var newLeft = ceNLeft - newWidth / 2;

                    var newShp = slide.Shapes.AddShape(MsoAutoShapeType.msoShapeOval, newLeft, newTop, newWidth, newHeight);

                    newShp.Fill.ForeColor.RGB = F.RGB2Int(255, 255, 255);
                    newShp.Line.ForeColor.RGB = F.RGB2Int(216, 216, 216);
                    newShp.Line.Weight = 0;
                    newShp.Shadow.Type = MsoShadowType.msoShadow23;
                    newShp.Shadow.ForeColor.RGB = F.RGB2Int(0, 0, 128);
                    newShp.Shadow.Transparency = 0.9f;
                    newShp.Shadow.OffsetX = 0;
                    newShp.Shadow.OffsetY = 0;
                    newShp.ZOrder(MsoZOrderCmd.msoSendToBack);
                    shps.Add(newShp.Name);
                    }

                sel.Unselect();
                slide.Shapes.Range(shps.ToArray()).Select(MsoTriState.msoTrue);
                }
            }

        private void button118_Click(object sender, RibbonControlEventArgs e)
            {
            Selection sel = app.ActiveWindow.Selection;
            Slide slide = app.ActiveWindow.View.Slide;
            MyFunction F = new MyFunction();

            if (sel.Type != PpSelectionType.ppSelectionShapes)
                {
                Growl.WarningGlobal("请选择内容后再试！");
                }
            else
                {
                List<PowerPoint.Shape> newShapes = new List<PowerPoint.Shape>();

                foreach (PowerPoint.Shape selectedShape in sel.ShapeRange)
                    {
                    float oldTop = selectedShape.Top;
                    float oldLeft = selectedShape.Left;
                    float oldWidth = selectedShape.Width;
                    float oldHeight = selectedShape.Height;
                    float ceNTop = oldTop + oldHeight / 2;
                    float ceNLeft = oldLeft + oldWidth / 2;

                    float newWidth, newHeight, newTop, newLeft;
                    if (oldHeight >= oldWidth)
                        {
                        newWidth = oldHeight / 4 + oldWidth;
                        newHeight = oldHeight + oldHeight / 2;
                        newTop = ceNTop - newHeight / 2;
                        newLeft = ceNLeft - newWidth / 2;
                        }
                    else
                        {
                        newWidth = oldWidth * 2;
                        newHeight = newWidth / 4 + oldHeight;
                        newTop = ceNTop - newHeight / 2;
                        newLeft = ceNLeft - newWidth / 2;
                        }

                    PowerPoint.Shape newShape = slide.Shapes.AddShape(MsoAutoShapeType.msoShapeRectangle, newLeft, newTop, newWidth, newHeight);
                    newShape.Fill.ForeColor.RGB = F.RGB2Int(255, 255, 255);
                    newShape.Line.ForeColor.RGB = F.RGB2Int(216, 216, 216);
                    newShape.Line.Weight = 0;
                    newShape.Shadow.Type = MsoShadowType.msoShadow23;
                    newShape.Shadow.ForeColor.RGB = F.RGB2Int(0, 0, 128);
                    newShape.Shadow.Transparency = 0.9f;
                    newShape.Shadow.OffsetX = 0;
                    newShape.Shadow.OffsetY = 0;
                    newShape.ZOrder(MsoZOrderCmd.msoSendToBack);
                    newShapes.Add(newShape);
                    }

                sel.Unselect();
                PowerPoint.ShapeRange newShapeRange = slide.Shapes.Range(newShapes.Select(shape => shape.Name).ToArray());
                newShapeRange.Select(MsoTriState.msoTrue);
                }
            }

        private void toggleButton4_Click(object sender, RibbonControlEventArgs e)
            {
            //自适应窗格剪切板
            #region
            if (toggleButton4.Checked)
                {
                // 初始化TaskPane
                Con_TextTools con_Tools = new Con_TextTools();
                TaskPaneShared.taskPane = Globals.ThisAddIn.CustomTaskPanes.Add(con_Tools, "文本助手");
                TaskPaneShared.taskPane.DockPosition = MsoCTPDockPosition.msoCTPDockPositionRight;
                TaskPaneShared.taskPane.DockPositionRestrict = MsoCTPDockPositionRestrict.msoCTPDockPositionRestrictNone;
                TaskPaneShared.taskPane.Width = 340;
                TaskPaneShared.taskPane.Visible = true;
                }
            else
                {
                // 隐藏并移除TaskPane
                if (TaskPaneShared.taskPane != null)
                    {
                    TaskPaneShared.taskPane.Visible = false;
                    Globals.ThisAddIn.CustomTaskPanes.Remove(TaskPaneShared.taskPane);
                    TaskPaneShared.taskPane.Dispose();
                    TaskPaneShared.taskPane = null; // 清理引用
                    }
                }

            // VisibleChange Event
            if (TaskPaneShared.taskPane != null)
                {
                TaskPaneShared.taskPane.VisibleChanged += new System.EventHandler(taskpane4_VisibleChanged);
                }
            #endregion
            }

        private void taskpane4_VisibleChanged(object sender, EventArgs e)//回调用户窗体事件
            {
            MyRibbon ribbon = Globals.Ribbons.GetRibbon<MyRibbon>();//获得功能区
            if (TaskPaneShared.taskPane.Visible)
                {
                ribbon.toggleButton4.Checked = true;
                }
            else
                {
                ribbon.toggleButton4.Checked = false;
                }
            }

        /// <summary>
        /// 单例窗体调用 GenericSingleton
        /// </summary>
        /// <typeparam name="T"></typeparam>
        public class GenericSingleton<T> where T : Form, new()
            {
            private static T t = null;

            public static T CreateInstrance()
                {
                if (t == null || t.IsDisposed)
                    {
                    t = new T();
                    }
                else
                    {
                    t.Activate(); //如果已经打开过就让其获得焦点
                    t.WindowState = FormWindowState.Normal;//使Form恢复正常窗体大小
                    }
                return t;
                }
            }

        private void group5_DialogLauncherClick(object sender, RibbonControlEventArgs e)
            {
            //Wpf_Tools wpf_Tools = new Wpf_Tools();
            //wpf_Tools.Show();
            //var ribbon = Globals.Ribbons.Ribbon1; // 获取功能区实例
            //ribbon.group5.Visible = false;
            }

        private void button119_Click(object sender, RibbonControlEventArgs e)
            {
            }

        private void button108_Click_1(object sender, RibbonControlEventArgs e)
            {
            app = Globals.ThisAddIn.Application;
            Selection sel = app.ActiveWindow.Selection;
            if (sel.Type == PpSelectionType.ppSelectionShapes)
                {
                int count = sel.ShapeRange.Count;
                for (int i = 1 ; i <= count ; i++)
                    {
                    PowerPoint.Shape shp2 = sel.ShapeRange[i];
                    string Text = i.ToString();
                    if (i < 10)
                        {
                        Text = Text.PadLeft(2, '0');
                        }
                    shp2.TextFrame.TextRange.Text = Text;
                    }
                }
            else
                {
                Growl.WarningGlobal("请选择形状后再试！");
                }
            }

        private void button89_Click(object sender, RibbonControlEventArgs e)
            {
            Presentation presentation = app.ActivePresentation;
            Slide slide = app.ActiveWindow.View.Slide;
            Selection selection = app.ActiveWindow.Selection;

            // 获取临时文件夹路径
            string tempFolderPath = Path.GetTempPath();

            // 禁用幻灯片 1 的跟随母版背景
            presentation.Slides[1].FollowMasterBackground = MsoTriState.msoFalse;

            // 检查选择是否为图形
            if (selection.Type == PpSelectionType.ppSelectionShapes)
                {
                PowerPoint.Shape shape = selection.ShapeRange[1];
                // 检查选定的形状是否为图片
                if (shape.Type == MsoShapeType.msoPicture)
                    {
                    try
                        {
                        // 导出图片
                        string fileName = DateTime.Now.ToFileTimeUtc().ToString() + ".png";
                        string exportedImagePath = Path.Combine(tempFolderPath, fileName);
                        shape.Export(exportedImagePath, PpShapeFormat.ppShapeFormatPNG, 0, 0);

                        // 设置导出的图片为 PowerPoint 背景
                        slide.Background.Fill.UserPicture(exportedImagePath);
                        }
                    catch (Exception ex)
                        {
                        // 处理异常
                        Console.WriteLine("Error exporting image and setting background: " + ex.Message);
                        }
                    }
                }
            }

        private void button7_Click(object sender, RibbonControlEventArgs e)
            {
            }

        private void button109_Click_1(object sender, RibbonControlEventArgs e)
            {
            Wpf_shapeCohesion wpf_ShapeCohesion = new Wpf_shapeCohesion();
            wpf_ShapeCohesion.Show();
            }

        /// <summary>
        /// 获取形状位置坐标
        /// </summary>
        public float[] GetPosation()
            {
            float x = 2, y = 1;
            float[] result = new float[] { x, y };
            Selection sel = app.ActiveWindow.Selection;
            Slide slide = app.ActiveWindow.View.Slide;
            return result;
            }

        private void button110_Click_1(object sender, RibbonControlEventArgs e)
            {
            Wpf_ShpStyle wpf_ShpStyle = new Wpf_ShpStyle();
            wpf_ShpStyle.Show();
            //Form_ShpStyle ShpStyle = GenericSingleton<Form_ShpStyle>.CreateInstrance();
            //ShpStyle.ShowDialog();
            }

        private void button111_Click_1(object sender, RibbonControlEventArgs e)
            {
            PowerPoint.Selection sel = app.ActiveWindow.Selection;
            if (sel.Type == PowerPoint.PpSelectionType.ppSelectionShapes)
                {
                PowerPoint.Slide slide = app.ActiveWindow.View.Slide;
                float w = app.ActivePresentation.PageSetup.SlideWidth;
                float h = app.ActivePresentation.PageSetup.SlideHeight;
                PowerPoint.Shape shape = slide.Shapes.AddShape(Office.MsoAutoShapeType.msoShapeRectangle, 0, 0, w, h);
                if (shape.Line.Visible == Office.MsoTriState.msoTrue)
                    {
                    shape.Line.Visible = Office.MsoTriState.msoFalse;
                    }
                shape.Line.Weight = 0;
                shape.Select();
                }
            else
                {
                PowerPoint.SlideRange srange = sel.SlideRange;
                float w = app.ActivePresentation.PageSetup.SlideWidth;
                float h = app.ActivePresentation.PageSetup.SlideHeight;
                if (sel.SlideRange.Count == 1)
                    {
                    PowerPoint.Shape shape = sel.SlideRange[1].Shapes.AddShape(Office.MsoAutoShapeType.msoShapeRoundedRectangle, 0, 0, w, h);
                    if (shape.Line.Visible == Office.MsoTriState.msoTrue)
                        {
                        shape.Line.Visible = Office.MsoTriState.msoFalse;
                        }
                    shape.Line.Weight = 0;
                    shape.Select();
                    }
                else
                    {
                    foreach (PowerPoint.Slide item in srange)
                        {
                        PowerPoint.Shape shape = item.Shapes.AddShape(Office.MsoAutoShapeType.msoShapeRoundedRectangle, 0, 0, w, h);
                        if (shape.Line.Visible == Office.MsoTriState.msoTrue)
                            {
                            shape.Line.Visible = Office.MsoTriState.msoFalse;
                            }
                        shape.Line.Weight = 0;
                        }
                    }
                }
            }

        private void button98_Click(object sender, RibbonControlEventArgs e)
            {
            }

        private void button59_Click(object sender, RibbonControlEventArgs e)
            {
            Selection sel = app.ActiveWindow.Selection;
            Slide slide = app.ActiveWindow.View.Slide;
            List<string> shps = new List<string> { };
            MyFunction F = new MyFunction();
            if (sel.Type != PpSelectionType.ppSelectionShapes)
                {
                Growl.WarningGlobal("请选择内容后再试！");
                }
            else
                {
                int count = sel.ShapeRange.Count;
                for (int i = 0 ; i < count ; count--)
                    {
                    PowerPoint.Shape shp = sel.ShapeRange[count];//选择的图形
                    float oldTop = shp.Top;
                    float oldLeft = shp.Left;
                    float oldWidth = shp.Width;
                    float oldHeight = shp.Height;
                    float ceNTop = oldTop + oldHeight;
                    float ceNLeft = oldLeft + oldWidth;
                    //添加角标
                    PowerPoint.Shape newShp = slide.Shapes.AddShape(MsoAutoShapeType.msoShapeOval, ceNLeft - oldHeight / 4, ceNTop - oldHeight / 4, oldHeight / 2, oldHeight / 2);
                    newShp.ZOrder(MsoZOrderCmd.msoSendToBack);
                    newShp.Line.Visible = MsoTriState.msoFalse;
                    shps.Add(newShp.Name);
                    }
                }
            sel.Unselect();
            slide.Shapes.Range(shps.ToArray()).Select(MsoTriState.msoTrue);
            }

        private void button98_Click_1(object sender, RibbonControlEventArgs e)
            {
            }

        private void button98_Click_2(object sender, RibbonControlEventArgs e)
            {
            }

        private void button112_Click_1(object sender, RibbonControlEventArgs e)
            {
            Selection sel = app.ActiveWindow.Selection;
            if (sel.Type == PpSelectionType.ppSelectionShapes)
                {
                if (sel.ShapeRange.TextFrame.AutoSize == PpAutoSize.ppAutoSizeNone)
                    {
                    sel.ShapeRange.TextFrame.AutoSize = PpAutoSize.ppAutoSizeShapeToFitText;
                    }
                else
                    {
                    sel.ShapeRange.TextFrame.AutoSize = PpAutoSize.ppAutoSizeNone;
                    }
                }
            }

        private void button113_Click(object sender, RibbonControlEventArgs e)
            {
            try
                {
                Selection sel = app.ActiveWindow.Selection;
                if (sel.Type == PpSelectionType.ppSelectionShapes)
                    {
                    foreach (PowerPoint.Shape shp in sel.ShapeRange)
                        {
                        int count = shp.Nodes.Count;
                        for (int i = 0 ; i < count ; count--)
                            {
                            shp.Nodes.SetEditingType(count, MsoEditingType.msoEditingSmooth);//平滑顶点
                            shp.Nodes.SetEditingType(count, MsoEditingType.msoEditingAuto);//平滑顶点
                            }
                        }
                    }
                }
            catch
                {
                return;
                }
            }

        private void button62_Click(object sender, RibbonControlEventArgs e)
            {
            Selection sel = app.ActiveWindow.Selection;
            Slide slide = app.ActiveWindow.View.Slide;
            List<string> shps = new List<string> { };
            MyFunction F = new MyFunction();
            if (sel.Type != PpSelectionType.ppSelectionShapes)
                {
                Growl.WarningGlobal("请选择内容后再试！");
                }
            else
                {
                int count = sel.ShapeRange.Count;
                for (int i = 0 ; i < count ; count--)
                    {
                    PowerPoint.Shape shp = sel.ShapeRange[count];//选择的图形
                    float oldTop = shp.Top;
                    float oldLeft = shp.Left;
                    float oldWidth = shp.Width;
                    float oldHeight = shp.Height;
                    float ceNTop = oldTop + oldHeight;
                    float ceNLeft = oldLeft + oldWidth;
                    //添加角标
                    PowerPoint.Shape newShp = slide.Shapes.AddShape(MsoAutoShapeType.msoShapeOval, ceNLeft - oldHeight / 4, ceNTop - oldHeight / 4, oldHeight / 2, oldHeight / 2);
                    newShp.ZOrder(MsoZOrderCmd.msoSendToBack);
                    newShp.Line.Visible = MsoTriState.msoFalse;
                    shps.Add(newShp.Name);
                    }
                }
            sel.Unselect();
            slide.Shapes.Range(shps.ToArray()).Select(MsoTriState.msoTrue);
            }

        private void button114_Click(object sender, RibbonControlEventArgs e)
            {
            }

        private void button119_Click_1(object sender, RibbonControlEventArgs e)
            {
            Selection sel = app.ActiveWindow.Selection;
            if (sel.Type == PpSelectionType.ppSelectionShapes && sel.ShapeRange.Count > 0)
                {
                Slide slide = app.ActiveWindow.View.Slide;
                List<PowerPoint.Shape> oshapes = new List<PowerPoint.Shape>();

                foreach (PowerPoint.Shape shape in sel.ShapeRange)
                    {
                    oshapes.Add(shape);
                    }

                foreach (var oshape in oshapes)
                    {
                    if (oshape.Type == MsoShapeType.msoTextBox)
                        {
                        // 选中文本框
                        oshape.Select(MsoTriState.msoTrue);

                        // 执行“转换为形状”命令
                        string strCmd = "ConvertTextToSmartArt";
                        Globals.ThisAddIn.Application.CommandBars.ExecuteMso(strCmd); // 转换为形状

                        // 在这里可以进行进一步处理，比如调整形状属性等
                        }
                    }
                }
            else
                {
                Growl.Warning("请选择一个或多个文本框内容后再操作！", "温馨提示");
                }
            }

        private void button77_Click(object sender, RibbonControlEventArgs e)
            {
            Selection sel = app.ActiveWindow.Selection;
            Slide slide = app.ActiveWindow.View.Slide;
            List<string> shps = new List<string>();

            if (sel.Type != PpSelectionType.ppSelectionShapes)
                {
                Growl.WarningGlobal("请选择文字后再试！");
                }
            else
                {
                PowerPoint.Shape shp2 = slide.Shapes.AddShape(MsoAutoShapeType.msoShapeRectangle, -30, 0, 30, 30);
                shps.Add(shp2.Name);
                int count = sel.ShapeRange.Count;
                for (int i = 0 ; i < count ; count--)
                    {
                    PowerPoint.Shape shp = sel.ShapeRange[count];//选择的图形
                    shps.Add(shp.Name);
                    }
                slide.Shapes.Range(shps.ToArray()).Select(MsoTriState.msoTrue);
                sel.ShapeRange.MergeShapes(MsoMergeCmd.msoMergeFragment, null);
                foreach (PowerPoint.Shape shp in slide.Shapes)
                    {
                    if (shp.Left == -30)
                        {
                        shp.Delete();
                        }
                    }
                }
            }

        private void button120_Click(object sender, RibbonControlEventArgs e)
            {
            PowerPoint.Selection sel = app.ActiveWindow.Selection;
            PowerPoint.Slide slide = app.ActiveWindow.View.Slide;
            if (sel.Type == PowerPoint.PpSelectionType.ppSelectionNone)
                {
                Growl.WarningGlobal("请选中至少1个文本框");
                return;
                }

            PowerPoint.ShapeRange range = sel.ShapeRange;
            if (sel.HasChildShapeRange)
                {
                range = sel.ChildShapeRange;
                }

            foreach (PowerPoint.Shape shape in range)
                {
                string txt = shape.TextEffect.Text;
                if (txt.Contains("\r") || txt.Contains("\v"))
                    {
                    string[] arr = txt.Split(new string[] { "\r", "\v" }, StringSplitOptions.None);
                    int tcount = arr.Length;
                    shape.PickUp();
                    for (int j = 1 ; j <= tcount ; j++)
                        {
                        float left = shape.Left + shape.Width;
                        float top = shape.Top + shape.Height * (j - 1) / tcount;
                        float width = shape.Width;
                        float height = shape.Height / tcount;

                        PowerPoint.Shape text = slide.Shapes.AddTextbox(Office.MsoTextOrientation.msoTextOrientationHorizontal, left, top, width, height);
                        text.TextFrame.TextRange.Text = arr[j - 1];
                        text.Apply();
                        text.TextFrame.AutoSize = PowerPoint.PpAutoSize.ppAutoSizeShapeToFitText;
                        }
                    }
                else
                    {
                    Growl.SuccessGlobal("存在没有分段的文本框");
                    }
                }
            }

        private void button121_Click(object sender, RibbonControlEventArgs e)
            {
            Selection sel = app.ActiveWindow.Selection;
            Slide slide = app.ActiveWindow.View.Slide;

            // 检查是否按下了Ctrl键
            if (Control.ModifierKeys.HasFlag(Keys.Control))
                {
                // 如果按下了 Ctrl 键，但没有选中任何形状
                if (sel.Type == PpSelectionType.ppSelectionNone)
                    {
                    Growl.WarningGlobal("请选中至少1个文本框");
                    }
                else
                    {
                    PowerPoint.ShapeRange range = sel.HasChildShapeRange ? sel.ChildShapeRange : sel.ShapeRange;

                    // 遍历选定的每个形状
                    foreach (PowerPoint.Shape shape in range)
                        {
                        if (shape.HasTextFrame == MsoTriState.msoTrue && shape.TextFrame.HasText == MsoTriState.msoTrue)
                            {
                            string txt = shape.TextFrame.TextRange.Text;
                            int tcount = txt.Length;

                            // 处理每个字符
                            for (int j = 0 ; j < tcount ; j++)
                                {
                                float left = shape.Left + shape.Width + 24 * j;
                                float top = shape.Top;
                                float width = 24;
                                float height = 20;

                                // 创建新的文本框形状
                                PowerPoint.Shape text = slide.Shapes.AddTextbox(MsoTextOrientation.msoTextOrientationHorizontal, left, top, width, height);
                                text.TextFrame.TextRange.Text = txt.Substring(j, 1);

                                // 获取原始字符的字体格式
                                var charFont = shape.TextFrame2.TextRange.Characters[j + 1].Font;
                                var textFont = text.TextFrame2.TextRange.Font;

                                // 应用字体格式
                                textFont.Size = charFont.Size;
                                textFont.Name = charFont.Name;
                                textFont.NameFarEast = charFont.NameFarEast;
                                textFont.Fill.ForeColor.RGB = charFont.Fill.ForeColor.RGB;
                                textFont.Fill.Transparency = charFont.Fill.Transparency;

                                // 自动调整文本框大小以适应文本
                                text.TextFrame.AutoSize = PpAutoSize.ppAutoSizeShapeToFitText;
                                }

                            // 删除原始形状
                            //   shape.Delete();
                            }
                        else
                            {
                            Growl.WarningGlobal("请选择一个包含文本的文本框。");
                            }
                        }
                    }
                }
            else
                {
                // 如果未按下Ctrl键，执行标准文本拆分逻辑
                if (sel.Type == PpSelectionType.ppSelectionShapes)
                    {
                    PowerPoint.ShapeRange shapeRange = sel.ShapeRange;

                    // 遍历选定的每个形状
                    foreach (PowerPoint.Shape shape in shapeRange)
                        {
                        if (shape.HasTextFrame == MsoTriState.msoTrue && shape.TextFrame.HasText == MsoTriState.msoTrue)
                            {
                            string text = shape.TextFrame.TextRange.Text;
                            float left = shape.Left;
                            float top = shape.Top;
                            float widthPerChar = shape.Width / text.Length;
                            float height = shape.Height;

                            // 将每个字符分离成独立的形状
                            for (int i = 0 ; i < text.Length ; i++)
                                {
                                PowerPoint.Shape newShape = shape.Duplicate()[1];
                                newShape.Left = left + i * widthPerChar;
                                newShape.Top = top;
                                newShape.Width = widthPerChar;
                                newShape.Height = height;
                                newShape.TextFrame.TextRange.Text = text[i].ToString();
                                }

                            // 删除原始形状
                            shape.Delete();
                            }
                        else
                            {
                            Growl.WarningGlobal("请选择一个包含文本的文本框。");
                            }
                        }
                    }
                else
                    {
                    Growl.WarningGlobal("请选择一个文本框。");
                    }
                }
            }

        private void button122_Click(object sender, RibbonControlEventArgs e)
            {
            PowerPoint.Selection sel = app.ActiveWindow.Selection;
            PowerPoint.Slide slide = app.ActiveWindow.View.Slide;

            if (sel.Type == PowerPoint.PpSelectionType.ppSelectionNone)
                {
                Growl.WarningGlobal("请选中要合并的文本框，或者幻灯片");
                }
            else
                {
                PowerPoint.Shape text = slide.Shapes.AddTextbox(Office.MsoTextOrientation.msoTextOrientationHorizontal, -200, 0, 200, 200);
                text.TextFrame.AutoSize = PowerPoint.PpAutoSize.ppAutoSizeShapeToFitText;
                text.Name = "textcount";
                text.TextFrame.TextRange.Text = "";

                PowerPoint.ShapeRange range = null;
                if (sel.Type == PowerPoint.PpSelectionType.ppSelectionSlides)
                    {
                    range = (PowerPoint.ShapeRange)sel.SlideRange.Shapes;
                    }
                else if (sel.Type == PowerPoint.PpSelectionType.ppSelectionShapes)
                    {
                    range = sel.ShapeRange;
                    if (sel.HasChildShapeRange)
                        {
                        range = sel.ChildShapeRange;
                        }
                    }

                if (range != null)
                    {
                    foreach (PowerPoint.Shape shape in range)
                        {
                        if (shape.Type == Office.MsoShapeType.msoGroup)
                            {
                            foreach (PowerPoint.Shape groupShape in shape.GroupItems)
                                {
                                if (groupShape.HasTextFrame == Office.MsoTriState.msoTrue && groupShape.Name != "textcount")
                                    {
                                    text.TextFrame.TextRange.Text += Environment.NewLine + groupShape.TextFrame.TextRange.Text;
                                    }
                                }
                            }
                        else if (shape.HasTextFrame == Office.MsoTriState.msoTrue && shape.Name != "textcount")
                            {
                            text.TextFrame.TextRange.Text += Environment.NewLine + shape.TextFrame.TextRange.Text;
                            }
                        }
                    }

                text.Select();
                }
            }

        private void button123_Click(object sender, RibbonControlEventArgs e)
            {
            PowerPoint.Selection sel = app.ActiveWindow.Selection;
            PowerPoint.Slide slide = app.ActiveWindow.View.Slide;
            if (sel.Type == PowerPoint.PpSelectionType.ppSelectionNone || sel.TextRange.Count < 2)
                {
                Growl.WarningGlobal("请选中至少2个文本框");
                }
            else
                {
                PowerPoint.ShapeRange range = sel.ShapeRange;
                if (sel.HasChildShapeRange)
                    {
                    range = sel.ChildShapeRange;
                    }
                PowerPoint.Shape text = slide.Shapes.AddTextbox(Office.MsoTextOrientation.msoTextOrientationHorizontal, sel.ShapeRange[sel.ShapeRange.Count].Left + sel.ShapeRange[sel.ShapeRange.Count].Width, sel.ShapeRange[1].Top, sel.ShapeRange[1].Width * sel.ShapeRange.Count / 2, sel.ShapeRange[1].Height);
                PowerPoint.TextFrame2 tframe = text.TextFrame2;
                int count = sel.ShapeRange.Count;
                for (int i = 1 ; i <= count ; i++)
                    {
                    tframe.TextRange.Text = tframe.TextRange.Text + range[i].TextFrame2.TextRange.Text;
                    tframe.AutoSize = Office.MsoAutoSize.msoAutoSizeShapeToFitText;
                    }
                }
            }

        private void button50_Click(object sender, RibbonControlEventArgs e)
            {
            Wpf_crossPage wpf_CrossPage = new Wpf_crossPage();
            wpf_CrossPage.Show();
            }

        private void button8_Click_1(object sender, RibbonControlEventArgs e)
            {
            Globals.ThisAddIn.Application.CommandBars.ExecuteMso("MSPPTInsertTableofContents");//插入缩放
            }

        //发光质感功能开始
        private void button9_Click_1(object sender, RibbonControlEventArgs e)
            {
            MyFunction F = new MyFunction();
            Selection selection = this.app.ActiveWindow.Selection;
            int num1 = Properties.Settings.Default.GradeColorN;
            if ((System.Windows.Forms.Control.ModifierKeys & Keys.Control) == Keys.Control)  //判断Ctrl键
                {
                Wpf_ShapeGlow wpf_ShapeGlow = new Wpf_ShapeGlow();
                wpf_ShapeGlow.ShowDialog();
                }
            else if (selection.Type != PpSelectionType.ppSelectionShapes)
                {
                Growl.WarningGlobal("请选择形状！");
                }
            else
                {
                int GlowNum = Properties.Settings.Default.GlowNum;
                double num2 = Properties.Settings.Default.GlowTra;
                double GlowTra = Math.Round(num2 / 100, 1);
                Color color = Properties.Settings.Default.GlowColor;

                foreach (PowerPoint.Shape shp in selection.ShapeRange)
                    {
                    shp.Glow.Color.RGB = F.RGB2Int(color.R, color.G, color.B);
                    shp.Glow.Transparency = (float)GlowTra;
                    shp.Glow.Radius = GlowNum;
                    }
                }
            }

        private void splitButton8_Click(object sender, RibbonControlEventArgs e)
            {
            MyFunction F = new MyFunction();
            Selection selection = this.app.ActiveWindow.Selection;
            int num1 = Properties.Settings.Default.GradeColorN;
            if ((System.Windows.Forms.Control.ModifierKeys & Keys.Control) == Keys.Control)  //判断Ctrl键
                {
                Wpf_Gradient wpf_Gradient = new Wpf_Gradient();
                wpf_Gradient.ShowDialog();
                }
            else
                {
                if (selection.Type == PpSelectionType.ppSelectionNone)
                    {
                    Growl.WarningGlobal("请选择内容后再试！");
                    return;
                    }
                PowerPoint.ShapeRange shapeRange = selection.ShapeRange;
                if (selection.HasChildShapeRange)
                    {
                    shapeRange = selection.ChildShapeRange;
                    }
                int num = num1;//设置渐变色差
                foreach (object obj in shapeRange)
                    {
                    PowerPoint.Shape shape = (PowerPoint.Shape)obj;
                    if (shape.Type == MsoShapeType.msoGroup)
                        {
                        IEnumerator enumerator2 = shape.GroupItems.GetEnumerator();
                            {
                            while (enumerator2.MoveNext())
                                {
                                object obj2 = enumerator2.Current;
                                PowerPoint.Shape shape2 = (PowerPoint.Shape)obj2;
                                if (shape2.Fill.Type == MsoFillType.msoFillSolid)
                                    {
                                    int rgb = shape2.Fill.ForeColor.RGB;
                                    int r = rgb % 256;
                                    int g = rgb / 256 % 256;
                                    int b = rgb / 256 / 256 % 256;
                                    int num2 = F.Rgb2Hsl(r, g, b);
                                    int num3 = num2 % 256;
                                    int s = num2 / 256 % 256;
                                    int l = num2 / 256 / 256 % 256;
                                    int num4 = num3 + num;
                                    if (num4 > 255)
                                        {
                                        num4 -= 256;
                                        }
                                    else if (num4 < 0)
                                        {
                                        num4 = 256 - num4;
                                        }
                                    int rgb2 = F.Hsl2Rgb(num4, s, l);
                                    shape2.Fill.OneColorGradient(MsoGradientStyle.msoGradientDiagonalUp, 1, 1f);
                                    shape2.Fill.GradientStops[1].Color.RGB = rgb;
                                    shape2.Fill.GradientStops[2].Color.RGB = rgb2;
                                    shape2.Fill.GradientAngle = 0f;
                                    }
                                else if (shape2.Fill.Type == MsoFillType.msoFillGradient)
                                    {
                                    int rgb3 = shape2.Fill.GradientStops[1].Color.RGB;
                                    int r2 = rgb3 % 256;
                                    int g2 = rgb3 / 256 % 256;
                                    int b2 = rgb3 / 256 / 256 % 256;
                                    int num5 = F.Rgb2Hsl(r2, g2, b2);
                                    int num6 = num5 % 256;
                                    int s2 = num5 / 256 % 256;
                                    int l2 = num5 / 256 / 256 % 256;
                                    int num7 = num6 + num;
                                    if (num7 > 255)
                                        {
                                        num7 -= 256;
                                        }
                                    else if (num7 < 0)
                                        {
                                        num7 = 256 - num7;
                                        }
                                    int rgb4 = F.Hsl2Rgb(num7, s2, l2);
                                    shape2.Fill.GradientStops[2].Color.RGB = rgb4;
                                    }
                                else
                                    {
                                    Growl.WarningGlobal("所选非渐变！");
                                    }
                                }
                            continue;
                            }
                        }
                    if (shape.Fill.Type == MsoFillType.msoFillSolid)
                        {
                        int rgb5 = shape.Fill.ForeColor.RGB;
                        int r3 = rgb5 % 256;
                        int g3 = rgb5 / 256 % 256;
                        int b3 = rgb5 / 256 / 256 % 256;
                        int num8 = F.Rgb2Hsl(r3, g3, b3);
                        int num9 = num8 % 256;
                        int s3 = num8 / 256 % 256;
                        int l3 = num8 / 256 / 256 % 256;
                        int num10 = num9 + num;
                        if (num10 > 255)
                            {
                            num10 -= 256;
                            }
                        else if (num10 < 0)
                            {
                            num10 = 256 - num10;
                            }
                        int rgb6 = F.Hsl2Rgb(num10, s3, l3);
                        shape.Fill.OneColorGradient(MsoGradientStyle.msoGradientDiagonalUp, 1, 1f);
                        shape.Fill.GradientStops[1].Color.RGB = rgb5;
                        shape.Fill.GradientStops[2].Color.RGB = rgb6;
                        shape.Fill.GradientAngle = 0f;
                        }
                    else if (shape.Fill.Type == MsoFillType.msoFillGradient)
                        {
                        int rgb7 = shape.Fill.GradientStops[1].Color.RGB;
                        int r4 = rgb7 % 256;
                        int g4 = rgb7 / 256 % 256;
                        int b4 = rgb7 / 256 / 256 % 256;
                        int num11 = F.Rgb2Hsl(r4, g4, b4);
                        int num12 = num11 % 256;
                        int s4 = num11 / 256 % 256;
                        int l4 = num11 / 256 / 256 % 256;
                        int num13 = num12 + num;
                        if (num13 > 255)
                            {
                            num13 -= 256;
                            }
                        else if (num13 < 0)
                            {
                            num13 = 256 - num13;
                            }
                        int rgb8 = F.Hsl2Rgb(num13, s4, l4);
                        shape.Fill.GradientStops[2].Color.RGB = rgb8;
                        }
                    else
                        {
                        Growl.WarningGlobal("请选择内容后再试！");
                        }
                    }
                }
            }

        private void button10_Click_1(object sender, RibbonControlEventArgs e)
            {
            Selection sel = app.ActiveWindow.Selection;
            double num2 = Properties.Settings.Default.ShodwTra;
            double ShadowTra = Math.Round(num2 / 100, 1);
            float ShadowSize = Properties.Settings.Default.ShodwSize;
            float ShadowBlur = Properties.Settings.Default.ShodwBlur;
            float ShadowX = Properties.Settings.Default.ShodwX;
            Color ShadowColor = Properties.Settings.Default.ShodwColor;
            MsoTriState msoTriState = Properties.Settings.Default.ShodwCheck;
            MyFunction F = new MyFunction();
            try
                {
                int num1 = Properties.Settings.Default.GradeColorN;
                if ((System.Windows.Forms.Control.ModifierKeys & Keys.Control) == Keys.Control)  //判断Ctrl键
                    {
                    Wpf_ShapeShodw wpf_ShapeShodw = new Wpf_ShapeShodw();
                    wpf_ShapeShodw.ShowDialog();
                    }
                else
                    {
                    if (sel.Type != PpSelectionType.ppSelectionShapes)
                        {
                        Growl.WarningGlobal("请选择内容后再试！");
                        }
                    else
                        {
                        int count = sel.ShapeRange.Count;
                        for (int i = 0 ; i <= count ; count--)
                            {
                            PowerPoint.Shape shp = sel.ShapeRange[count];
                            shp.Line.Visible = msoTriState;//无边框
                            shp.Shadow.ForeColor.RGB = F.RGB2Int(ShadowColor.R, ShadowColor.G, ShadowColor.B);
                            shp.Shadow.Transparency = (float)ShadowTra;//透明度
                            shp.Shadow.Size = ShadowSize;//大小
                            shp.Shadow.Blur = ShadowBlur;//虚化半径
                            shp.Shadow.OffsetX = ShadowX;//偏移量
                            shp.Shadow.OffsetY = 1.5f;
                            shp.Shadow.RotateWithShape = MsoTriState.msoTrue;//旋转角度
                            shp.Shadow.IncrementOffsetX(0);
                            shp.Shadow.IncrementOffsetY(0);
                            }
                        }
                    }
                }
            catch
                {
                int count = sel.ShapeRange.Count;
                for (int i = 0 ; i <= count ; count--)
                    {
                    PowerPoint.Shape shp = sel.ShapeRange[count];
                    shp.Line.Visible = msoTriState;//无边框
                    shp.Shadow.ForeColor.RGB = F.RGB2Int(ShadowColor.R, ShadowColor.G, ShadowColor.B);
                    shp.Shadow.Transparency = (float)ShadowTra;//透明度
                    shp.Shadow.Size = ShadowSize;//大小
                    shp.Shadow.Blur = ShadowBlur;//虚化半径
                    shp.Shadow.OffsetX = ShadowX;//偏移量
                    shp.Shadow.OffsetY = 1.5f;
                    shp.Shadow.RotateWithShape = MsoTriState.msoTrue;//旋转角度
                    shp.Shadow.IncrementOffsetX(0);
                    shp.Shadow.IncrementOffsetY(0);
                    }
                }
            }

        private void button86_Click(object sender, RibbonControlEventArgs e)
            {
            Wpf_ImageExport wpf_ImageExport = new Wpf_ImageExport();
            wpf_ImageExport.Show();
            }

        private void button52_Click_1(object sender, RibbonControlEventArgs e)
            {
            DelShpe("随机色块");
            Selection sel = app.ActiveWindow.Selection;
            if (sel.Type != PpSelectionType.ppSelectionShapes)
                {
                Growl.Warning("请选择形状后再试！", "温馨提示");
                return;
                }

            MyFunction F = new MyFunction();
            int oldColor = sel.ShapeRange.Fill.ForeColor.RGB;
            Color baseColor = F.Int2RGB(oldColor);
            System.Drawing.Color[] colors = new System.Drawing.Color[6];
            Random ro = new Random();
            int colorDiff = ro.Next(10, 60);

            // 生成颜色数组
            colors[0] = Color.FromArgb(baseColor.A, Math.Max(0, baseColor.R - colorDiff), Math.Min(baseColor.G + colorDiff, 255), Math.Max(0, baseColor.B - colorDiff));
            colors[1] = Color.FromArgb(baseColor.A, Math.Max(0, baseColor.R - colorDiff), Math.Max(0, Math.Min(baseColor.G - colorDiff, 255)), Math.Max(0, Math.Min(baseColor.B - colorDiff, 255)));
            colors[2] = Color.FromArgb(baseColor.A, Math.Max(0, baseColor.R - colorDiff), Math.Max(0, baseColor.G - colorDiff), Math.Min(Math.Max(0, baseColor.B - colorDiff), 255));
            colors[3] = Color.FromArgb(Math.Max(0, baseColor.A - 15), Math.Max(0, Math.Min(baseColor.R - colorDiff, 255)), Math.Max((int)0, (int)baseColor.G), Math.Min(Math.Max(0, baseColor.B - colorDiff), 255));
            colors[4] = Color.FromArgb(Math.Max(0, baseColor.A - 25), Math.Min(255, Math.Max(baseColor.R + colorDiff, 0)), Math.Max((int)0, (int)baseColor.G), Math.Min(Math.Max(0, baseColor.B - colorDiff), 255));
            colors[5] = Color.FromArgb(Math.Max(0, baseColor.A - 5), Math.Max((int)0, (int)baseColor.R), Math.Max(0, Math.Min(baseColor.G - colorDiff, 255)), Math.Max(0, Math.Min(baseColor.B - colorDiff, 255)));

            // 生成形状并设置属性
            PowerPoint.ShapeRange shp0 = sel.ShapeRange;
            shp0.Copy();
            float toTop = shp0.Top;

            for (int i = 1 ; i < 7 ; i++)
                {
                PowerPoint.ShapeRange shape = shp0.Duplicate();
                shape.Tags.Add("配色", "随机色块");
                shape.Top = toTop;
                shape.Left = shp0.Left + shp0.Width * i;
                // 设置填充颜色
                int r = Math.Min((int)colors[i - 1].R, 255);
                int g = Math.Min((int)colors[i - 1].G, 255);
                int b = Math.Min((int)colors[i - 1].B, 255);
                Color newColor = Color.FromArgb(baseColor.A, r, g, b);
                shape.Fill.ForeColor.RGB = F.RGB2Int(newColor.R, newColor.G, newColor.B);

                // 设置文本框颜色和文字内容
                shape.TextFrame.TextRange.Font.Color.RGB = shape.Fill.ForeColor.RGB;
                shape.TextFrame.TextRange.Text = F.RGB2Int(r, g, b).ToString();
                }
            }

        private void button124_Click(object sender, RibbonControlEventArgs e)
            {
            }

        private void button125_Click(object sender, RibbonControlEventArgs e)
            {
            PowerPoint.Selection sel = app.ActiveWindow.Selection;

            if (sel.Type == PpSelectionType.ppSelectionText || sel.Type == PpSelectionType.ppSelectionShapes)
                {
                // 保存初始位置和边距
                float initialTop = sel.ShapeRange.Top;
                float initialLeft = sel.ShapeRange.Left;
                float initialMarginTop = sel.ShapeRange.TextFrame.MarginTop;
                float initialMarginLeft = sel.ShapeRange.TextFrame.MarginLeft;

                // 选择形状范围并设置边距为0
                sel.ShapeRange.Select();
                sel.ShapeRange.TextFrame.MarginLeft = 0;
                sel.ShapeRange.TextFrame.MarginRight = 0;
                sel.ShapeRange.TextFrame.MarginTop = 0;
                sel.ShapeRange.TextFrame.MarginBottom = 0;
                sel.ShapeRange.TextFrame.AutoSize = PpAutoSize.ppAutoSizeShapeToFitText;

                // 恢复字段距离
                float newTop = sel.ShapeRange.Top;
                float newLeft = sel.ShapeRange.Left;
                sel.ShapeRange.Top = newTop + initialMarginTop;
                sel.ShapeRange.Left = newLeft + initialMarginLeft;
                }
            else
                {
                Growl.WarningGlobal("请选择文本内容！");
                }
            }

        private void button27_Click_1(object sender, RibbonControlEventArgs e)
            {
            PowerPoint.Selection sel = app.ActiveWindow.Selection;
            if (sel.Type != PowerPoint.PpSelectionType.ppSelectionShapes)
                {
                Growl.SuccessGlobal("同时选中当前页面中与所选形状相同填充颜色的形状！");
                }
            else
                {
                PowerPoint.Slide slide = app.ActiveWindow.View.Slide;
                PowerPoint.ShapeRange range = sel.ShapeRange;
                if (sel.HasChildShapeRange)
                    {
                    range = sel.ChildShapeRange;
                    }
                PowerPoint.Shape shape = range[1];
                if (shape.Fill.Type == Office.MsoFillType.msoFillSolid)
                    {
                    for (int i = 1 ; i <= slide.Shapes.Count ; i++)
                        {
                        PowerPoint.Shape item = slide.Shapes[i];
                        if (item.Name == shape.Name && item.Id != shape.Id)
                            {
                            item.Name = item.Name + "_" + i;
                            }
                        }
                    List<string> list = new List<string>();
                    foreach (PowerPoint.Shape item in slide.Shapes)
                        {
                        if (item.Type == shape.Type && item.Fill.Type == shape.Fill.Type && item.Fill.ForeColor.RGB == shape.Fill.ForeColor.RGB)
                            {
                            list.Add(item.Name);
                            }
                        }
                    slide.Shapes.Range(list.ToArray()).Select();
                    }
                else
                    {
                    Growl.Warning("形状不是纯色填充");
                    }
                }
            }

        private void button63_Click(object sender, RibbonControlEventArgs e)
            {
            PowerPoint.Selection sel = app.ActiveWindow.Selection;
            if (sel.Type != PowerPoint.PpSelectionType.ppSelectionShapes)
                {
                Growl.SuccessGlobal("同时选中当前页面中与所选形状相同线条颜色的形状！");
                }
            else
                {
                PowerPoint.Slide slide = app.ActiveWindow.View.Slide;
                PowerPoint.ShapeRange range = sel.ShapeRange;
                if (sel.HasChildShapeRange)
                    {
                    range = sel.ChildShapeRange;
                    }
                PowerPoint.Shape shape = range[1];
                if (shape.Line.Visible == Office.MsoTriState.msoTrue)
                    {
                    for (int i = 1 ; i <= slide.Shapes.Count ; i++)
                        {
                        PowerPoint.Shape item = slide.Shapes[i];
                        if (item.Name == shape.Name && item.Id != shape.Id)
                            {
                            item.Name = item.Name + "_" + i;
                            }
                        }
                    List<string> list = new List<string>();
                    foreach (PowerPoint.Shape item in slide.Shapes)
                        {
                        if (item.Type == shape.Type && item.Line.Visible == Office.MsoTriState.msoTrue && item.Line.ForeColor.RGB == shape.Line.ForeColor.RGB)
                            {
                            list.Add(item.Name);
                            }
                        }
                    slide.Shapes.Range(list.ToArray()).Select();
                    }
                else
                    {
                    Growl.Warning("形状无线条");
                    }
                }
            }

        private void button64_Click(object sender, RibbonControlEventArgs e)
            {
            PowerPoint.Selection sel = app.ActiveWindow.Selection;
            if (sel.Type != PowerPoint.PpSelectionType.ppSelectionShapes)
                {
                Growl.SuccessGlobal("同时选中当前页面中与所选形状相同线条粗细的形状！");
                }
            else
                {
                PowerPoint.Slide slide = app.ActiveWindow.View.Slide;
                PowerPoint.ShapeRange range = sel.ShapeRange;
                if (sel.HasChildShapeRange)
                    {
                    range = sel.ChildShapeRange;
                    }
                PowerPoint.Shape shape = range[1];
                if (shape.Line.Visible == Office.MsoTriState.msoTrue)
                    {
                    for (int i = 1 ; i <= slide.Shapes.Count ; i++)
                        {
                        PowerPoint.Shape item = slide.Shapes[i];
                        if (item.Name == shape.Name && item.Id != shape.Id)
                            {
                            item.Name = item.Name + "_" + i;
                            }
                        }
                    List<string> list = new List<string>();
                    foreach (PowerPoint.Shape item in slide.Shapes)
                        {
                        if (item.Type == shape.Type && item.Line.Visible == Office.MsoTriState.msoTrue && item.Line.Weight == shape.Line.Weight)
                            {
                            list.Add(item.Name);
                            }
                        }
                    slide.Shapes.Range(list.ToArray()).Select();
                    }
                else
                    {
                    Growl.WarningGlobal("形状无线条！");
                    }
                }
            }

        private void button65_Click(object sender, RibbonControlEventArgs e)
            {
            PowerPoint.Selection sel = app.ActiveWindow.Selection;
            if (sel.Type != PowerPoint.PpSelectionType.ppSelectionShapes)
                {
                Growl.SuccessGlobal("同时选中当前页面中与所选形状相同类型的形状！");
                }
            else
                {
                PowerPoint.Slide slide = app.ActiveWindow.View.Slide;
                PowerPoint.ShapeRange range = sel.ShapeRange;
                if (sel.HasChildShapeRange)
                    {
                    range = sel.ChildShapeRange;
                    }
                PowerPoint.Shape shape = range[1];
                for (int i = 1 ; i <= slide.Shapes.Count ; i++)
                    {
                    PowerPoint.Shape item = slide.Shapes[i];
                    if (item.Name == shape.Name && item.Id != shape.Id)
                        {
                        item.Name = item.Name + "_" + i;
                        }
                    }
                List<string> list = new List<string>();
                foreach (PowerPoint.Shape item in slide.Shapes)
                    {
                    if (item.Type == shape.Type)
                        {
                        list.Add(item.Name);
                        }
                    }
                slide.Shapes.Range(list.ToArray()).Select();
                }
            }

        private void button126_Click(object sender, RibbonControlEventArgs e)
            {
            PowerPoint.Selection sel = app.ActiveWindow.Selection;
            if (sel.Type != PowerPoint.PpSelectionType.ppSelectionShapes)
                {
                Growl.SuccessGlobal("同时选中当前页面中与所选形状相同类型的形状！");
                }
            else
                {
                PowerPoint.Slide slide = app.ActiveWindow.View.Slide;
                PowerPoint.ShapeRange range = sel.ShapeRange;
                if (sel.HasChildShapeRange)
                    {
                    range = sel.ChildShapeRange;
                    }
                PowerPoint.Shape shape = range[1];
                for (int i = 1 ; i <= slide.Shapes.Count ; i++)
                    {
                    PowerPoint.Shape item = slide.Shapes[i];
                    if (item.Name == shape.Name && item.Id != shape.Id)
                        {
                        item.Name = item.Name + "_" + i;
                        }
                    }
                List<string> list = new List<string>();
                foreach (PowerPoint.Shape item in slide.Shapes)
                    {
                    if (item.Type == Office.MsoShapeType.msoAutoShape)
                        {
                        if (item.AutoShapeType == shape.AutoShapeType)
                            {
                            list.Add(item.Name);
                            }
                        }
                    else if (item.Type == Office.MsoShapeType.msoFreeform)
                        {
                        if (item.Type == Office.MsoShapeType.msoFreeform)
                            {
                            list.Add(item.Name);
                            }
                        }
                    }
                slide.Shapes.Range(list.ToArray()).Select();
                }
            }

        private void button128_Click(object sender, RibbonControlEventArgs e)
            {
            PowerPoint.Selection sel = app.ActiveWindow.Selection;
            if (sel.Type != PowerPoint.PpSelectionType.ppSelectionShapes)
                {
                Growl.WarningGlobal("同时选中当前页面中与所选形状相同尺寸的形状");
                }
            else
                {
                PowerPoint.Slide slide = app.ActiveWindow.View.Slide;
                PowerPoint.ShapeRange range = sel.ShapeRange;
                if (sel.HasChildShapeRange)
                    {
                    range = sel.ChildShapeRange;
                    }
                PowerPoint.Shape shape = range[1];
                for (int i = 1 ; i <= slide.Shapes.Count ; i++)
                    {
                    PowerPoint.Shape item = slide.Shapes[i];
                    if (item.Name == shape.Name && item.Id != shape.Id)
                        {
                        item.Name = item.Name + "_" + i;
                        }
                    }
                List<string> list = new List<string>();
                foreach (PowerPoint.Shape item in slide.Shapes)
                    {
                    if (item.Height == shape.Height && item.Width == shape.Width)
                        {
                        list.Add(item.Name);
                        }
                    }
                slide.Shapes.Range(list.ToArray()).Select();
                }
            }

        private void button129_Click(object sender, RibbonControlEventArgs e)
            {
            Wpf_Mockup wpf_Mockup = new Wpf_Mockup();
            wpf_Mockup.Show();
            }

        private void button130_Click(object sender, RibbonControlEventArgs e)
            {
            Wpf_InfoDelete wpf_InfoDelete = new Wpf_InfoDelete();
            wpf_InfoDelete.ShowDialog();
            }

        private void button131_Click(object sender, RibbonControlEventArgs e)
            {
            saveFileDialog1.Filter = "PPT文件（*.PPTX)|*.PPTX";
            //设置默认文件类型显示顺序（可以不设置）
            saveFileDialog1.FilterIndex = 2;
            //保存对话框是否记忆上次打开的目录
            saveFileDialog1.RestoreDirectory = true;
            //设置默认的文件名
            string Name = app.ActivePresentation.Name;
            saveFileDialog1.FileName = "全图型-" + Name;
            DialogResult dr = saveFileDialog1.ShowDialog();
            string fileName = saveFileDialog1.FileName;
            if (dr == System.Windows.Forms.DialogResult.OK && !string.IsNullOrEmpty(fileName))
                {
                app.Presentations[1].SaveAs(fileName, PpSaveAsFileType.ppSaveAsOpenXMLPicturePresentation, MsoTriState.msoTriStateMixed);//导出PDF
                }
            else
                {
                Growl.WarningGlobal("您未选择文件夹，导出失败！");
                }
            }

        private void button133_Click(object sender, RibbonControlEventArgs e)
            {
            }

        private void group3_DialogLauncherClick(object sender, RibbonControlEventArgs e)
            {
            ////独立窗口
            //Wpf_Clipboard wpf_Clipboard = new Wpf_Clipboard();
            //wpf_Clipboard.Show();
            }

        /// <summary>
        /// 设置文本行距  ,num为
        /// </summary>
        /// <param name="num"></param>
        public void Paragraph(float num)
            {
            PowerPoint.Selection sel = app.ActiveWindow.Selection;
            PowerPoint.Slide slide = app.ActiveWindow.View.Slide;

            if (sel.Type == PowerPoint.PpSelectionType.ppSelectionNone)
                {
                Growl.WarningGlobal("请选中至少1个文本框");
                }
            else
                {
                PowerPoint.ShapeRange range = sel.ShapeRange;
                if (sel.HasChildShapeRange)
                    {
                    range = sel.ChildShapeRange;
                    }
                int count = range.Count;
                for (int i = 1 ; i <= count ; i++)
                    {
                    PowerPoint.Shape shape = range[i];
                    ParagraphFormat paragraphFormat = shape.TextFrame.TextRange.ParagraphFormat;
                    paragraphFormat.LineRuleWithin = MsoTriState.msoTrue;//设置为行数
                    paragraphFormat.SpaceWithin = num;//段落间距
                                                      // paragraphFormat.Alignment = PpParagraphAlignment.ppAlignJustify;//两端对齐
                    }
                }
            }

        private void splitButton9_Click(object sender, RibbonControlEventArgs e)
            {
            //Form_Dialog form_Dialog = new Form_Dialog();
            //double value = Settings.Default.Pa_Spaced;
            //if ((System.Windows.Forms.Control.ModifierKeys & Keys.Control) == Keys.Control)  //判断Ctrl键
            //{
            //    if (form_Dialog.InputDoubleDialog(ref value, 2, true, "请输入设定行距值！", true))
            //    {
            //        Settings.Default.Pa_Spaced = (float)value;
            //        Settings.Default.Save();

            //    }
            //}
            //else
            //{
            //    Paragraph((float)value);
            //}
            }

        private void button133_Click_1(object sender, RibbonControlEventArgs e)
            {
            Paragraph((float)1.2);
            }

        private void button98_Click_3(object sender, RibbonControlEventArgs e)
            {
            Paragraph((float)1);
            }

        private void button134_Click(object sender, RibbonControlEventArgs e)
            {
            Paragraph((float)1.3);
            }

        private void button135_Click(object sender, RibbonControlEventArgs e)
            {
            Paragraph((float)1.5);
            }

        private void button136_Click(object sender, RibbonControlEventArgs e)
            {
            Paragraph((float)2);
            }

        private void button137_Click(object sender, RibbonControlEventArgs e)
            {
            //Form_Dialog form_Dialog = new Form_Dialog();
            //double value = Settings.Default.Pa_Spaced;
            //if (form_Dialog.InputDoubleDialog(ref value, 2, true, "请输入设定行距值！", true))
            //{
            //    Settings.Default.Pa_Spaced = (float)value;
            //    Settings.Default.Save();

            //}
            }

        private void button138_Click(object sender, RibbonControlEventArgs e)
            {
            }

        private void splitButton10_Click(object sender, RibbonControlEventArgs e)
            {
            Selection sel = app.ActiveWindow.Selection;
            if (checkBox2.Checked)
                {
                AutoTextBox(sel);
                }
            else
                {
                AdjustTextBox(sel, 0, 0, 0, 0);
                }
            }

        /// <summary>
        /// 自适应调整文本边框
        /// </summary>
        /// <param name="shr"></param>
        public void AutoTextBox(Selection sel)
            {
            PowerPoint.Slide slide = app.ActiveWindow.View.Slide;

            if (sel.Type == PowerPoint.PpSelectionType.ppSelectionNone)
                {
                Growl.WarningGlobal("请选中至少1个文本框");
                }
            else
                {
                PowerPoint.ShapeRange range = sel.ShapeRange;
                if (sel.HasChildShapeRange)
                    {
                    range = sel.ChildShapeRange;
                    }
                int count = range.Count;
                for (int i = 1 ; i <= count ; i++)
                    {
                    PowerPoint.Shape shape = range[i];
                    shape.TextFrame.AutoSize = PpAutoSize.ppAutoSizeShapeToFitText;
                    }
                }
            }

        /// <summary>
        /// 调整文本边距
        /// </summary>
        /// <param name="sel"></param>
        /// <param name="leftMargin"></param>
        /// <param name="topMargin"></param>
        /// <param name="rightMargin"></param>
        /// <param name="bottomMargin"></param>
        public void AdjustTextBox(Selection sel, float leftMargin, float topMargin, float rightMargin, float bottomMargin)
            {
            PowerPoint.Slide slide = app.ActiveWindow.View.Slide;
            if (sel.Type == PowerPoint.PpSelectionType.ppSelectionNone)
                {
                Growl.WarningGlobal("请选中至少1个文本框");
                }
            else
                {
                PowerPoint.ShapeRange range = sel.ShapeRange;
                if (sel.HasChildShapeRange)
                    {
                    range = sel.ChildShapeRange;
                    }
                int count = range.Count;
                for (int i = 1 ; i <= count ; i++)
                    {
                    PowerPoint.Shape shape = range[i];
                    shape.TextFrame.MarginLeft = leftMargin;
                    shape.TextFrame.MarginTop = topMargin;
                    shape.TextFrame.MarginRight = rightMargin;
                    shape.TextFrame.MarginBottom = bottomMargin;
                    }
                }
            }

        private void splitButton11_Click(object sender, RibbonControlEventArgs e)
            {
            try
                {
                Slide oSld = Globals.ThisAddIn.Application.ActiveWindow.Selection.SlideRange[1];
                foreach (PowerPoint.Shape oShp in Globals.ThisAddIn.Application.ActiveWindow.Selection.ShapeRange)
                    {
                    if (oShp.HasTextFrame == MsoTriState.msoTrue)
                        {
                        oShp.TextFrame.TextRange.Words(1).Text = "文头" + oShp.TextFrame.TextRange.Words(1).Text;
                        float oMargin = oShp.TextFrame.TextRange.Words(1, 2).BoundWidth;
                        oShp.TextFrame.TextRange.Words(1, 2).Delete();

                        for (int i = 1 ; i <= oShp.TextFrame.Ruler.Levels.Count ; i++)
                            {
                            if (oShp.TextFrame.Ruler.Levels[i].FirstMargin == 0)
                                {
                                oShp.TextFrame.Ruler.Levels[i].FirstMargin = oMargin;
                                }
                            else
                                {
                                oShp.TextFrame.Ruler.Levels[i].FirstMargin = 0;
                                }
                            oShp.TextFrame.Ruler.Levels[i].LeftMargin = 0;
                            }
                        }
                    }
                }
            catch (Exception ex)
                {
                Console.WriteLine("Error in IndentFirstLine: " + ex.Message);
                }
            }

        private void button139_Click(object sender, RibbonControlEventArgs e)
            {
            // 自定义Clamp方法
            int Clamp(int value, int min, int max)
                {
                return Math.Max(min, Math.Min(value, max));
                }

            // 获取当前幻灯片和演示文稿
            Slide slide = app.ActiveWindow.View.Slide;
            Presentation presentation = app.ActivePresentation;

            // 获取幻灯片的宽度和高度
            var height = presentation.PageSetup.SlideHeight;
            var width = presentation.PageSetup.SlideWidth;

            // 存储已有标签的顶部位置
            List<float> tops = new List<float>();

            // 遍历当前幻灯片中的形状，删除已有标签并记录位置
            foreach (PowerPoint.Shape item in slide.Shapes)
                {
                if (item.Tags["标签"] == "定稿")
                    {
                    // 删除已有标签
                    // item.Delete();

                    // 记录标签位置
                    tops.Add(item.Top + item.Height);
                    }
                }

            // 对顶部位置进行排序
            tops.Sort();

            // 确定新矩形形状的顶部位置
            float newTopPosition = (tops.Count > 0) ? tops.Last() + 3 : 0; // 默认从0开始

            // 如果有空位，调整位置
            for (float y = 0 ; y < height ; y += 53) // 53是矩形的高度加上间距
                {
                if (!tops.Any(top => y >= top - 50 && y <= top)) // 确保不与已有标签重叠
                    {
                    newTopPosition = y; // 找到空位
                    break;
                    }
                }

            // 初始化颜色的RGB值
            int initR = 192, initG = 62, initB = 28; // 可以根据需要调整初始颜色

            // 随机颜色生成器
            Random random = new Random();
            int rAdjust = random.Next(-15, 30);
            int gAdjust = random.Next(-15, 30);
            int bAdjust = random.Next(-15, 30);

            // 计算新的RGB值并确保其在0-255范围内
            int newR = Clamp(initR + rAdjust, 0, 255);
            int newG = Clamp(initG + gAdjust, 0, 255);
            int newB = Clamp(initB + bAdjust, 0, 255);

            // 添加新的矩形形状
            PowerPoint.Shape shape1 = slide.Shapes.AddShape(MsoAutoShapeType.msoShapeRectangle, width, newTopPosition, 100, 50);

            // 添加标签
            shape1.Tags.Add("标签", "定稿");
            shape1.Line.Visible = MsoTriState.msoFalse;

            // 设置矩形的填充颜色和文本
            MyFunction F = new MyFunction();
            shape1.Fill.ForeColor.RGB = F.RGB2Int(newR, newG, newB);
            shape1.TextFrame.TextRange.Text = "已定稿";
            }

        private void button140_Click(object sender, RibbonControlEventArgs e)
            {
            Selection sel = app.ActiveWindow.Selection;
            Slide slide = app.ActiveWindow.View.Slide;
            MyFunction F = new MyFunction();
            if (sel.Type != PpSelectionType.ppSelectionShapes)
                {
                return;
                }
            else
                {
                int count = sel.ShapeRange.Count;

                for (int i = 0 ; i < count ; count--)
                    {
                    PowerPoint.Shape shp = sel.ShapeRange[count];//选择的图形
                    shp.Flip(MsoFlipCmd.msoFlipVertical);
                    }
                }
            }

        private void button141_Click(object sender, RibbonControlEventArgs e)
            {
            Selection sel = app.ActiveWindow.Selection;
            Slide slide = app.ActiveWindow.View.Slide;
            MyFunction F = new MyFunction();
            if (sel.Type != PpSelectionType.ppSelectionShapes)
                {
                return;
                }
            else
                {
                int count = sel.ShapeRange.Count;

                for (int i = 0 ; i < count ; count--)
                    {
                    PowerPoint.Shape shp = sel.ShapeRange[count];//选择的图形
                    shp.Flip(MsoFlipCmd.msoFlipHorizontal);
                    }
                }
            }

        private void button114_Click_1(object sender, RibbonControlEventArgs e)
            {
            try
                {
                Globals.ThisAddIn.Application.CommandBars.ExecuteMso("ObjectsAlignCenterHorizontalSmart");
                Globals.ThisAddIn.Application.CommandBars.ExecuteMso("ObjectsAlignMiddleVerticalSmart");
                }
            catch
                {
                // Handle exception if needed, or ignore it if it's expected
                }
            }

        public void button142_Click(object sender, RibbonControlEventArgs e)
            {
            PowerPoint.Selection selection = app.ActiveWindow.Selection;
            PowerPoint.Presentation presentation = app.ActivePresentation;

            if (selection.Type == PowerPoint.PpSelectionType.ppSelectionShapes && selection.ShapeRange.Count > 0)
                {
                float slideWidth = presentation.PageSetup.SlideWidth;
                float slideHeight = presentation.PageSetup.SlideHeight;
                bool isSingleShape = selection.ShapeRange.Count == 1;

                foreach (PowerPoint.Shape shape in selection.ShapeRange)
                    {
                    if (isSingleShape)
                        {
                        shape.Left = 0;
                        shape.Top = 0;
                        shape.Width = slideWidth;
                        shape.Height = slideHeight;
                        }
                    else
                        {
                        shape.Width = selection.ShapeRange[1].Width;
                        shape.Height = selection.ShapeRange[1].Height;
                        }
                    }
                }
            }

        private void button143_Click(object sender, RibbonControlEventArgs e)
            {
            PowerPoint.Selection selection = app.ActiveWindow.Selection;
            PowerPoint.Presentation presentation = app.ActivePresentation;

            if (selection.Type == PowerPoint.PpSelectionType.ppSelectionShapes && selection.ShapeRange.Count > 0)
                {
                float slideHeight = presentation.PageSetup.SlideHeight;
                bool isSingleShape = selection.ShapeRange.Count == 1;

                foreach (PowerPoint.Shape shape in selection.ShapeRange)
                    {
                    if (isSingleShape)
                        {
                        shape.Height = slideHeight;
                        shape.Top = 0;
                        shape.Left = shape.Left; // 保持左边距离不变
                        }
                    else
                        {
                        float shapeHeight = selection.ShapeRange[1].Height;
                        float shapeTop = selection.ShapeRange[1].Top;
                        float shapeLeft = selection.ShapeRange[1].Left;
                        shape.Height = shapeHeight;
                        }
                    }
                }
            }

        private void button146_Click(object sender, RibbonControlEventArgs e)
            {
            Selection selection = app.ActiveWindow.Selection;
            Presentation pre = app.ActivePresentation;
            if (selection.Type == PpSelectionType.ppSelectionShapes)
                {
                PowerPoint.Shape firstShape = selection.ShapeRange[1];
                float width = firstShape.Width;
                float PageWidth = pre.PageSetup.SlideWidth;
                if (selection.ShapeRange.Count == 1)
                    {
                    PowerPoint.Shape shape = selection.ShapeRange[1];
                    shape.Width = PageWidth;
                    shape.Left = 0;
                    }
                else
                    {
                    foreach (PowerPoint.Shape shape in selection.ShapeRange)
                        {
                        shape.Width = width;
                        }
                    }
                }
            }

        private void button144_Click(object sender, RibbonControlEventArgs e)
            {
            }

        /// <summary>
        /// 判断形状是否在周围
        /// </summary>
        /// <param name="selected"></param>
        /// <param name="shape"></param>
        /// <returns></returns>
        private bool IsShapeAroundSelected(PowerPoint.Shape selected, PowerPoint.Shape shape)
            {
            float selectedLeft = selected.Left;
            float selectedTop = selected.Top;
            float selectedWidth = selected.Width;
            float selectedHeight = selected.Height;

            float shapeLeft = shape.Left;
            float shapeTop = shape.Top;
            float shapeWidth = shape.Width;
            float shapeHeight = shape.Height;

            return ((shapeLeft >= selectedLeft - shapeWidth && shapeLeft <= selectedLeft + selectedWidth) &&
                (shapeTop >= selectedTop - shapeHeight && shapeTop <= selectedTop + selectedHeight)
            );
            }

        private void button145_Click(object sender, RibbonControlEventArgs e)
            {
            PowerPoint.Selection selection = app.ActiveWindow.Selection;
            PowerPoint.ShapeRange shapeRange = selection.ShapeRange;
            int count = shapeRange.Count;

            if (count < 2)
                {
                // 选择的形状数量小于2，无法进行操作
                return;
                }

            // 获取最后一个形状
            PowerPoint.Shape lastShape = shapeRange[count];

            // 获取最后一个形状格式
            lastShape.PickUp();

            if (lastShape.AnimationSettings.Animate == MsoTriState.msoTrue)
                {
                lastShape.PickupAnimation();
                }

            // 应用格式和动画到其余形状
            for (int i = 1 ; i < count ; i++)
                {
                PowerPoint.Shape shape = shapeRange[i];
                shape.Apply();
                if (lastShape.AnimationSettings.Animate == MsoTriState.msoTrue)
                    {
                    shape.ApplyAnimation();
                    }
                }

            // 清除剪贴板中的内容
            System.Windows.Clipboard.Clear();
            }

        private void button147_Click(object sender, RibbonControlEventArgs e)
            {
            app.CommandBars.ExecuteMso("AutoShapeInsert");
            }

        private void button66_Click(object sender, RibbonControlEventArgs e)
            {
            Slide slide = app.ActiveWindow.View.Slide;
            PowerPoint.Selection selection = app.ActiveWindow.Selection;
            List<string> Shps = new List<string>();
            if (selection.Type == PpSelectionType.ppSelectionShapes || selection.Type == PpSelectionType.ppSelectionText)
                {
                PowerPoint.Shape firstshp = selection.ShapeRange[1];
                if (firstshp.HasTextFrame == MsoTriState.msoTrue)
                    {
                    float FontSize = firstshp.TextFrame.TextRange.Font.Size;
                    int count = slide.Shapes.Count;
                    for (int i = 0 ; i < count ; count--)
                        {
                        PowerPoint.Shape shp = slide.Shapes[count];
                        if (shp.HasTextFrame == MsoTriState.msoTrue)
                            {
                            if (shp.TextFrame.TextRange.Font.Size == FontSize)
                                {
                                Shps.Add(shp.Name);
                                }
                            }
                        }

                    selection.Unselect();
                    slide.Shapes.Range(Shps.ToArray()).Select(MsoTriState.msoTrue);
                    }
                }
            else
                {
                Growl.WarningGlobal("请选择内容后再试");
                }
            }

        private void button67_Click(object sender, RibbonControlEventArgs e)
            {
            Slide slide = app.ActiveWindow.View.Slide;
            Selection selection = app.ActiveWindow.Selection;
            List<string> Shps = new List<string>();
            if (selection.Type == PpSelectionType.ppSelectionShapes || selection.Type == PpSelectionType.ppSelectionText)
                {
                PowerPoint.Shape firstshp = selection.ShapeRange[1];
                if (firstshp.HasTextFrame == MsoTriState.msoTrue)
                    {
                    int FontColor = firstshp.TextFrame.TextRange.Font.Color.RGB;
                    int count = slide.Shapes.Count;
                    for (int i = 0 ; i < count ; count--)
                        {
                        PowerPoint.Shape shp = slide.Shapes[count];
                        if (shp.HasTextFrame == MsoTriState.msoTrue)
                            {
                            if (shp.TextFrame.TextRange.Font.Color.RGB == FontColor)
                                {
                                Shps.Add(shp.Name);
                                }
                            }
                        }

                    selection.Unselect();
                    slide.Shapes.Range(Shps.ToArray()).Select(MsoTriState.msoTrue);
                    }
                }
            else
                {
                Growl.WarningGlobal("请选择内容后再试");
                }
            }

        private void button68_Click(object sender, RibbonControlEventArgs e)
            {
            Slide slide = app.ActiveWindow.View.Slide;
            Selection selection = app.ActiveWindow.Selection;
            List<string> Shps = new List<string>();
            if (selection.Type == PpSelectionType.ppSelectionShapes || selection.Type == PpSelectionType.ppSelectionText)
                {
                PowerPoint.Shape firstshp = selection.ShapeRange[1];
                if (firstshp.HasTextFrame == MsoTriState.msoTrue)
                    {
                    string FontName = firstshp.TextFrame.TextRange.Font.Name;
                    int count = slide.Shapes.Count;
                    for (int i = 0 ; i < count ; count--)
                        {
                        PowerPoint.Shape shp = slide.Shapes[count];
                        if (shp.HasTextFrame == MsoTriState.msoTrue)
                            {
                            if (shp.TextFrame.TextRange.Font.Name == FontName)
                                {
                                Shps.Add(shp.Name);
                                }
                            }
                        }

                    selection.Unselect();
                    slide.Shapes.Range(Shps.ToArray()).Select(MsoTriState.msoTrue);
                    }
                }
            else
                {
                Growl.WarningGlobal("请选择内容后再试");
                }
            }

        private void checkBox2_Click(object sender, RibbonControlEventArgs e)
            {
            Settings.Default.TextBoxAuto = checkBox2.Checked;
            Settings.Default.Save();
            if (checkBox2.Checked)
                {
                checkBox2.Label = "自适应模式";
                }
            else
                {
                checkBox2.Label = "无边框模式";
                }
            }

        private void button148_Click(object sender, RibbonControlEventArgs e)
            {
            Slide slide = app.ActiveWindow.View.Slide;
            Selection sel = app.ActiveWindow.Selection;
            if (sel.Type == PpSelectionType.ppSelectionSlides)
                {
                foreach (Slide item in sel.SlideRange)
                    {
                    foreach (PowerPoint.Shape shp in item.Shapes)
                        {
                        if (checkBox2.Checked)
                            {
                            shp.TextFrame.AutoSize = PpAutoSize.ppAutoSizeShapeToFitText;
                            }
                        else
                            {
                            shp.TextFrame.MarginLeft = 0;
                            shp.TextFrame.MarginTop = 0;
                            shp.TextFrame.MarginRight = 0;
                            shp.TextFrame.MarginBottom = 0;
                            }
                        }
                    }
                }
            else
                {
                Growl.WarningGlobal("请先选择幻灯片！");
                }
            }

        private void button149_Click(object sender, RibbonControlEventArgs e)
            {
            Slide slide = app.ActiveWindow.View.Slide;
            int count = slide.Shapes.Count;
            for (int i = 0 ; i < count ; count--)
                {
                PowerPoint.Shape shp = slide.Shapes[count];
                if (shp.Type == MsoShapeType.msoTextBox)
                    {
                    if (checkBox2.Checked)
                        {
                        shp.TextFrame.AutoSize = PpAutoSize.ppAutoSizeShapeToFitText;
                        }
                    else
                        {
                        shp.TextFrame.MarginLeft = 0;
                        shp.TextFrame.MarginTop = 0;
                        shp.TextFrame.MarginRight = 0;
                        shp.TextFrame.MarginBottom = 0;
                        }
                    }
                }
            }

        private void button150_Click(object sender, RibbonControlEventArgs e)
            {
            Slide slide = app.ActiveWindow.View.Slide;
            Presentation pre = app.ActivePresentation;
            foreach (Slide item in pre.Slides)
                {
                foreach (PowerPoint.Shape shp in item.Shapes)
                    {
                    if (checkBox2.Checked)
                        {
                        shp.TextFrame.AutoSize = PpAutoSize.ppAutoSizeShapeToFitText;
                        }
                    else
                        {
                        shp.TextFrame.MarginLeft = 0;
                        shp.TextFrame.MarginTop = 0;
                        shp.TextFrame.MarginRight = 0;
                        shp.TextFrame.MarginBottom = 0;
                        }
                    }
                }

            int count = pre.Slides.Count;
            string Text = "已处理" + count + "页幻灯片的文本内容";
            Growl.WarningGlobal(Text);
            }

        private void button107_Click_1(object sender, RibbonControlEventArgs e)
            {
            Wpf_UniPa wpf_UniPa = new Wpf_UniPa();
            wpf_UniPa.Show();
            }

        private void button151_Click(object sender, RibbonControlEventArgs e)
            {
            string svgData = System.Windows.Forms.Clipboard.GetText();
            if (!svgData.Contains("<svg"))
                {
                Growl.WarningGlobal("剪贴板不含svg数据");

                return;
                }

            string svgPath = System.IO.Path.GetTempPath() + "temp.svg";
            File.WriteAllText(svgPath, svgData);
            Presentation presentation = app.ActivePresentation;
            Slide slide = app.ActiveWindow.Selection.SlideRange[1];
            PowerPoint.Shape shapePic = slide.Shapes.AddPicture(svgPath, MsoTriState.msoFalse, MsoTriState.msoTrue, 0, 0);

            if (app.ActiveWindow.Selection.Type == PpSelectionType.ppSelectionShapes || app.ActiveWindow.Selection.Type == PpSelectionType.ppSelectionText)
                {
                PowerPoint.Shape selectedShape = app.ActiveWindow.Selection.ShapeRange[1];
                shapePic.LockAspectRatio = MsoTriState.msoTrue;
                shapePic.Width = selectedShape.Width;
                shapePic.Left = selectedShape.Left;
                shapePic.Top = selectedShape.Top - shapePic.Height;
                }
            shapePic.Select();
            if ((System.Windows.Forms.Control.ModifierKeys & Keys.Control) == Keys.Control)  //判断Ctrl键
                {
                app.CommandBars.ExecuteMso("SVGEdit"); // 转化为形状
                }
            else
                {
                return;
                }
            }

        private void button152_Click(object sender, RibbonControlEventArgs e)
            {
            PowerPoint.Application app = Globals.ThisAddIn.Application;
            PowerPoint.Slide slide = app.ActiveWindow.Selection.SlideRange[1];
            PowerPoint.Shape selectedShape = app.ActiveWindow.Selection.ShapeRange[1];
            List<PowerPoint.Shape> surroundingShapes = new List<PowerPoint.Shape>();

            // Get all shapes around the selected shape
            foreach (PowerPoint.Shape shape in slide.Shapes)
                {
                if (shape.Id != selectedShape.Id && IsShapeAroundSelected(selectedShape, shape))
                    {
                    surroundingShapes.Add(shape);
                    }
                }
            PowerPoint.Shape groupShape = slide.Shapes.Range(new Object[] { selectedShape.Name }.Concat(surroundingShapes.Select(x => x.Name)).ToArray()).Group();
            groupShape.Select();
            }

        private void button53_Click_1(object sender, RibbonControlEventArgs e)
            {
            Selection sel = app.ActiveWindow.Selection;
            if (sel.ShapeRange.Count == 1)
                {
                PowerPoint.Shape shape = sel.ShapeRange[1];
                // 获取形状的坐标和大小
                float x = shape.Left;
                float y = shape.Top;
                float width = shape.Width;
                float height = shape.Height;

                // 创建临时文件用于保存 SVG
                string tempFileName = System.IO.Path.GetTempFileName();
                string svgFileName = System.IO.Path.ChangeExtension(tempFileName, ".wmf");
                // 将形状转换为 SVG 文件
                shape.Export(svgFileName, PpShapeFormat.ppShapeFormatWMF, (int)width, (int)height, PpExportMode.ppRelativeToSlide);
                //将图片添加回PPT
                app.ActiveWindow.View.Slide.Shapes.AddPicture(svgFileName, MsoTriState.msoFalse, MsoTriState.msoTrue, x, y);
                // 删除临时文件
                shape.Delete();
                File.Delete(tempFileName);
                File.Delete(svgFileName);
                }
            else
                {
                Growl.WarningGlobal("请选择单个文件");
                }
            }

        private void button61_Click(object sender, RibbonControlEventArgs e)
            {
            Selection sel = app.ActiveWindow.Selection;
            Slide slide = app.ActiveWindow.View.Slide;
            List<string> shps = new List<string> { };
            MyFunction F = new MyFunction();
            if (sel.Type != PpSelectionType.ppSelectionShapes)
                {
                Growl.WarningGlobal("请选择内容后再试！");
                }
            else
                {
                int count = sel.ShapeRange.Count;
                for (int i = 0 ; i < count ; count--)
                    {
                    PowerPoint.Shape shp = sel.ShapeRange[count];//选择的图形
                    float oldTop = shp.Top;
                    float oldLeft = shp.Left;
                    float oldWidth = shp.Width;
                    float oldHeight = shp.Height;
                    float ceNTop = oldTop + oldHeight;
                    float ceNLeft = oldLeft + oldWidth;
                    //添加角标
                    PowerPoint.Shape newShp = slide.Shapes.AddShape(MsoAutoShapeType.msoShapeRectangle, ceNLeft - oldHeight / 4, ceNTop - oldHeight / 4, oldHeight / 2, oldHeight / 2);
                    newShp.ZOrder(MsoZOrderCmd.msoSendToBack);
                    newShp.Line.Visible = MsoTriState.msoFalse;
                    shps.Add(newShp.Name);
                    }
                }
            sel.Unselect();
            slide.Shapes.Range(shps.ToArray()).Select(MsoTriState.msoTrue);
            }

        private void splitButton12_Click(object sender, RibbonControlEventArgs e)
            {
            }

        private void button147_Click_1(object sender, RibbonControlEventArgs e)
            {
            // 获取选择的形状
            var selectedShapes = Globals.ThisAddIn.Application.ActiveWindow.Selection.ShapeRange;

            if (selectedShapes.Count < 2)
                {
                // 如果选择的形状数量小于2，则无法进行左对齐，直接返回
                return;
                }

            // 获取最后一个形状
            var lastShape = selectedShapes[selectedShapes.Count];

            // 获取最后一个形状的左侧位置
            var lastShapeLeft = lastShape.Left;

            // 遍历选择的形状，将它们的左侧位置与最后一个形状的左侧位置对齐
            for (int i = 1 ; i <= selectedShapes.Count - 1 ; i++)
                {
                var shape = selectedShapes[i];
                shape.Left = lastShapeLeft;
                }
            }

        private void button153_Click(object sender, RibbonControlEventArgs e)
            {
            // 获取选择的形状
            var selectedShapes = Globals.ThisAddIn.Application.ActiveWindow.Selection.ShapeRange;

            if (selectedShapes.Count < 2)
                {
                // 如果选择的形状数量小于2，则无法进行左对齐，直接返回
                return;
                }

            // 获取最后一个形状
            var lastShape = selectedShapes[selectedShapes.Count];

            // 获取最后一个形状的右侧位置
            var lastShapeLeft = lastShape.Left + lastShape.Width;

            // 遍历选择的形状，将它们的左侧位置与最后一个形状的左侧位置对齐
            for (int i = 1 ; i <= selectedShapes.Count - 1 ; i++)
                {
                var shape = selectedShapes[i];
                shape.Left = lastShapeLeft - shape.Width;
                }
            }

        private void button154_Click(object sender, RibbonControlEventArgs e)
            {       // 获取选择的形状
            var selectedShapes = Globals.ThisAddIn.Application.ActiveWindow.Selection.ShapeRange;

            if (selectedShapes.Count < 2)
                {
                // 如果选择的形状数量小于2，则无法进行左对齐，直接返回
                return;
                }

            // 获取最后一个形状
            var lastShape = selectedShapes[selectedShapes.Count];

            // 获取最后一个形状的右侧位置
            var lastShapeLeft = lastShape.Left + lastShape.Width / 2;

            // 遍历选择的形状，将它们的左侧位置与最后一个形状的左侧位置对齐
            for (int i = 1 ; i <= selectedShapes.Count - 1 ; i++)
                {
                var shape = selectedShapes[i];
                shape.Left = lastShapeLeft - shape.Width / 2;
                }
            }

        private void button156_Click(object sender, RibbonControlEventArgs e)
            {
            // 获取选择的形状
            var selectedShapes = Globals.ThisAddIn.Application.ActiveWindow.Selection.ShapeRange;

            if (selectedShapes.Count < 2)
                {
                // 如果选择的形状数量小于2，则无法进行左对齐，直接返回
                return;
                }

            // 获取最后一个形状
            var lastShape = selectedShapes[selectedShapes.Count];

            // 获取最后一个形状的左侧位置
            var lastShapeLeft = lastShape.Top;

            // 遍历选择的形状，将它们的左侧位置与最后一个形状的左侧位置对齐
            for (int i = 1 ; i <= selectedShapes.Count - 1 ; i++)
                {
                var shape = selectedShapes[i];
                shape.Top = lastShapeLeft;
                }
            }

        private void button155_Click(object sender, RibbonControlEventArgs e)
            {
            // 获取选择的形状
            var selectedShapes = Globals.ThisAddIn.Application.ActiveWindow.Selection.ShapeRange;

            if (selectedShapes.Count < 2)
                {
                // 如果选择的形状数量小于2，则无法进行左对齐，直接返回
                return;
                }

            // 获取最后一个形状
            var lastShape = selectedShapes[selectedShapes.Count];

            // 获取最后一个形状的左侧位置
            var lastShapeLeft = lastShape.Top + lastShape.Height / 2;

            // 遍历选择的形状，将它们的左侧位置与最后一个形状的左侧位置对齐
            for (int i = 1 ; i <= selectedShapes.Count - 1 ; i++)
                {
                var shape = selectedShapes[i];
                shape.Top = lastShapeLeft - shape.Height / 2;
                }
            }

        private void button157_Click(object sender, RibbonControlEventArgs e)
            {
            // 获取选定的形状
            var selectedShapes = Globals.ThisAddIn.Application.ActiveWindow.Selection.ShapeRange;

            // 确保选定了至少两个形状
            if (selectedShapes.Count < 2)
                {
                return;
                }

            // 获取最后一个形状的底部位置
            var lastShape = selectedShapes[selectedShapes.Count];
            var lastBottom = lastShape.Top + lastShape.Height;

            // 遍历每个选定的形状，将其底部位置移动到最后一个形状的底部位置
            for (int i = 1 ; i < selectedShapes.Count ; i++)
                {
                var shape = selectedShapes[i];
                var shapeBottom = shape.Top + shape.Height;
                var offsetY = lastBottom - shapeBottom;
                shape.Top += offsetY;
                }
            }

        /// <summary>
        /// 获取选中组中的子项名称
        /// </summary>
        /// <returns></returns>
        /// <exception cref="Exception"></exception>
        public List<PowerPoint.Shape> GetShapesInSelectedGroup()
            {
            PowerPoint.Application pptApplication = Globals.ThisAddIn.Application;
            PowerPoint.Selection selection = pptApplication.ActiveWindow.Selection;

            if (selection.Type != PowerPoint.PpSelectionType.ppSelectionShapes)
                {
                throw new Exception("当前选中的不是形状，请选择组合形状并重试。");
                }

            List<PowerPoint.Shape> shapes = new List<PowerPoint.Shape>();
            for (int i = 1 ; i <= selection.ShapeRange.Count ; i++)
                {
                PowerPoint.Shape shape = selection.ShapeRange[i];

                if (shape.Type == MsoShapeType.msoGroup)
                    {
                    // 对于组合形状，获取其子项列表
                    var groupItems = shape.GroupItems;
                    foreach (PowerPoint.Shape item in groupItems)
                        {
                        shapes.Add(item);
                        }
                    }
                else
                    {
                    // 对于非组合形状，直接将其添加到子项列表中
                    shapes.Add(shape);
                    }
                }

            return shapes;
            }

        private void button127_Click(object sender, RibbonControlEventArgs e)
            {
            PowerPoint.Application pptApp = Globals.ThisAddIn.Application;
            PowerPoint.Selection sel = pptApp.ActiveWindow.Selection;

            if (sel.Type == PowerPoint.PpSelectionType.ppSelectionShapes)
                {
                foreach (PowerPoint.Shape shape in sel.ShapeRange)
                    {
                    // 获取形状的填充颜色
                    int rgb = shape.Fill.ForeColor.RGB;

                    // 计算互补颜色
                    int r = 255 - (rgb & 0xFF);
                    int g = 255 - ((rgb & 0xFF00) >> 8);
                    int b = 255 - ((rgb & 0xFF0000) >> 16);
                    int complementaryRgb = (b << 16) + (g << 8) + r;

                    // 将形状的填充颜色设置为互补颜色
                    shape.Fill.ForeColor.RGB = complementaryRgb;
                    }
                }
            else
                {
                Growl.WarningGlobal("请选择内容后再试！ ");
                }
            }

        private void button132_Click(object sender, RibbonControlEventArgs e)
            {
            // 获取选定的形状
            Selection selection = app.ActiveWindow.Selection;
            if (selection.Type == PpSelectionType.ppSelectionShapes || selection.Type == PpSelectionType.ppSelectionText)
                {
                PowerPoint.ShapeRange shapeRange = selection.ShapeRange;
                // 遍历选定的形状并更改颜色
                for (int i = 1 ; i <= shapeRange.Count ; i++)
                    {
                    PowerPoint.Shape shape = shapeRange[i];
                    PowerPoint.ColorFormat colorFormat = shape.Fill.ForeColor;

                    // 获取当前形状的颜色
                    int red = colorFormat.RGB & 0xff;
                    int green = (colorFormat.RGB >> 8) & 0xff;
                    int blue = (colorFormat.RGB >> 16) & 0xff;

                    // 计算颜色的相近色
                    int newRed = (red + 30) % 256;
                    int newGreen = (green + 30) % 256;
                    int newBlue = (blue + 30) % 256;

                    // 更改当前形状的颜色为相近色
                    colorFormat.RGB = (newBlue << 128) | (newGreen << 64) | newRed;
                    }
                }
            else
                {
                Growl.WarningGlobal("请选择内容后再试！ ");
                }
            }

        /// <summary>
        /// 跳转网址到浏览器
        /// </summary>
        /// <param name="Url"></param>
        public void Gotoweb(string Url)
            {
            if (!NetworkInterface.GetIsNetworkAvailable())
                {
                Growl.Warning("没有可用的网络连接");
                return;
                }

            // Check if the network is connected
            var connections = NetworkInterface.GetAllNetworkInterfaces();
            foreach (var connection in connections)
                {
                if (connection.OperationalStatus == OperationalStatus.Up)
                    {
                    System.Diagnostics.Process.Start(Url);
                    return;
                    }
                }

            Growl.Warning("没有可用的网络连接");
            }

        private void button56_Click(object sender, RibbonControlEventArgs e)
            {
            Gotoweb("https://unsplash.com/");
            }

        private void button159_Click(object sender, RibbonControlEventArgs e)
            {
            Gotoweb("https://pixabay.com/zh/");
            }

        private void button161_Click(object sender, RibbonControlEventArgs e)
            {
            Gotoweb("https://www.hippopx.com/zh");
            }

        private void button162_Click(object sender, RibbonControlEventArgs e)
            {
            Gotoweb("https://toppng.com/");
            }

        private void button163_Click(object sender, RibbonControlEventArgs e)
            {
            Gotoweb("http://www.51yuansu.com/");
            }

        private void button164_Click(object sender, RibbonControlEventArgs e)
            {
            Gotoweb("http://pngimg.com/");
            }

        private void button165_Click(object sender, RibbonControlEventArgs e)
            {
            Gotoweb("https://www.3dicons.com/");
            }

        private void button166_Click(object sender, RibbonControlEventArgs e)
            {
            Gotoweb("https://picular.co/");
            }

        private void button167_Click(object sender, RibbonControlEventArgs e)
            {
            Gotoweb("https://www.grabient.com/");
            }

        private void button168_Click(object sender, RibbonControlEventArgs e)
            {
            Gotoweb("http://www.colorlisa.com/");
            }

        private void button169_Click(object sender, RibbonControlEventArgs e)
            {
            Gotoweb("https://paletton.com/");
            }

        private void button170_Click(object sender, RibbonControlEventArgs e)
            {
            Gotoweb("https://www.fonts.net.cn/");
            }

        private void button171_Click(object sender, RibbonControlEventArgs e)
            {
            Growl.WarningGlobal("AI功能更新中");
            }

        /// <summary>
        /// 搜索替换文字内容
        /// </summary>
        /// <param name="searchText"></param>
        /// <param name="replaceText"></param>
        /// <returns>返回</returns>
        private void ReplaceText(PowerPoint.Shape shape, string searchText, string replaceText)
            {
            if (shape.HasTextFrame == Microsoft.Office.Core.MsoTriState.msoTrue)
                {
                PowerPoint.TextFrame textFrame = shape.TextFrame;
                if (textFrame.HasText == Microsoft.Office.Core.MsoTriState.msoTrue)
                    {
                    string text = textFrame.TextRange.Text;
                    if (text.Contains(searchText))
                        {
                        text = text.Replace(searchText, replaceText);
                        textFrame.TextRange.Text = text;
                        }
                    }
                }
            }

        private void button172_Click(object sender, RibbonControlEventArgs e)
            {
            var selection = Globals.ThisAddIn.Application.ActiveWindow.Selection;
            var slide = app.ActiveWindow.View.Slide;

            // 检查是否有选择
            if (selection.Type == PowerPoint.PpSelectionType.ppSelectionNone) return;

            IEnumerable<PowerPoint.Shape> shapesToProcess;

            if (Control.ModifierKeys == Keys.Control)
                {
                // 在控制键模式下，处理幻灯片上的所有形状
                shapesToProcess = slide.Shapes.Cast<PowerPoint.Shape>();
                }
            else
                {
                // 处理当前选择的形状
                shapesToProcess = selection.ShapeRange.Cast<PowerPoint.Shape>().Where(shape => shape.HasTextFrame == MsoTriState.msoTrue);
                }

            // 替换文本
            foreach (PowerPoint.Shape shape in shapesToProcess)
                {
                ReplaceText(shape, " ", "");
                }
            }

        private void button158_Click(object sender, RibbonControlEventArgs e)
            {
            Presentation presentation = app.ActivePresentation;
            string WithoutExtension = System.IO.Path.GetFileNameWithoutExtension(presentation.Name);

            if (presentation != null)
                {
                // 打开文件夹选择对话框，让用户选择导出文件夹
                var dialog = new FolderBrowserDialog();
                dialog.Description = "请选择拆分文件存放文件夹:";
                if (dialog.ShowDialog() == DialogResult.OK)
                    {
                    string exportPath = dialog.SelectedPath;
                    int slideCount = presentation.Slides.Count;
                    int threadsCount = System.Environment.ProcessorCount; // 获取 CPU 核心数
                    int slidesPerThread = (slideCount + threadsCount - 1) / threadsCount; // 计算每个线程处理的幻灯片数
                    var tasks = new List<System.Threading.Tasks.Task>();
                    for (int i = 0 ; i < threadsCount ; i++)
                        {
                        int startSlideIndex = i * slidesPerThread + 1;
                        int endSlideIndex = Math.Min(startSlideIndex + slidesPerThread - 1, slideCount);
                        System.Threading.Tasks.Task task = System.Threading.Tasks.Task.Run(() =>
                        {
                            for (int j = startSlideIndex ; j <= endSlideIndex ; j++)
                                {
                                string slidePath = System.IO.Path.Combine(exportPath, WithoutExtension + "_" + j + ".pptx");
                                presentation.Slides[j].Export(slidePath, "pptx", 0, 0);
                                }
                        });
                        tasks.Add(task);
                        }
                    System.Threading.Tasks.Task.WaitAll(tasks.ToArray());

                    Growl.SuccessGlobal("导出成功！");
                    }
                }
            else
                {
                Growl.Warning("请打开要拆分的 PowerPoint 文件！");
                }
            }

        private void button173_Click(object sender, RibbonControlEventArgs e)
            {
            //OpenFileDialog openFileDialog = new OpenFileDialog();
            //openFileDialog.Filter = "PPT文件 (*.ppt*;*.ppa*;*.pps*)|*.ppt*;*.ppa*;*.pps*|所有文件 (*.*)|*.*";
            //openFileDialog.Multiselect = true;
            //if (openFileDialog.ShowDialog()==DialogResult.OK)
            //    {
            //    PowerPoint.Application app = Globals.ThisAddIn.Application;
            //    PowerPoint.Presentation destinationPresentation = app.ActivePresentation;
            //    foreach (string file in openFileDialog.FileNames)
            //        {
            //        PowerPoint.Presentation sourcePresentation = app.Presentations.Open(file, Office.MsoTriState.msoFalse, Office.MsoTriState.msoFalse, Office.MsoTriState.msoFalse);
            //        foreach (PowerPoint.Slide sourceSlide in sourcePresentation.Slides)
            //            {
            //            PowerPoint.CustomLayout customLayout = sourceSlide.CustomLayout;
            //            PowerPoint.Slide newSlide = destinationPresentation.Slides.AddSlide(destinationPresentation.Slides.Count + 1, customLayout);
            //            sourceSlide.Copy();
            //            newSlide.Shapes.Paste();
            //            }
            //        sourcePresentation.Close();
            //        }
            //    }
            }

        private void button174_Click(object sender, RibbonControlEventArgs e)
            {
            var sel = app.ActiveWindow.Selection;
            var slide = app.ActiveWindow.View.Slide;

            switch (sel.Type)
                {
                case PowerPoint.PpSelectionType.ppSelectionNone:

                    Growl.WarningGlobal("可选中形状和图片元素导出为Bmp；选中多页幻灯片，只导出其中的图片元素");

                    break;

                case PowerPoint.PpSelectionType.ppSelectionShapes:
                    var range = sel.ShapeRange;
                    var name = Path.GetFileNameWithoutExtension(app.ActivePresentation.Name);
                    var cPath = Path.Combine(app.ActivePresentation.Path, $"{name} 的元素");

                    Directory.CreateDirectory(cPath);

                    var tasks = new List<Task>();
                    for (int i = 1 ; i <= range.Count ; i++)
                        {
                        var shape = range[i];
                        var dir = new DirectoryInfo(cPath);
                        var k = dir.GetFiles().Length + i;
                        var shname = $"{name}_{k}";

                        var task = Task.Run(() =>
                        {
                            shape.Export(Path.Combine(cPath, $"{shname}.bmp"), PowerPoint.PpShapeFormat.ppShapeFormatBMP);
                        });
                        tasks.Add(task);
                        }
                    Task.WaitAll(tasks.ToArray());

                    Process.Start("Explorer.exe", cPath);
                    break;

                case PowerPoint.PpSelectionType.ppSelectionSlides:
                    name = Path.GetFileNameWithoutExtension(app.ActivePresentation.Name);
                    cPath = Path.Combine(app.ActivePresentation.Path, $"{name} 的元素");

                    Directory.CreateDirectory(cPath);

                    tasks = new List<Task>();
                    foreach (PowerPoint.Slide item in sel.SlideRange)
                        {
                        for (int i = 1 ; i <= item.Shapes.Count ; i++)
                            {
                            var shape = item.Shapes[i];
                            if (shape.Type == Office.MsoShapeType.msoPicture)
                                {
                                var dir = new DirectoryInfo(cPath);
                                var k = dir.GetFiles().Length + i;
                                var shname = $"{name}_{k}";

                                var task = Task.Run(() =>
                                {
                                    shape.Export(Path.Combine(cPath, $"{shname}.bmp"), PowerPoint.PpShapeFormat.ppShapeFormatBMP);
                                });
                                tasks.Add(task);
                                }
                            }
                        }
                    Task.WaitAll(tasks.ToArray());

                    Process.Start("Explorer.exe", cPath);
                    break;
                }
            }

        private void button175_Click(object sender, RibbonControlEventArgs e)
            {
            }

        private void splitButton13_Click(object sender, RibbonControlEventArgs e)
            {
            Slide slide = Globals.ThisAddIn.Application.ActiveWindow.View.Slide;
            var app = Globals.ThisAddIn.Application;

            Selection sel = app.ActiveWindow.Selection;

            if (sel.Type == PpSelectionType.ppSelectionSlides)
                {
                // 如果是选中幻灯片，则遍历选中的幻灯片
                foreach (Slide item in sel.SlideRange)
                    {
                    DeleteShapes(item.Shapes);
                    }
                }
            else
                {
                // 如果没有选中幻灯片，则清空当前幻灯片中的所有形状
                DeleteShapes(slide.Shapes);
                }

            ShowResultMessage();

            void ShowResultMessage()
                {
                if (app.ActiveWindow.Selection.Type == PpSelectionType.ppSelectionNone)
                    {
                    // 如果幻灯片中没有任何形状，则显示警告提示信息
                    Growl.WarningGlobal("当前页面无内容！");
                    }
                else
                    {
                    // 显示操作成功提示信息
                    Growl.WarningGlobal("页面清理成功！");
                    }
                }
            }

        /// <summary>
        /// 形状类型删除
        /// </summary>
        /// <param name="shapeType"></param>
        public void DeleteShapesByType(MsoShapeType shapeType, string name)
            {
            PowerPoint.Selection sel = Globals.ThisAddIn.Application.ActiveWindow.Selection;
            int count = 0;

            if (Control.ModifierKeys.HasFlag(Keys.Control))
                {
                // 删除整个文档的指定类型的形状
                foreach (PowerPoint.Slide slide in Globals.ThisAddIn.Application.ActivePresentation.Slides)
                    {
                    foreach (PowerPoint.Shape shape in slide.Shapes)
                        {
                        if (shape.Type == shapeType)
                            {
                            shape.Delete();
                            count++;
                            }
                        }
                    }
                if (count > 0)
                    {
                    Growl.SuccessGlobal(string.Format("已删除 {0} 个 {1} 类型的媒体对象！", count, name));
                    }
                else
                    {
                    Growl.SuccessGlobal(string.Format("未找到类型为{0}的形状！", name));
                    }
                }
            else
                {
                // 获取当前活动窗口的选中对象
                if (sel.Type == PowerPoint.PpSelectionType.ppSelectionSlides)
                    {
                    // 如果是选中幻灯片，则遍历选中的幻灯片
                    foreach (PowerPoint.Slide slide in sel.SlideRange)
                        {
                        for (int i = slide.Shapes.Count ; i >= 1 ; i--)
                            {
                            PowerPoint.Shape shape = slide.Shapes[i];
                            if (shape.Type == shapeType)
                                {
                                shape.Delete();
                                count++;
                                }
                            }
                        }
                    if (count > 0)
                        {
                        Growl.SuccessGlobal(string.Format("已删除{0}个{1}类型的媒体对象！", count, name));
                        }
                    else
                        {
                        Growl.SuccessGlobal(string.Format("未找到类型为{0}的形状！", name));
                        }
                    }
                else
                    {
                    // 如果没有选中幻灯片，则删除当前幻灯片中的指定类型的形状
                    PowerPoint.Slide slide = Globals.ThisAddIn.Application.ActiveWindow.View.Slide;
                    for (int i = slide.Shapes.Count ; i >= 1 ; i--)
                        {
                        PowerPoint.Shape shape = slide.Shapes[i];
                        if (shape.Type == shapeType)
                            {
                            shape.Delete();
                            count++;
                            }
                        }
                    if (count > 0)
                        {
                        Growl.SuccessGlobal(string.Format("已删除{0}个{1}类型的媒体对象！", count, name));
                        }
                    else
                        {
                        Growl.SuccessGlobal(string.Format("未找到类型为{0}的形状！", name));
                        }
                    }
                }
            }

        private void button91_Click(object sender, RibbonControlEventArgs e)
            {
            // 获取当前活动窗口的幻灯片
            Slide slide = Globals.ThisAddIn.Application.ActiveWindow.View.Slide;

            // 判断是否按下 Ctrl 键
            if (Control.ModifierKeys.HasFlag(Keys.Control))
                {
                // 删除整个文档的动画
                foreach (Slide s in Globals.ThisAddIn.Application.ActivePresentation.Slides)
                    {
                    int num1 = s.TimeLine.MainSequence.Count;
                    for (int i1 = num1 ; i1 > 0 ; i1--)
                        {
                        Effect effect = s.TimeLine.MainSequence[i1];
                        effect.Delete();
                        }
                    }
                // 显示操作成功提示信息

                Growl.SuccessGlobal("整个文档的动画已删除！");
                }
            else
                {
                // 获取当前活动窗口的选中对象
                Selection sel = Globals.ThisAddIn.Application.ActiveWindow.Selection;
                // 判断选中对象的类型
                if (sel.Type == PpSelectionType.ppSelectionSlides)
                    {
                    // 如果是选中幻灯片，则遍历选中的幻灯片
                    foreach (Slide item in sel.SlideRange)
                        {
                        // 遍历幻灯片中的所有动画效果，并逐个删除
                        int num1 = item.TimeLine.MainSequence.Count;
                        for (int i1 = num1 ; i1 > 0 ; i1--)
                            {
                            Effect effect = item.TimeLine.MainSequence[i1];
                            effect.Delete();
                            }
                        }
                    Growl.SuccessGlobal("所选页面的动画已删除！");
                    }
                else
                    {
                    // 如果没有选中幻灯片，则删除当前幻灯片中的所有动画效果
                    int num = Globals.ThisAddIn.Application.ActiveWindow.View.Slide.TimeLine.MainSequence.Count;
                    for (int i = num ; i > 0 ; i--)
                        {
                        Effect effect = Globals.ThisAddIn.Application.ActiveWindow.View.Slide.TimeLine.MainSequence[i];
                        effect.Delete();
                        }
                    }
                Growl.SuccessGlobal("当前页面的动画已删除！");
                }
            }

        private void button176_Click(object sender, RibbonControlEventArgs e)
            {
            PowerPoint.CustomLayouts cLayouts = app.ActivePresentation.SlideMaster.CustomLayouts;
            int deletedCount = 0;
            for (int i = cLayouts.Count ; i >= 1 ; i--)
                {
                bool isUsed = false;
                foreach (PowerPoint.Slide slide in app.ActivePresentation.Slides)
                    {
                    if (slide.CustomLayout == cLayouts[i])
                        {
                        isUsed = true;
                        break;
                        }
                    }
                if (!isUsed)
                    {
                    cLayouts[i].Delete();
                    deletedCount++;
                    }
                }
            Growl.Success($"已删除 {deletedCount} 张未使用版式", "温馨提示");
            }

        private void button177_Click(object sender, RibbonControlEventArgs e)
            {
            PowerPoint.Presentation presentation = Globals.ThisAddIn.Application.ActivePresentation;
            PowerPoint.Slides slides = presentation.Slides;
            int count = slides.Count;
            int n = 0;

            for (int i = count ; i > 0 ; i--)
                {
                PowerPoint.Slide slide = slides[i];
                if (slide.Shapes.Count == 0 && slide.HeadersFooters.SlideNumber.Visible == Microsoft.Office.Core.MsoTriState.msoFalse)
                    {
                    slide.Delete();
                    n++;
                    }
                }

            string message;
            if (n > 0)
                {
                message = "已删除 " + n + " 张空白页。";
                Globals.ThisAddIn.Application.StartNewUndoEntry();
                }
            else
                {
                message = "未找到空白页。";
                }
            Growl.Success(message, "删除结果");
            }

        private void button178_Click(object sender, RibbonControlEventArgs e)
            {
            }

        private void button179_Click(object sender, RibbonControlEventArgs e)
            {
            Presentation pre = this.app.ActivePresentation;//定义幻灯片对象
            pre.RemoveDocumentInformation(PpRemoveDocInfoType.ppRDIAll);//删除演示文档信息
            }

        private void button180_Click(object sender, RibbonControlEventArgs e)
            {
            PowerPoint.Selection sel = app.ActiveWindow.Selection;
            PowerPoint.Slide slide = app.ActiveWindow.View.Slide;
            int count = 0;
            if (sel.Type == PowerPoint.PpSelectionType.ppSelectionShapes)
                {
                for (int i = slide.Shapes.Placeholders.Count ; i >= 1 ; i--)
                    {
                    slide.Shapes.Placeholders[i].Delete();
                    count++;
                    }
                }
            else
                {
                foreach (PowerPoint.Slide item in sel.SlideRange)
                    {
                    for (int i = item.Shapes.Placeholders.Count ; i >= 1 ; i--)
                        {
                        item.Shapes.Placeholders[i].Delete();
                        count++;
                        }
                    }
                }

            if (count > 0)
                {
                Growl.Success("已删除" + count + "个占位符。", "删除成功");
                }
            else
                {
                Growl.Success("未找到任何占位符。", "提示");
                }
            }

        private void button181_Click(object sender, RibbonControlEventArgs e)
            {
            DeleteShapesByType(MsoShapeType.msoPicture, "图片");
            }

        private void button182_Click(object sender, RibbonControlEventArgs e)
            {
            DeleteShapesByType(MsoShapeType.msoTable, "表格");
            }

        private void button183_Click(object sender, RibbonControlEventArgs e)
            {
            DeleteShapesByType(MsoShapeType.msoChart, "图表");
            }

        private void button184_Click(object sender, RibbonControlEventArgs e)
            {
            DeleteShapesByType(MsoShapeType.msoDiagram, "图");
            }

        private void button185_Click(object sender, RibbonControlEventArgs e)
            {
            DeleteShapesByType(MsoShapeType.msoCallout, "标注");
            }

        private void button186_Click(object sender, RibbonControlEventArgs e)
            {
            DeleteShapesByType(MsoShapeType.msoSmartArt, "Smart");
            }

        private void button187_Click(object sender, RibbonControlEventArgs e)
            {
            DeleteShapesByType(MsoShapeType.msoInkComment, "墨迹批注");
            }

        private void button188_Click(object sender, RibbonControlEventArgs e)
            {
            DeleteShapesByType(MsoShapeType.msoComment, "评论");
            }

        private void button189_Click(object sender, RibbonControlEventArgs e)
            {
            DeleteShapesByType(MsoShapeType.msoGroup, "组合");
            }

        private void button190_Click(object sender, RibbonControlEventArgs e)
            {// 获取当前活动窗口的选中对象
            Selection sel = Globals.ThisAddIn.Application.ActiveWindow.Selection;
            int count = 0; // 统计删除的文本框数目

            // 判断选中对象的类型
            if (sel.Type == PpSelectionType.ppSelectionSlides)
                {
                // 如果是选中幻灯片，则遍历选中的幻灯片
                foreach (Slide item in sel.SlideRange)
                    {
                    // 遍历幻灯片中的所有文本框
                    foreach (PowerPoint.Shape shape in item.Shapes)
                        {
                        if (shape.Type == MsoShapeType.msoTextBox)
                            {
                            PowerPoint.TextFrame textFrame = shape.TextFrame;
                            if (textFrame.TextRange.Text.Trim() == "")
                                {
                                shape.Delete();
                                count++;
                                }
                            }
                        }
                    }
                }
            else
                {
                // 如果没有选中幻灯片，则遍历当前幻灯片中的所有文本框
                Slide slide = Globals.ThisAddIn.Application.ActiveWindow.View.Slide;
                foreach (PowerPoint.Shape shape in slide.Shapes)
                    {
                    if (shape.Type == MsoShapeType.msoTextBox)
                        {
                        PowerPoint.TextFrame textFrame = shape.TextFrame;
                        if (textFrame.TextRange.Text.Trim() == "")
                            {
                            shape.Delete();
                            count++;
                            }
                        }
                    }
                }

            if (count > 0)
                {
                Growl.Success($"已删除 {count} 个空白文本框。");
                }
            else
                {
                Growl.Warning("当前幻灯片中不存在空白文本框。");
                }
            }

        /// <summary>
        /// 删除媒体
        /// </summary>
        /// <param name="mediaType"></param>
        public void DeleteMediaByType(PowerPoint.PpMediaType mediaType, string Name)
            {
            // 获取当前活动窗口的幻灯片
            PowerPoint.Slide slide = Globals.ThisAddIn.Application.ActiveWindow.View.Slide;

            int deletedCount = 0;
            foreach (PowerPoint.Shape shape in slide.Shapes)
                {
                // 判断是否为媒体对象
                if (shape.Type == Office.MsoShapeType.msoMedia && shape.MediaType == mediaType)
                    {
                    shape.Delete();
                    deletedCount++;
                    }
                }
            // 显示操作成功提示信息
            if (deletedCount > 0)
                {
                Growl.SuccessGlobal(string.Format("已删除 {0} 个 {1} 类型的媒体对象！", deletedCount, Name));
                }
            else
                {
                Growl.SuccessGlobal(string.Format("当前幻灯片中没有 {0} 类型的媒体对象。", Name));
                }
            }

        private void button192_Click(object sender, RibbonControlEventArgs e)
            {
            DeleteMediaByType(PpMediaType.ppMediaTypeSound, "音频");
            }

        private void button191_Click(object sender, RibbonControlEventArgs e)
            {
            DeleteMediaByType(PpMediaType.ppMediaTypeMovie, "视频");
            }

        private void button193_Click(object sender, RibbonControlEventArgs e)
            {
            DeleteMediaByType(PpMediaType.ppMediaTypeOther, "其他");
            }

        private void button194_Click(object sender, RibbonControlEventArgs e)
            {
            DeleteMediaByType(PpMediaType.ppMediaTypeMixed, "混合媒体");
            }

        /// <summary>
        /// 删除形状属性
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>

        public static void RemoveShapeProperties(PowerPoint.Shape shape, bool removeFillColor, bool removeFillTransparency, bool removeLineDashStyle, bool removeLineStyle, bool removeLineWidth, bool removeShadowType, bool removeSoftEdgeRadius)
            {
            // 检查 shape 是否为 null，如果是则返回
            if (shape == null) return;

            // 如果 removeFillColor 为 true，则将 shape 的填充颜色设置为透明
            if (removeFillColor)
                {
                shape.Fill.ForeColor.RGB = 0;
                shape.Fill.Transparency = 1.0f;
                }

            // 如果 removeFillTransparency 为 true，则将 shape 的填充透明度设置为 1.0
            if (removeFillTransparency)
                {
                shape.Fill.Transparency = 1.0f;
                }

            // 如果 removeLineDashStyle 为 true，则将 shape 的线条样式设置为实线
            if (removeLineDashStyle)
                {
                shape.Line.DashStyle = MsoLineDashStyle.msoLineSolid;
                }

            // 如果 removeLineStyle 为 true，则将 shape 的线条样式设置为单线
            if (removeLineStyle)
                {
                shape.Line.Style = MsoLineStyle.msoLineSingle;
                }

            // 如果 removeLineWidth 为 true，则将 shape 的线条宽度设置为 0
            if (removeLineWidth)
                {
                shape.Line.Weight = 0;
                }

            // 如果 removeShadowType 为 true，则将 shape 的阴影类型设置为无阴影
            if (removeShadowType)
                {
                shape.Shadow.Type = 0;
                }

            // 如果 removeSoftEdgeRadius 为 true，则将 shape 的柔化边缘半径设置为 0
            if (removeSoftEdgeRadius)
                {
                shape.SoftEdge.Radius = 0;
                }
            }

        private void button195_Click(object sender, RibbonControlEventArgs e)
            {
            }

        private void button197_Click(object sender, RibbonControlEventArgs e)
            {
            Gotoweb("https://www.pptx.cn/");
            }

        private void button196_Click(object sender, RibbonControlEventArgs e)
            {
            Gotoweb("https://www.aboutppt.com/");
            }

        private void button198_Click(object sender, RibbonControlEventArgs e)
            {
            Gotoweb("https://picwish.cn/remove-background");
            }

        private void button199_Click(object sender, RibbonControlEventArgs e)
            {
            Gotoweb("https://photokit.com/tools/cutout/?lang=zh");
            }

        private void button200_Click(object sender, RibbonControlEventArgs e)
            {
            Gotoweb("https://pixmiller.com/zh-hans/");
            }

        private void button201_Click(object sender, RibbonControlEventArgs e)
            {
            Gotoweb("https://huaban.com/search?q=PPT");
            }

        private void button82_Click(object sender, RibbonControlEventArgs e)
            {
            Wpf__Pagesize wpf__Pagesize = new Wpf__Pagesize();
            wpf__Pagesize.ShowDialog();
            }

        private void button202_Click(object sender, RibbonControlEventArgs e)
            {
            }

        private void button203_Click(object sender, RibbonControlEventArgs e)
            {
            // 获取当前选中的对象和活动幻灯片
            Selection selection = app.ActiveWindow.Selection;
            Slide slide = app.ActiveWindow.View.Slide;

            // 检查是否有选中对象
            bool flag = selection.Type == PpSelectionType.ppSelectionNone;
            if (flag)
                {
                Growl.Warning("请选中一张图片"); // 提示用户选择一张图片
                }
            else
                {
                PowerPoint.ShapeRange shapeRange = selection.ShapeRange;

                // 检查选中的对象是否为图片
                bool flag2 = shapeRange[1].Type == MsoShapeType.msoPicture;
                if (flag2)
                    {
                    PowerPoint.Shape shape = shapeRange[1];
                    float slideWidth = this.app.ActivePresentation.PageSetup.SlideWidth; // 获取幻灯片宽度
                    float slideHeight = this.app.ActivePresentation.PageSetup.SlideHeight; // 获取幻灯片高度

                    shape.LockAspectRatio = MsoTriState.msoTrue; // 锁定纵横比
                    shape.Width = slideWidth * 0.073f; // 设置形状宽度
                    shape.Cut(); // 剪切形状

                    // 粘贴形状为位图
                    slide.Shapes.PasteSpecial(PpPasteDataType.ppPasteBitmap, MsoTriState.msoFalse, "", 0, "", MsoTriState.msoFalse);
                    int count = slide.Shapes.Count; // 获取当前形状数量

                    // 设置粘贴形状的大小和位置
                    slide.Shapes[count].LockAspectRatio = MsoTriState.msoFalse; // 解锁纵横比
                    slide.Shapes[count].Left = 0f; // 设置左边距
                    slide.Shapes[count].Top = 0f; // 设置上边距
                    slide.Shapes[count].Width = slideWidth; // 设置宽度
                    slide.Shapes[count].Height = slideHeight; // 设置高度

                    // 添加模糊效果
                    PictureEffect pictureEffect = slide.Shapes[count].Fill.PictureEffects.Insert(MsoPictureEffectType.msoEffectBlur, 0);
                    pictureEffect.EffectParameters[1].Value = 10; // 设置模糊参数

                    slide.Shapes[count].Cut(); // 再次剪切形状
                    slide.Shapes.PasteSpecial(PpPasteDataType.ppPasteBitmap, MsoTriState.msoFalse, "", 0, "", MsoTriState.msoFalse); // 粘贴形状为位图

                    // 添加胶卷颗粒效果
                    PictureEffect pictureEffect2 = slide.Shapes[count].Fill.PictureEffects.Insert(MsoPictureEffectType.msoEffectFilmGrain, 0);
                    pictureEffect2.EffectParameters[1].Value = 0.3; // 设置颗粒度
                    pictureEffect2.EffectParameters[2].Value = 20; // 设置颗粒大小
                    }
                else
                    {
                    Growl.Warning("请选中一张图片"); // 提示用户选择一张图片
                    }
                }
            }

        private void button55_Click_1(object sender, RibbonControlEventArgs e)
            {
            }

        private void button204_Click(object sender, RibbonControlEventArgs e)
            {
            }

        private void bindingSource1_CurrentChanged(object sender, EventArgs e)
            {
            }

        private void button204_Click_1(object sender, RibbonControlEventArgs e)
            {
            Gotoweb("https://www.zcool.com.cn/");
            }

        private void button205_Click(object sender, RibbonControlEventArgs e)
            {
            Gotoweb("https://www.ui.cn/?ref=ppask.cn");
            }

        private void button206_Click(object sender, RibbonControlEventArgs e)
            {
            Gotoweb("https://dribbble.com/?ref=ppask.cn");
            }

        private void button138_Click_1(object sender, RibbonControlEventArgs e)
            {
            }

        private void button55_Click_2(object sender, RibbonControlEventArgs e)
            {
            }

        private void button7_Click_1(object sender, RibbonControlEventArgs e)
            {
            Wpf_shapeCopy wpf_ShapeCopy = new Wpf_shapeCopy();
            wpf_ShapeCopy.Show();
            }

        //// 创建并显示 XWindow1
        //XWindow1 window = new XWindow1();
        //window.OpenWindow(@"https://www.yuque.com/hiubook/vba");

        private void button138_Click_2(object sender, RibbonControlEventArgs e)
            {
            //Wpf_NameUnification wpf_NameUnification = new Wpf_NameUnification();
            //wpf_NameUnification.Show();
            // 创建并显示 XWindow1
            //XWindow1 window = new XWindow1();
            //window.OpenWindow(@"file:///C:/Users/Administrator/Desktop/Demo/Untitled-6.html", 560, 600);
            ////window.OpenWindow(@"https://chat.tinycms.xyz:2024/#/profile",1200,600);
            }

        #region

        private float CalculateOpacity(byte alpha, float minOpacity, float maxOpacity)
            {
            float opacityRange = maxOpacity - minOpacity;
            float normalizedAlpha = alpha / 255f;
            return minOpacity + normalizedAlpha * opacityRange;
            }

        #endregion

        /// <summary>
        /// 插入SVG
        /// </summary>
        /// <param name="svgFilePath"></param>
        public void InsertSvgIntoPpt(string svgFilePath)
            {
            PowerPoint.Application app = Globals.ThisAddIn.Application;
            PowerPoint.Slide slide = app.ActiveWindow.View.Slide;

            // 插入SVG文件到当前幻灯片
            slide.Shapes.AddPicture(svgFilePath, Office.MsoTriState.msoFalse, Office.MsoTriState.msoCTrue, 0, 0, -1, -1);
            }

        private void button207_Click(object sender, RibbonControlEventArgs e)
            {
            PresPio.Public_Wpf.Wpf_VoiceAssistant wpf_VoiceAssistant = new PresPio.Public_Wpf.Wpf_VoiceAssistant();
            wpf_VoiceAssistant.Show();
            }

        private void button208_Click(object sender, RibbonControlEventArgs e)
            {
            PowerPoint.Slide slide = null;
            PowerPoint.Shape shp = null;
            int i = 0;
            float slideWidth = 0;
            float slideHeight = 0;
            string outOfBoundsShapes = "幻灯片区域外的形状：" + Environment.NewLine;

            // 获取当前活动的幻灯片
            slide = app.ActiveWindow.View.Slide;

            // 获取幻灯片的宽度和高度
            slideWidth = slide.Master.Width;
            slideHeight = slide.Master.Height;

            // 遍历幻灯片中的所有形状
            for (i = slide.Shapes.Count ; i >= 1 ; i--)
                {
                shp = slide.Shapes[i];

                // 检查形状是否在画布范围内
                if (shp.Left + shp.Width < 0 || shp.Left > slideWidth ||
                    shp.Top + shp.Height < 0 || shp.Top > slideHeight)
                    {
                    // 记录形状名称
                    outOfBoundsShapes += shp.Name + Environment.NewLine;

                    // 删除形状
                    shp.Delete();
                    }
                }
            }

        private void button209_Click(object sender, RibbonControlEventArgs e)
            {
            }

        private void button210_Click(object sender, RibbonControlEventArgs e)
            {
            }

        /// <summary>
        /// 对比放大
        /// </summary>
        /// <param name="Sel"></param>
        /// <param name="scale"></param>
        private void CopyAndResizeSelection(Selection Sel, float scale)
            {
            if (Sel.Type == PpSelectionType.ppSelectionShapes)
                {
                PowerPoint.Shape shp = Sel.ShapeRange[1];
                Slide slide = app.ActiveWindow.View.Slide as Slide;

                shp.Copy();
                PowerPoint.Shape newShp = slide.Shapes.Paste()[1];

                newShp.LockAspectRatio = Office.MsoTriState.msoFalse;
                newShp.Width = shp.Width * scale;
                newShp.Height = shp.Height * scale;
                // newShp.Left = shp.Left + shp.Width / 2 - newShp.Width / 2;
                newShp.Left = shp.Left;
                newShp.Top = shp.Top - newShp.Height;

                if (newShp.HasTextFrame == Office.MsoTriState.msoTrue)
                    {
                    if (newShp.TextFrame.HasText == Office.MsoTriState.msoTrue)
                        {
                        var textRange = newShp.TextFrame.TextRange;
                        float originalFontSize = textRange.Font.Size;
                        textRange.Font.Size = originalFontSize * scale;
                        }
                    }
                }
            else if (Sel.Type == PpSelectionType.ppSelectionText)
                {
                var textRange = Sel.TextRange;
                Slide slide = app.ActiveWindow.View.Slide as Slide;

                PowerPoint.Shape newShp = slide.Shapes.AddTextbox(Office.MsoTextOrientation.msoTextOrientationHorizontal,
                     textRange.BoundLeft, textRange.BoundTop - textRange.BoundHeight * scale, textRange.BoundWidth, textRange.BoundHeight);

                newShp.TextFrame.TextRange.Text = textRange.Text;
                newShp.Width = textRange.BoundWidth * scale;
                newShp.Height = textRange.BoundHeight * scale;
                newShp.Left = textRange.BoundLeft + textRange.BoundWidth / 2 - newShp.Width / 2;
                newShp.Top = textRange.BoundTop - newShp.Height;

                float originalFontSize = textRange.Font.Size;
                newShp.TextFrame.TextRange.Font.Size = originalFontSize * scale;
                }
            else
                {
                MessageBox.Show("请选择一个形状或文本。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                }
            }

        private void button211_Click(object sender, RibbonControlEventArgs e)
            {
            try
                {
                Slide oSld = Globals.ThisAddIn.Application.ActiveWindow.Selection.SlideRange[1];
                foreach (PowerPoint.Shape oShp in Globals.ThisAddIn.Application.ActiveWindow.Selection.ShapeRange)
                    {
                    if (oShp.HasTextFrame == MsoTriState.msoTrue)
                        {
                        oShp.TextFrame.TextRange.Words(1).Text = "文头" + oShp.TextFrame.TextRange.Words(1).Text;
                        float oMargin = oShp.TextFrame.TextRange.Words(1, 2).BoundWidth;
                        oShp.TextFrame.TextRange.Words(1, 2).Delete();

                        for (int i = 1 ; i <= oShp.TextFrame.Ruler.Levels.Count ; i++)
                            {
                            if (oShp.TextFrame.Ruler.Levels[i].FirstMargin == 0)
                                {
                                oShp.TextFrame.Ruler.Levels[i].FirstMargin = oMargin;
                                }
                            else
                                {
                                oShp.TextFrame.Ruler.Levels[i].FirstMargin = 0;
                                }
                            oShp.TextFrame.Ruler.Levels[i].LeftMargin = 0;
                            }
                        }
                    }
                }
            catch (Exception ex)
                {
                Console.WriteLine("Error in IndentFirstLine: " + ex.Message);
                }
            }

        private void button213_Click(object sender, RibbonControlEventArgs e)
            {
            // 获取当前演示文稿
            Presentation presentation = app.ActivePresentation;
            FolderBrowserDialog folderBrowserDialog = new FolderBrowserDialog();
            folderBrowserDialog.Description = "选择要保存备注文件的文件夹路径";

            if (folderBrowserDialog.ShowDialog() == DialogResult.OK)
                {
                string saveFolderPath = folderBrowserDialog.SelectedPath;

                // 获取演示文稿的文件名（不含扩展名）
                string pptFileName = Path.GetFileNameWithoutExtension(presentation.FullName);

                // 构造备注文件的完整路径
                string notesFilePath = Path.Combine(saveFolderPath, pptFileName + "_Notes.txt");

                // 创建一个字符串来存储所有备注文本
                string allNotesText = "";

                // 循环遍历每个幻灯片
                foreach (Slide slide in presentation.Slides)
                    {
                    // 如果幻灯片有备注
                    if (slide.HasNotesPage == MsoTriState.msoTrue)
                        {
                        // 获取备注页
                        var notesPage = slide.NotesPage;

                        // 获取幻灯片的备注文本
                        string slideNotesText = notesPage.Shapes[2].TextFrame.TextRange.Text;

                        // 将幻灯片的备注文本添加到总的备注文本中
                        allNotesText += "Slide " + slide.SlideNumber + " Notes:\n" + slideNotesText + "\n\n";
                        }
                    }

                // 将所有备注文本写入到文件中
                File.WriteAllText(notesFilePath, allNotesText);

                Growl.SuccessGlobal("备注导出完成。文件保存在：" + notesFilePath);
                }
            else
                {
                Growl.WarningGlobal("未选择文件夹路径。");
                }
            }

        private void button214_Click(object sender, RibbonControlEventArgs e)
            {
            // 获取当前演示文稿
            Presentation presentation = app.ActivePresentation;
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "文本文件 (*.txt)|*.txt";
            openFileDialog.Title = "选择之前导出的备注文本文件";

            if (openFileDialog.ShowDialog() == DialogResult.OK)
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
                        Slide slide = presentation.Slides[i];

                        // 获取备注文本
                        string notesText = pageNotes[i - 1].Trim();

                        // 如果幻灯片有备注页
                        if (slide.HasNotesPage == MsoTriState.msoTrue)
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
                            TextRange textRange = notesSlide.Shapes[2].TextFrame.TextRange;
                            textRange.Text = notesText;
                            }
                        }

                    Growl.SuccessGlobal("备注导入完成。");
                    }
                catch (Exception ex)
                    {
                    Growl.ErrorGlobal("导入备注时发生错误：" + ex.Message);
                    }
                }
            else
                {
                Growl.WarningGlobal("未选择备注文本文件。");
                }
            }

        private void button55_Click_3(object sender, RibbonControlEventArgs e)
            {
            Selection selection = app.ActiveWindow.Selection;

            // 检查选择条件
            if (!IsValidSelection(selection))
                {
                return; // 提前退出以简化后续逻辑
                }

            Shape lineShape = selection.ShapeRange[1]; // 获取第一形状
            List<Shape> shapesToDistribute = CollectShapesToDistribute(selection);

            if (shapesToDistribute.Count > 0)
                {
                DistributeShapesAlongLine(lineShape, shapesToDistribute);
                }
            else
                {
                Growl.WarningGlobal("没有其他对象可以分布。");
                }

            // 验证选择是否有效
            bool IsValidSelection(Selection sel)
                {
                if (sel.Type != PpSelectionType.ppSelectionShapes || sel.ShapeRange.Count < 2)
                    {
                    Growl.WarningGlobal("请至少选择一条线段或曲线和一个对象。");
                    return false;
                    }

                Shape firstShape = sel.ShapeRange[1];
                if (firstShape.Type != MsoShapeType.msoLine && firstShape.Type != MsoShapeType.msoFreeform)
                    {
                    Growl.WarningGlobal("第一个选择的对象必须是线段或自由绘制的曲线。");
                    return false;
                    }

                return true;
                }

            // 收集可以分布的形状
            List<Shape> CollectShapesToDistribute(Selection sel)
                {
                List<Shape> shapes = new List<Shape>();
                for (int i = 2 ; i <= sel.ShapeRange.Count ; i++)
                    {
                    Shape shape = sel.ShapeRange[i];
                    if (shape.Type != MsoShapeType.msoLine && shape.Type != MsoShapeType.msoFreeform)
                        {
                        shapes.Add(shape);
                        }
                    }
                return shapes;
                }
            }

        private void DistributeShapesAlongLine(Shape lineShape, List<Shape> shapesToDistribute)
            {
            PowerPoint.ShapeNodes nodes = lineShape.Nodes;

            // 检查线段是否有至少两个节点
            if (nodes.Count < 2)
                {
                Growl.WarningGlobal("线段或曲线必须至少有两个节点。");
                return;
                }

            List<(float X, float Y)> nodePoints = new List<(float X, float Y)>();

            // 获取线段节点的坐标
            for (int i = 1 ; i <= nodes.Count ; i++)
                {
                nodePoints.Add(((float)nodes[i].Points[1, 1], (float)nodes[i].Points[1, 2]));
                }

            int shapeCount = shapesToDistribute.Count;
            float totalLength = GetTotalLength(nodePoints);
            float interval = totalLength / (shapeCount + 1);
            float currentLength = 0f;

            // 分布形状
            for (int i = 0 ; i < shapeCount ; i++)
                {
                currentLength += interval;
                var (x, y) = GetPointAtLength(nodePoints, currentLength);
                Shape shape = shapesToDistribute[i];
                shape.Left = x - shape.Width / 2f;
                shape.Top = y - shape.Height / 2f;

                if (lineShape.Type == MsoShapeType.msoFreeform)
                    {
                    AdjustShapeToLineCenter(shape, lineShape, x, y);
                    }
                }
            }

        // 计算线段的总长度
        private float GetTotalLength(List<(float X, float Y)> points)
            {
            float length = 0f;
            for (int i = 0 ; i < points.Count - 1 ; i++)
                {
                float dx = points[i + 1].X - points[i].X;
                float dy = points[i + 1].Y - points[i].Y;
                length += (float)Math.Sqrt(dx * dx + dy * dy);
                }
            return length;
            }

        // 根据累积长度找到对应的点坐标
        private (float X, float Y) GetPointAtLength(List<(float X, float Y)> points, float length)
            {
            float accumulatedLength = 0f;
            for (int i = 0 ; i < points.Count - 1 ; i++)
                {
                float dx = points[i + 1].X - points[i].X;
                float dy = points[i + 1].Y - points[i].Y;
                float segmentLength = (float)Math.Sqrt(dx * dx + dy * dy);

                if (accumulatedLength + segmentLength >= length)
                    {
                    float remainingLength = length - accumulatedLength;
                    float ratio = remainingLength / segmentLength;
                    float x = points[i].X + ratio * dx;
                    float y = points[i].Y + ratio * dy;
                    return (x, y);
                    }

                accumulatedLength += segmentLength;
                }
            return points[1];
            }

        // 调整形状使其中心与线段中心对齐
        private void AdjustShapeToLineCenter(PowerPoint.Shape shape, PowerPoint.Shape lineShape, float centerX, float centerY)
            {
            float num = shape.Left + shape.Width / 2f;
            float num2 = shape.Top + shape.Height / 2f;
            float num3 = centerX - num;
            float num4 = centerY - num2;
            shape.Left += num3;
            shape.Top += num4;
            }

        private void button212_Click(object sender, RibbonControlEventArgs e)
            {
            PresPio.Public_Wpf.Wpf_MaterialExport wpf_MaterialExport = new Public_Wpf.Wpf_MaterialExport(app);
            wpf_MaterialExport.Show();
            }

        private void button114_Click_2(object sender, RibbonControlEventArgs e)
            {
            PowerPoint.Application app = Globals.ThisAddIn.Application;
            PowerPoint.Selection sel = app.ActiveWindow.Selection;

            if (sel.Type == PowerPoint.PpSelectionType.ppSelectionNone)
                {
                Growl.WarningGlobal("请选中元素或幻灯片，且做好备份！");
                }
            else if (sel.Type == PowerPoint.PpSelectionType.ppSelectionShapes)
                {
                PowerPoint.ShapeRange range = sel.ShapeRange;
                PowerPoint.Slide slide = app.ActiveWindow.View.Slide;

                foreach (PowerPoint.Shape item in range)
                    {
                    item.Copy();
                    PowerPoint.Shape pic = slide.Shapes.PasteSpecial(PowerPoint.PpPasteDataType.ppPastePNG)[1];
                    pic.ScaleHeight(1f, Office.MsoTriState.msoTrue, Office.MsoScaleFrom.msoScaleFromMiddle);
                    pic.Left = item.Left + item.Width / 2 - pic.Width / 2;
                    pic.Top = item.Top + item.Height / 2 - pic.Height / 2;
                    item.Delete();
                    }
                }
            else if (sel.Type == PowerPoint.PpSelectionType.ppSelectionSlides)
                {
                PowerPoint.SlideRange slideRange = sel.SlideRange;

                foreach (PowerPoint.Slide slide in slideRange)
                    {
                    for (int i = slide.Shapes.Count ; i >= 1 ; i--)
                        {
                        slide.Shapes[i].Copy();
                        PowerPoint.Shape pic = slide.Shapes.PasteSpecial(PowerPoint.PpPasteDataType.ppPastePNG)[1];
                        pic.ScaleHeight(1f, Office.MsoTriState.msoTrue, Office.MsoScaleFrom.msoScaleFromMiddle);
                        pic.Left = slide.Shapes[i].Left + slide.Shapes[i].Width / 2 - pic.Width / 2;
                        pic.Top = slide.Shapes[i].Top + slide.Shapes[i].Height / 2 - pic.Height / 2;
                        slide.Shapes[i].Delete();
                        }
                    }
                Growl.SuccessGlobal("已将所选页面中的所有元素转为png图片");
                }
            else
                {
                PowerPoint.Slide slide = app.ActiveWindow.View.Slide;

                foreach (PowerPoint.Shape shape in slide.Shapes)
                    {
                    shape.Copy();
                    PowerPoint.Shape pic = slide.Shapes.PasteSpecial(PowerPoint.PpPasteDataType.ppPastePNG)[1];
                    pic.ScaleHeight(1f, Office.MsoTriState.msoTrue, Office.MsoScaleFrom.msoScaleFromMiddle);
                    pic.Left = shape.Left + shape.Width / 2 - pic.Width / 2;
                    pic.Top = shape.Top + shape.Height / 2 - pic.Height / 2;
                    shape.Delete();
                    }
                Growl.SuccessGlobal("已将当前页面中的所有元素转为png图片");
                }
            }

        private void button140_Click_1(object sender, RibbonControlEventArgs e)
            {
            }

        private void button140_Click_2(object sender, RibbonControlEventArgs e)
            {
            align_shapes("left");
            }

        /// <summary>
        /// 对齐函数
        /// </summary>
        /// <param name="align_type"></param>
        #region
        private bool is_to_shape = true;

        private void align_shapes(string align_type)
            {
            try
                {
                if (Globals.ThisAddIn.Application.ActiveWindow.Selection.ShapeRange.Count > 0)
                    {
                    var shape = Globals.ThisAddIn.Application.ActiveWindow.Selection.ShapeRange;

                    int count = shape.Count;

                    bool need_slide = false;

                    if (count == 1)
                        {
                        if (align_type == "hori_dist" || align_type == "vert_dist") { return; }
                        }

                    if (!this.is_to_shape || count == 1) { need_slide = true; }

                    if (need_slide) { count += 1; }

                    int[] index = new int[count];
                    for (int j = 0 ; j < count ; j++)
                        {
                        index[j] = j;
                        }

                    float[] centers_x = new float[count];
                    float[] centers_y = new float[count];
                    float[,] sizes = new float[count, 2];
                    float[,] corners = new float[count, 2];

                    float left_top_x = 99999;
                    float left_top_y = 99999;
                    float right_bottom_x = 0;
                    float right_bottom_y = 0;

                    var shapeRange = Globals.ThisAddIn.Application.ActiveWindow.Selection.ShapeRange;
                    int i = 0;
                    foreach (PowerPoint.Shape shp in shapeRange)
                        {
                        float w = shp.Width;
                        float h = shp.Height;
                        float t = shp.Top;
                        float l = shp.Left;

                        centers_x[i] = l + w / 2;
                        centers_y[i] = t + h / 2;

                        sizes[i, 0] = w;
                        sizes[i, 1] = h;

                        corners[i, 0] = l;
                        corners[i, 1] = t;

                        float r = l + w;
                        float b = t + h;

                        if (left_top_x > l) left_top_x = l;
                        if (left_top_y > t) left_top_y = t;
                        if (right_bottom_x < r) right_bottom_x = r;
                        if (right_bottom_y < b) right_bottom_y = b;

                        i++;
                        }

                    i = count - 1;
                    if (need_slide)
                        {
                        float w = Globals.ThisAddIn.Application.ActivePresentation.PageSetup.SlideWidth;
                        float h = Globals.ThisAddIn.Application.ActivePresentation.PageSetup.SlideHeight;
                        float l = 0;
                        float t = 0;

                        centers_x[i] = l + w / 2;
                        centers_y[i] = t + h / 2;

                        sizes[i, 0] = w;
                        sizes[i, 1] = h;

                        corners[i, 0] = l;
                        corners[i, 1] = t;
                        }

                    // 开始对齐

                    if (align_type == "hori_dist")
                        {
                        Array.Sort(centers_x, index);

                        int left_id = index[0];
                        int right_id = index[count - 1];

                        float left_w = sizes[left_id, 0];
                        float right_w = sizes[right_id, 0];

                        float between_gap = (centers_x[count - 1] - right_w / 2) - (centers_x[0] + left_w / 2);
                        float total_len = 0;
                        foreach (var id in index)
                            {
                            if (id == left_id || id == right_id) continue;
                            total_len += sizes[id, 0];
                            }
                        between_gap = between_gap - total_len;

                        between_gap /= count - 1;

                        float pre_right = centers_x[0] + left_w / 2;

                        foreach (var id in index)
                            {
                            if (id == left_id || id == right_id) continue;

                            int k = 0;
                            foreach (PowerPoint.Shape shp in shapeRange)
                                {
                                if (id == k)
                                    {
                                    float shp_w = shp.Width;

                                    pre_right += between_gap;
                                    shp.Left = pre_right;

                                    pre_right += shp_w;
                                    }
                                k++;
                                }
                            }
                        }
                    else if (align_type == "vert_dist")
                        {
                        Array.Sort(centers_y, index);

                        int top_id = index[0];
                        int bottom_id = index[count - 1];

                        float top_h = sizes[top_id, 1];
                        float bottom_h = sizes[bottom_id, 1];

                        float between_gap = (centers_y[count - 1] - bottom_h / 2) - (centers_y[0] + top_h / 2);
                        float total_len = 0;
                        foreach (var id in index)
                            {
                            if (id == top_id || id == bottom_id) continue;
                            total_len += sizes[id, 1];
                            }
                        between_gap = between_gap - total_len;

                        between_gap /= count - 1;

                        float pre_botttom = centers_y[0] + top_h / 2;

                        foreach (var id in index)
                            {
                            if (id == top_id || id == bottom_id) continue;

                            int k = 0;
                            foreach (PowerPoint.Shape shp in shapeRange)
                                {
                                if (id == k)
                                    {
                                    float shp_h = shp.Height;

                                    pre_botttom += between_gap;
                                    shp.Top = pre_botttom;

                                    pre_botttom += shp_h;
                                    }
                                k++;
                                }
                            }
                        }
                    else if (align_type == "hori_group")
                        {
                        float group_center_y = (left_top_y + right_bottom_y) / 2;

                        float slide_h = Globals.ThisAddIn.Application.ActivePresentation.PageSetup.SlideHeight;

                        float slide_center_y = slide_h / 2;

                        float slide_y_offset = slide_center_y - group_center_y;

                        foreach (PowerPoint.Shape shp in shapeRange)
                            {
                            shp.Top += slide_y_offset;
                            }
                        }
                    else if (align_type == "vert_group")
                        {
                        float group_center_x = (left_top_x + right_bottom_x) / 2;

                        float slide_w = Globals.ThisAddIn.Application.ActivePresentation.PageSetup.SlideWidth;

                        float slide_center_x = slide_w / 2;

                        float slide_x_offset = slide_center_x - group_center_x;

                        foreach (PowerPoint.Shape shp in shapeRange)
                            {
                            shp.Left += slide_x_offset;
                            }
                        }
                    else
                        {
                        i = 0;
                        foreach (PowerPoint.Shape shp in shapeRange)
                            {
                            switch (align_type)
                                {
                                case "horizontal":
                                    shp.Top = centers_y[count - 1] - sizes[i, 1] / 2;
                                    break;

                                case "vertical":
                                    shp.Left = centers_x[count - 1] - sizes[i, 0] / 2;
                                    break;

                                case "left":
                                    shp.Left = corners[count - 1, 0];
                                    break;

                                case "right":
                                    shp.Left = (corners[count - 1, 0] + sizes[count - 1, 0] - sizes[i, 0]);
                                    break;

                                case "top":
                                    shp.Top = corners[count - 1, 1];
                                    break;

                                case "bottom":
                                    shp.Top = (corners[count - 1, 1] + sizes[count - 1, 1] - sizes[i, 1]);
                                    break;

                                default:
                                    break;
                                }
                            i++;
                            }
                        }
                    }
                }
            catch (Exception)
                {
                }
            }

        private void update_enable()
            {
            to_shape.Enabled = !is_to_shape;
            to_slide.Enabled = is_to_shape;
            }

        #endregion

        private void button142_Click_1(object sender, RibbonControlEventArgs e)
            {
            align_shapes("right");
            }

        private void button141_Click_1(object sender, RibbonControlEventArgs e)
            {
            align_shapes("horizontal");
            }

        private void button143_Click_1(object sender, RibbonControlEventArgs e)
            {
            align_shapes("top");
            }

        private void button144_Click_1(object sender, RibbonControlEventArgs e)
            {
            align_shapes("vertical");
            }

        private void button145_Click_1(object sender, RibbonControlEventArgs e)
            {
            align_shapes("bottom");
            }

        private void button146_Click_1(object sender, RibbonControlEventArgs e)
            {
            align_shapes("hori_dist");
            }

        private void button147_Click_2(object sender, RibbonControlEventArgs e)
            {
            align_shapes("vert_dist");
            }

        private void button153_Click_1(object sender, RibbonControlEventArgs e)
            {
            }

        private void button155_Click_1(object sender, RibbonControlEventArgs e)
            {
            is_to_shape = false;
            update_enable();
            }

        private void button156_Click_1(object sender, RibbonControlEventArgs e)
            {
            is_to_shape = true;
            update_enable();
            }

        private void splitButton14_Click(object sender, RibbonControlEventArgs e)
            {
            PowerPoint.Selection sel = app.ActiveWindow.Selection;
            PowerPoint.Slide slide = app.ActiveWindow.View.Slide;

            if (sel.Type == PowerPoint.PpSelectionType.ppSelectionNone)
                {
                Growl.Warning("请选择一个元素", "温馨提示");
                }
            else if (sel.Type == PowerPoint.PpSelectionType.ppSelectionShapes)
                {
                PowerPoint.ShapeRange range = sel.ShapeRange;
                int count = range.Count;
                string[] name = new string[count];

                for (int i = 1 ; i <= count ; i++)
                    {
                    PowerPoint.Shape shape = range[i];
                    PowerPoint.Shape cshape = shape.Duplicate()[1];
                    cshape.Left = shape.Left;
                    cshape.Top = shape.Top;
                    name[i - 1] = cshape.Name;
                    }

                slide.Shapes.Range(name).Select();
                }
            else if (sel.Type == PowerPoint.PpSelectionType.ppSelectionSlides)
                {
                PowerPoint.SlideRange slides = sel.SlideRange;
                int count = slides.Count;

                for (int i = 1 ; i <= count ; i++)
                    {
                    PowerPoint.Slide slide0 = slides[i];
                    PowerPoint.Slide nslide = slide0.Duplicate()[1];
                    }
                }

            Growl.SuccessGlobal("数量+1");
            }

        private void button14_Click_1(object sender, RibbonControlEventArgs e)
            {
            ComHelper comHelper = new ComHelper();
            comHelper.DuplicateShapes("left");     // 将形状复制到左侧
            }

        private void button153_Click_2(object sender, RibbonControlEventArgs e)
            {
            ComHelper comHelper = new ComHelper();
            comHelper.DuplicateShapes("right");      // 将形状复制到右侧
            }

        private void button154_Click_1(object sender, RibbonControlEventArgs e)
            {
            ComHelper comHelper = new ComHelper();
            comHelper.DuplicateShapes("up");        // 将形状复制到上侧
            }

        private void button157_Click_1(object sender, RibbonControlEventArgs e)
            {
            ComHelper comHelper = new ComHelper();
            comHelper.DuplicateShapes("down");      // 将形状复制到下侧
            }

        private void button215_Click(object sender, RibbonControlEventArgs e)
            {
            ComHelper comHelper = new ComHelper();
            comHelper.DuplicateShapes("leftup");    // 将形状复制到左上侧
            }

        private void button216_Click(object sender, RibbonControlEventArgs e)
            {
            ComHelper comHelper = new ComHelper();
            comHelper.DuplicateShapes("leftdown");  // 将形状复制到左下侧
            }

        private void button217_Click(object sender, RibbonControlEventArgs e)
            {
            ComHelper comHelper = new ComHelper();
            comHelper.DuplicateShapes("rightup");   // 将形状复制到右上侧
            }

        private void button218_Click(object sender, RibbonControlEventArgs e)
            {
            ComHelper comHelper = new ComHelper();
            comHelper.DuplicateShapes("rightdown"); // 将形状复制到右下侧
            }

        private void button219_Click(object sender, RibbonControlEventArgs e)
            {
            ComHelper comHelper = new ComHelper();
            comHelper.ArrangeShapesInCircle(150f, false);
            }

        private void splitButton15_Click(object sender, RibbonControlEventArgs e)
            {
            Wpf_Manuscript wpf_Manuscript = new Wpf_Manuscript();
            wpf_Manuscript.Show();
            MyRibbon RB = Globals.Ribbons.Ribbon1;
            RB.splitButton15.Enabled = false;
            }

        private void button220_Click(object sender, RibbonControlEventArgs e)
            {
            Wpf_Qrcode wpf_Qrcode = new Wpf_Qrcode();
            wpf_Qrcode.Show();
            }

        private void splitButton16_Click(object sender, RibbonControlEventArgs e)
            {
            Wpf_Colortheif wpf_Colortheif = new Wpf_Colortheif();
            wpf_Colortheif.Show();
            }

        /// <summary>
        /// 格式化当前选定文本的属性，包括字体、大小、颜色、对齐方式等。
        /// </summary>
        /// <param name="app">当前 PowerPoint 应用实例。</param>
        /// <param name="fontName">字体名称，默认为 "微软雅黑"。</param>
        /// <param name="fontSize">字体大小，默认为 20。</param>
        /// <param name="fontColor">字体颜色，默认为黑色。</param>
        /// <param name="alignment">段落对齐方式，默认为左对齐。</param>
        /// <param name="isBold">是否加粗，默认为 false。</param>
        /// <param name="isItalic">是否斜体，默认为 false。</param>
        /// <param name="spaceBefore">段落前间距，默认为 0。</param>
        /// <param name="spaceAfter">段落后间距，默认为 0。</param>
        public void FormatSelectedText(PowerPoint.Application app,
            string fontName = "微软雅黑",
            float fontSize = 20,
            Color? fontColor = null, // 字体颜色，默认值为黑色
            PpParagraphAlignment alignment = PpParagraphAlignment.ppAlignLeft,
            bool isBold = false,
            bool isItalic = false,
            float spaceBefore = 0,
            float spaceAfter = 0)
            {
            try
                {
                // 获取当前选定的文本范围
                Selection selection = app.ActiveWindow.Selection;
                TextRange2 textRange = selection.TextRange2;

                // 设置字体名称
                textRange.Font.Name = fontName;
                textRange.Font.NameFarEast = fontName;
                textRange.Font.NameAscii = fontName;

                // 设置字体大小
                textRange.Font.Size = fontSize;

                // 设置字体颜色，如果未指定则使用黑色
                if (fontColor == null)
                    {
                    fontColor = Color.Black; // 默认颜色
                    }
                textRange.Font.Fill.ForeColor.RGB = ColorTranslator.ToOle(fontColor.Value);

                // 设置对齐方式
                textRange.ParagraphFormat.Alignment = (MsoParagraphAlignment)alignment;

                // 设置加粗和斜体
                textRange.Font.Bold = isBold ? MsoTriState.msoTrue : MsoTriState.msoFalse;
                textRange.Font.Italic = isItalic ? MsoTriState.msoTrue : MsoTriState.msoFalse;

                // 设置段落间距
                textRange.ParagraphFormat.SpaceBefore = spaceBefore;
                textRange.ParagraphFormat.SpaceAfter = spaceAfter;

                Console.WriteLine("选定文本格式已成功设置。");
                }
            catch (Exception ex)
                {
                Console.WriteLine("Error in FormatSelectedText: " + ex.Message);
                }
            }

        private void button2_Click_1(object sender, RibbonControlEventArgs e)
            {
            // 调用 FormatSelectedText，设置所有参数
            FormatSelectedText(app,
                fontName: "微软雅黑",
                fontSize: 18,
                fontColor: Color.Blue,
                alignment: PpParagraphAlignment.ppAlignRight,
                isBold: true,
                isItalic: false,
                spaceBefore: 10,
                spaceAfter: 10);
            }

        private void splitButton17_Click(object sender, RibbonControlEventArgs e)
            {
            PowerPoint.Selection selection = app.ActiveWindow.Selection;
            PowerPoint.ShapeRange selectedShapes = null;
            int i = 0;
            float tempTop = 0;
            float tempLeft = 0;

            // 检查是否选中了至少两个形状
            if (selection.ShapeRange.Count > 1)
                {
                // 获取选中的所有形状
                selectedShapes = selection.ShapeRange;

                // 遍历选中的形状并依次交换位置
                for (i = 1 ; i < selectedShapes.Count ; i++)
                    {
                    // 交换当前形状和前一个形状的位置
                    PowerPoint.Shape currentShape = selectedShapes[i];
                    PowerPoint.Shape previousShape = selectedShapes[i + 1];

                    tempTop = currentShape.Top;
                    tempLeft = currentShape.Left;

                    // 设置当前形状为前一个形状的位置
                    currentShape.Top = previousShape.Top;
                    currentShape.Left = previousShape.Left;

                    // 设置前一个形状为当前形状原来的位置
                    previousShape.Top = tempTop;
                    previousShape.Left = tempLeft;
                    }
                }
            else
                {
                Growl.WarningGlobal("请选择至少两个形状以进行交换。");
                }
            }

        private void button222_Click(object sender, RibbonControlEventArgs e)
            {
            var commandBars = Globals.ThisAddIn.Application.CommandBars;
            commandBars.ExecuteMso("Distributed");
            }

        private void button224_Click(object sender, RibbonControlEventArgs e)
            {
            var commandBars = Globals.ThisAddIn.Application.CommandBars;
            commandBars.ExecuteMso("AlignJustify");
            }

        private void button223_Click(object sender, RibbonControlEventArgs e)
            {
            }

        private void button175_Click_1(object sender, RibbonControlEventArgs e)
            {
            var commandBars = Globals.ThisAddIn.Application.CommandBars;
            commandBars.ExecuteMso("AlignLeft");
            }

        private void button209_Click_1(object sender, RibbonControlEventArgs e)
            {
            var commandBars = Globals.ThisAddIn.Application.CommandBars;
            commandBars.ExecuteMso("AlignCenter");
            }

        private void button210_Click_1(object sender, RibbonControlEventArgs e)
            {
            var commandBars = Globals.ThisAddIn.Application.CommandBars;
            commandBars.ExecuteMso("AlignRight");
            }

        private void button226_Click(object sender, RibbonControlEventArgs e)
            {
            Selection selection = app.ActiveWindow.Selection;

            // 确保选择的是文本框
            if (selection.Type == PpSelectionType.ppSelectionShapes)
                {
                foreach (Shape osh in selection.ShapeRange)
                    {
                    if (osh.HasTextFrame == MsoTriState.msoTrue)
                        {
                        if (osh.TextFrame.HasText == MsoTriState.msoTrue)
                            {
                            TextRange textRange = osh.TextFrame.TextRange;

                            // 使用 Text 属性查找换行符并替换
                            string text = textRange.Text;
                            string updatedText = text.Replace("\r", ""); // 去掉换行符

                            // 更新文本框内容
                            textRange.Text = updatedText;
                            }
                        }
                    }
                }
            else
                {
                System.Diagnostics.Debug.Print("Please select a shape with text.");
                }
            }

        private void toggleButton1_Click_1(object sender, RibbonControlEventArgs e)
            {
            //自适应窗格剪切板
            #region
            if (toggleButton1.Checked)
                {
                // 初始化TaskPane
                Use_Color use_Color = new Use_Color();
                TaskPaneShared.taskPane = Globals.ThisAddIn.CustomTaskPanes.Add(use_Color, "颜色助手");
                TaskPaneShared.taskPane.DockPosition = MsoCTPDockPosition.msoCTPDockPositionRight;
                TaskPaneShared.taskPane.DockPositionRestrict = MsoCTPDockPositionRestrict.msoCTPDockPositionRestrictNone;
                TaskPaneShared.taskPane.Width = 360;
                TaskPaneShared.taskPane.Visible = true;
                }
            else
                {
                // 隐藏并移除TaskPane
                if (TaskPaneShared.taskPane != null)
                    {
                    TaskPaneShared.taskPane.Visible = false;
                    Globals.ThisAddIn.CustomTaskPanes.Remove(TaskPaneShared.taskPane);
                    TaskPaneShared.taskPane.Dispose();
                    TaskPaneShared.taskPane = null; // 清理引用
                    }
                }

            // VisibleChange Event
            if (TaskPaneShared.taskPane != null)
                {
                TaskPaneShared.taskPane.VisibleChanged += new System.EventHandler(taskpane1_VisibleChanged);
                }
            #endregion
            }

        private void taskpane1_VisibleChanged(object sender, EventArgs e)//回调用户窗体事件
            {
            MyRibbon ribbon = Globals.Ribbons.GetRibbon<MyRibbon>();//获得功能区
            if (TaskPaneShared.taskPane.Visible)
                {
                ribbon.toggleButton1.Checked = true;
                }
            else
                {
                ribbon.toggleButton1.Checked = false;
                }
            }

        private void button228_Click(object sender, RibbonControlEventArgs e)
            {
            Wpf_ColorAdjust wpf_ColorAdjust = new Wpf_ColorAdjust(app);
            wpf_ColorAdjust.Show();
            }

        private void splitButton19_Click(object sender, RibbonControlEventArgs e)
            {
            Wpf_shapeCopy wpf_ShapeCopy = new Wpf_shapeCopy();
            wpf_ShapeCopy.Show();
            }

        private void button15_Click_1(object sender, RibbonControlEventArgs e)
            {
            PresPio.Wpf_Form.Wpf_SplitPPT wpf_SplitPPT = new Wpf_Form.Wpf_SplitPPT();
            wpf_SplitPPT.Show();
            }

        private void button227_Click(object sender, RibbonControlEventArgs e)
            {
            }

        private void button124_Click_1(object sender, RibbonControlEventArgs e)
            {
            Page_NotePage page_NotePage = new Page_NotePage();
            page_NotePage.Show();
            MyRibbon RB = Globals.Ribbons.Ribbon1;
            RB.button124.Enabled = false;
            }

        private void button227_Click_1(object sender, RibbonControlEventArgs e)
            {
            PresPio.Public_Wpf.Wpf_Polygon wpf_Polygon = new PresPio.Public_Wpf.Wpf_Polygon();
            wpf_Polygon.Show();
            }

        private void button229_Click(object sender, RibbonControlEventArgs e)
            {
            PresPio.Public_Wpf.Wpf_MaterialExport wpf_MaterialExport = new Public_Wpf.Wpf_MaterialExport(app);
            wpf_MaterialExport.Show();
            }

        private void splitButton18_Click(object sender, RibbonControlEventArgs e)
            {
            Globals.ThisAddIn.ShowTaskPane(@"https://chat.deepseek.com/", "Deepseek", 480);
            }

        private void button2_Click_2(object sender, RibbonControlEventArgs e)
            {
            // 将 sender 转换为 RibbonButton 或其他合适的控件类型
            var button = sender as RibbonButton;

            if (button != null)
                {
                // 获取按钮的标签
                string label = button.Label;
                string url = @"https://www.doubao.com/chat/";
                Globals.ThisAddIn.ShowTaskPane(url, label, 480);
                }
            else
                {
                System.Diagnostics.Debug.WriteLine("Sender is not a RibbonButton");
                }
            }

        private void button221_Click(object sender, RibbonControlEventArgs e)
            {
            Globals.ThisAddIn.ShowTaskPane("https://aigc365.cc/", "aigc365", 480);
            }

        private void button225_Click(object sender, RibbonControlEventArgs e)
            {
            var button = sender as RibbonButton;

            if (button != null)
                {
                // 获取按钮的标签
                string label = button.Label;
                string url = @"https://www.doubao.com/chat/";
                Globals.ThisAddIn.ShowTaskPane(url, label, 480);
                }
            else
                {
                System.Diagnostics.Debug.WriteLine("Sender is not a RibbonButton");
                }
            }

        private void button230_Click(object sender, RibbonControlEventArgs e)
            {
            var button = sender as RibbonButton;

            if (button != null)
                {
                // 获取按钮的标签
                string label = button.Label;
                string url = @"https://www.aippt.cn/generate?from=home";
                Globals.ThisAddIn.ShowTaskPane(url, label, 480);
                }
            else
                {
                System.Diagnostics.Debug.WriteLine("Sender is not a RibbonButton");
                }
            }

        private void button231_Click(object sender, RibbonControlEventArgs e)
            {
            var button = sender as RibbonButton;

            if (button != null)
                {
                // 获取按钮的标签
                string label = button.Label;
                string url = @"https://yuanbao.tencent.com/chat/";
                Globals.ThisAddIn.ShowTaskPane(url, label, 480);
                }
            else
                {
                System.Diagnostics.Debug.WriteLine("Sender is not a RibbonButton");
                }
            }

        private void button232_Click(object sender, RibbonControlEventArgs e)
            {
            var button = sender as RibbonButton;

            if (button != null)
                {
                // 获取按钮的标签
                string label = button.Label;
                string url = @"https://cp.baidu.com/";
                Globals.ThisAddIn.ShowTaskPane(url, label, 480);
                }
            else
                {
                System.Diagnostics.Debug.WriteLine("Sender is not a RibbonButton");
                }
            }

        private void button233_Click(object sender, RibbonControlEventArgs e)
            {
            var button = sender as RibbonButton;

            if (button != null)
                {
                // 获取按钮的标签
                string label = button.Label;
                string url = @"https://chat.deepseek.com/";
                Globals.ThisAddIn.ShowTaskPane(url, label, 480);
                }
            else
                {
                //System.Diagnostics.Debug.WriteLine("Sender is not a RibbonButton");
                }
            }

        private void button234_Click(object sender, RibbonControlEventArgs e)
            {
            var button = sender as RibbonButton;

            if (button != null)
                {
                // 获取按钮的标签
                string label = button.Label;
                string url = @"https://copilot.wps.cn/";
                Globals.ThisAddIn.ShowTaskPane(url, label, 480);
                }
            else
                {
                //System.Diagnostics.Debug.WriteLine("Sender is not a RibbonButton");
                }
            }

        private void button235_Click(object sender, RibbonControlEventArgs e)
            {
            var button = sender as RibbonButton;

            if (button != null)
                {
                // 获取按钮的标签
                string label = button.Label;
                string url = @"https://chatglm.cn/?lang=zh";
                Globals.ThisAddIn.ShowTaskPane(url, label, 480);
                }
            else
                {
                //System.Diagnostics.Debug.WriteLine("Sender is not a RibbonButton");
                }
            }

        private void button236_Click(object sender, RibbonControlEventArgs e)
            {
            var button = sender as RibbonButton;

            if (button != null)
                {
                // 获取按钮的标签
                string label = button.Label;
                string url = @"https://www.volcengine.com/experience/ark";
                Globals.ThisAddIn.ShowTaskPane(url, label, 480);
                }
            else
                {
                //System.Diagnostics.Debug.WriteLine("Sender is not a RibbonButton");
                }
            }

        private void button237_Click(object sender, RibbonControlEventArgs e)
            {
            var button = sender as RibbonButton;

            if (button != null)
                {
                // 获取按钮的标签
                string label = button.Label;
                string url = @"https://www.aboutppt.com/";
                Globals.ThisAddIn.ShowTaskPane(url, label, 480);
                }
            else
                {
                //System.Diagnostics.Debug.WriteLine("Sender is not a RibbonButton");
                }
            }

        private void button238_Click(object sender, RibbonControlEventArgs e)
            {
            var button = sender as RibbonButton;

            if (button != null)
                {
                // 获取按钮的标签
                string label = button.Label;
                string url = @"https://www.pptx.cn/";
                Globals.ThisAddIn.ShowTaskPane(url, label, 480);
                }
            else
                {
                //System.Diagnostics.Debug.WriteLine("Sender is not a RibbonButton");
                }
            }

        private void button239_Click(object sender, RibbonControlEventArgs e)
            {
            var button = sender as RibbonButton;

            if (button != null)
                {
                // 获取按钮的标签
                string label = button.Label;
                string url = @"https://uzfqt.xetlk.com/s/2W8JQB/";
                Globals.ThisAddIn.ShowTaskPane(url, label, 480);
                }
            else
                {
                //System.Diagnostics.Debug.WriteLine("Sender is not a RibbonButton");
                }
            }

        private void button240_Click(object sender, RibbonControlEventArgs e)
            {
            var selection = app.ActiveWindow.Selection;
            if (selection.ShapeRange.Count > 1)
            {
                // 获取第一个形状的所有颜色属性
                var firstShape = selection.ShapeRange[1];
                
                // 获取填充颜色
                var fillColor = firstShape.Fill.ForeColor.RGB;
                
                // 获取边框颜色
                var lineColor = firstShape.Line.ForeColor.RGB;
                
                // 获取文字颜色
                var textColor = firstShape.HasTextFrame == MsoTriState.msoTrue && 
                              firstShape.TextFrame.HasText == MsoTriState.msoTrue ?
                              firstShape.TextFrame.TextRange.Font.Color.RGB : (int?)null;
                
                // 获取阴影颜色
                var shadowColor = firstShape.Shadow.Type == MsoShadowType.msoShadowMixed ? 
                                (int?)null : firstShape.Shadow.ForeColor.RGB;

                // 遍历所有选中的形状
                for (int i = 2; i <= selection.ShapeRange.Count; i++)
                {
                    var shape = selection.ShapeRange[i];
                    
                    // 设置填充颜色
                    if (shape.Fill.Visible == MsoTriState.msoTrue)
                    {
                        shape.Fill.ForeColor.RGB = fillColor;
                    }
                    
                    // 设置边框颜色
                    if (shape.Line.Visible == MsoTriState.msoTrue)
                    {
                        shape.Line.ForeColor.RGB = lineColor;
                    }
                    
                    // 设置文字颜色
                    if (textColor.HasValue && 
                        shape.HasTextFrame == MsoTriState.msoTrue &&
                        shape.TextFrame.HasText == MsoTriState.msoTrue)
                    {
                        shape.TextFrame.TextRange.Font.Color.RGB = textColor.Value;
                    }
                    
                    // 设置阴影颜色
                    if (shadowColor.HasValue && 
                        shape.Shadow.Type != MsoShadowType.msoShadowMixed)
                    {
                        shape.Shadow.ForeColor.RGB = shadowColor.Value;
                    }
                }
            }
            else
            {
                Growl.Warning("请选择多个对象以统一颜色");
            }
                
            }

        private void button243_Click(object sender, RibbonControlEventArgs e)
            {

            var selection = app.ActiveWindow.Selection;
            if (selection.ShapeRange.Count > 1)
                {
                // 获取第一个形状作为基准
                var firstShape = selection.ShapeRange[1];

                // 获取基准尺寸和字体信息
                float baseWidth = firstShape.Width;
                float baseHeight = firstShape.Height;
                float? baseFontSize = null;

                // 如果形状包含文本，获取字体大小
                if (firstShape.HasTextFrame == MsoTriState.msoTrue &&
                    firstShape.TextFrame.HasText == MsoTriState.msoTrue)
                    {
                    baseFontSize = firstShape.TextFrame.TextRange.Font.Size;
                    }

                // 遍历所有选中的形状
                for (int i = 2 ; i <= selection.ShapeRange.Count ; i++)
                    {
                    var shape = selection.ShapeRange[i];

                    // 应用基准尺寸
                    shape.Width = baseWidth;
                    shape.Height = baseHeight;

                    // 如果基准有字体大小且当前形状有文本，应用字体大小
                    if (baseFontSize.HasValue &&
                        shape.HasTextFrame == MsoTriState.msoTrue &&
                        shape.TextFrame.HasText == MsoTriState.msoTrue)
                        {
                        shape.TextFrame.TextRange.Font.Size = baseFontSize.Value;
                        }
                    }

               // Growl.SuccessGlobal("尺寸和字体镜像已完成");
                }
            else
                {
               // Growl.WarningGlobal("请选择多个对象以应用尺寸镜像");
                }

            }

        private void button241_Click(object sender, RibbonControlEventArgs e)
            {
                 var selection = app.ActiveWindow.Selection;
                 if (selection.ShapeRange.Count > 1)
                     {
                     // 获取第一个形状作为基准
                     var firstShape = selection.ShapeRange[1];
                     float baseHeight = firstShape.Height;

                     // 遍历所有选中的形状
                     for (int i = 2; i <= selection.ShapeRange.Count; i++)
                         {
                         var shape = selection.ShapeRange[i];
                         // 应用基准高度
                         shape.Height = baseHeight;
                         }

                     //Growl.SuccessGlobal("高度统一已完成");
                     }
                 else
                     {
                     //Growl.WarningGlobal("请选择多个对象以应用高度统一");
                     }
            }

        private void button242_Click(object sender, RibbonControlEventArgs e)
            {
            var selection = app.ActiveWindow.Selection;
            if (selection.ShapeRange.Count > 1)
            {
                // 获取第一个形状作为基准
                var firstShape = selection.ShapeRange[1];
                float baseWidth = firstShape.Width;

                // 遍历所有选中的形状
                for (int i = 2; i <= selection.ShapeRange.Count; i++)
                {
                    var shape = selection.ShapeRange[i];
                    // 应用基准宽度
                    shape.Width = baseWidth;
                }

                //Growl.SuccessGlobal("宽度统一已完成");
            }
            else
            {
                //Growl.WarningGlobal("请选择多个对象以应用宽度统一");
            }
            }

        private void button244_Click(object sender, RibbonControlEventArgs e)
            {
           
            var selection = app.ActiveWindow.Selection;
            if (selection.ShapeRange.Count > 1)
                {
                var firstShape = selection.ShapeRange[1];

                // 复制更多属性：透明度、线条样式、线条粗细、渐变设置等
                var fillTransparency = firstShape.Fill.Transparency;
                var lineWeight = firstShape.Line.Weight;
                var lineDashStyle = firstShape.Line.DashStyle;

                for (int i = 2 ; i <= selection.ShapeRange.Count ; i++)
                    {
                    var shape = selection.ShapeRange[i];

                    // 应用所有属性
                    shape.Fill.Transparency = fillTransparency;
                    shape.Line.Weight = lineWeight;
                    shape.Line.DashStyle = lineDashStyle;
                    // 可添加更多属性
                    }

                Growl.SuccessGlobal("完整格式镜像已完成");
                }
           
        }

        private void button245_Click(object sender, RibbonControlEventArgs e)
            {
      
            var selection = app.ActiveWindow.Selection;
            if (selection.ShapeRange.Count > 1)
                {
                // 获取所有形状相对于第一个形状的位置关系
                var firstShape = selection.ShapeRange[1];
                float baseLeft = firstShape.Left;
                float baseTop = firstShape.Top;

                // 将所有形状按照相对于另一点的位置进行排列
                float targetLeft = 100; // 新的基准点x坐标
                float targetTop = 100;  // 新的基准点y坐标

                for (int i = 1 ; i <= selection.ShapeRange.Count ; i++)
                    {
                    var shape = selection.ShapeRange[i];
                    float offsetX = shape.Left - baseLeft;
                    float offsetY = shape.Top - baseTop;

                    shape.Left = targetLeft + offsetX;
                    shape.Top = targetTop + offsetY;
                    }

                Growl.SuccessGlobal("位置关系镜像已完成");
                }
             }

        private void button246_Click(object sender, RibbonControlEventArgs e)
            {
            var selection = app.ActiveWindow.Selection;
            if (selection.ShapeRange.Count > 1)
                {
                var firstShape = selection.ShapeRange[1];
                float rotation = firstShape.Rotation;

                for (int i = 2 ; i <= selection.ShapeRange.Count ; i++)
                    {
                    selection.ShapeRange[i].Rotation = rotation;
                    }

                Growl.SuccessGlobal("旋转角度镜像已完成");
                }
            }

        private void button247_Click(object sender, RibbonControlEventArgs e)
            {
            var selection = app.ActiveWindow.Selection;
            if (selection.ShapeRange.Count > 1)
                {
                var firstShape = selection.ShapeRange[1];
                if (firstShape.HasTextFrame == MsoTriState.msoTrue &&
                    firstShape.TextFrame.HasText == MsoTriState.msoTrue)
                    {
                    var textRange = firstShape.TextFrame.TextRange;
                    var fontName = textRange.Font.Name;
                    var fontBold = textRange.Font.Bold;
                    var fontItalic = textRange.Font.Italic;
                    var fontUnderline = textRange.Font.Underline;
                    var alignment = textRange.ParagraphFormat.Alignment;
                    var lineSpacing = textRange.ParagraphFormat.LineRuleWithin;

                    for (int i = 2 ; i <= selection.ShapeRange.Count ; i++)
                        {
                        var shape = selection.ShapeRange[i];
                        if (shape.HasTextFrame == MsoTriState.msoTrue &&
                            shape.TextFrame.HasText == MsoTriState.msoTrue)
                            {
                            var targetText = shape.TextFrame.TextRange;
                            targetText.Font.Name = fontName;
                            targetText.Font.Bold = fontBold;
                            targetText.Font.Italic = fontItalic;
                            targetText.Font.Underline = fontUnderline;
                            targetText.ParagraphFormat.Alignment = alignment;
                            targetText.ParagraphFormat.LineRuleWithin = lineSpacing;
                            }
                        }

                    Growl.SuccessGlobal("文本格式镜像已完成");
                    }
                }
            }

        private void button248_Click(object sender, RibbonControlEventArgs e)
            {
            var selection = app.ActiveWindow.Selection;
            if (selection.ShapeRange.Count > 1)
                {
                var firstShape = selection.ShapeRange[1];
                var slide = app.ActiveWindow.View.Slide;

                // 查找第一个形状的动画
                PowerPoint.Effect firstEffect = null;
                foreach (PowerPoint.Effect effect in slide.TimeLine.MainSequence)
                    {
                    if (effect.Shape.Name == firstShape.Name)
                        {
                        firstEffect = effect;
                        break;
                        }
                    }

                if (firstEffect != null)
                    {
                    for (int i = 2 ; i <= selection.ShapeRange.Count ; i++)
                        {
                        var shape = selection.ShapeRange[i];
                        // 删除该形状的现有动画
                        for (int j = slide.TimeLine.MainSequence.Count ; j >= 1 ; j--)
                            {
                            if (slide.TimeLine.MainSequence[j].Shape.Name == shape.Name)
                                {
                                slide.TimeLine.MainSequence[j].Delete();
                                }
                            }

                        // 添加与第一个形状相同的动画
                        slide.TimeLine.MainSequence.AddEffect(
                            shape,
                            firstEffect.EffectType,
                            firstEffect.Timing.TriggerType,
                            firstEffect.Index);
                        }

                    Growl.SuccessGlobal("动画效果镜像已完成");
                    }
                }
            }
        }
    }