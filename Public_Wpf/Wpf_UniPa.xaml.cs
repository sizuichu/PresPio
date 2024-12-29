using HandyControl.Controls;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using System.Windows;

using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace PresPio
    {
    /// <summary>
    /// Wpf_UniPa.xaml 的交互逻辑
    /// </summary>
    public partial class Wpf_UniPa
        {
        private PowerPoint.Application app; //加载PPT项目

        public Wpf_UniPa()
            {
            InitializeComponent();
            app = Globals.ThisAddIn.Application;
            string[] arr = {
    "左对齐", // 左对齐
    "居中对齐", // 居中对齐
    "右对齐", // 右对齐
    "两端对齐", // 两端对齐
    "分散对齐" // 分散对齐
};

            // 假设 uiComboBox1 是你的 ComboBox 控件的名称
            uiComboBox1.ItemsSource = arr;
            uiComboBox1.SelectedIndex = 3;
            }

        private void uniWindow_Loaded(object sender, RoutedEventArgs e)
            {
            uiComboBox1.SelectedIndex = 3;
            }

        private void uiCheckBox1_Click(object sender, RoutedEventArgs e)
            {
            // 使用三元运算符简化代码
            uiTextBox1.IsEnabled = uiCheckBox1.IsChecked.Value ? true : false;
            }

        private void uiCheckBox2_Click(object sender, RoutedEventArgs e)
            {
            // 使用三元运算符简化代码
            uiTextBox2.IsEnabled = uiCheckBox2.IsChecked.Value ? true : false;
            }

        private void uiCheckBox3_Click(object sender, RoutedEventArgs e)
            {
            // 使用三元运算符简化代码
            uiTextBox3.IsEnabled = uiCheckBox3.IsChecked.Value ? true : false;
            }

        private void uiCheckBox4_Click(object sender, RoutedEventArgs e)
            {
            // 使用三元运算符简化代码
            uiComboBox1.IsEnabled = uiCheckBox3.IsChecked.Value ? true : false;
            }

        private void Button_Click(object sender, RoutedEventArgs e)
            {
            Slide slide = app.ActiveWindow.View.Slide;
            Presentation pre = app.ActivePresentation;
            Selection sel = app.ActiveWindow.Selection;
            if (uiRadioButton1.IsChecked == true)
                {
                if (sel.Type == PpSelectionType.ppSelectionSlides && sel.SlideRange.Count > 0)
                    {
                    foreach (Slide slide1 in sel.SlideRange)
                        {
                        foreach (PowerPoint.Shape shape in slide1.Shapes)
                            {
                            if (shape.HasTextFrame == MsoTriState.msoTrue && shape.TextFrame.HasText == MsoTriState.msoTrue)
                                {
                                UniPara(shape);//调用函数
                                }
                            }
                        }
                    Growl.SuccessGlobal("所选页面段落统一成功");
                    }
                else
                    {
                    Growl.WarningGlobal("请选择幻灯片页面");
                    }
                }
            //全部文档
            if (uiRadioButton2.IsChecked == true)
                {
                // 遍历所有的幻灯片
                foreach (Slide slide2 in pre.Slides)
                    {
                    // 遍历当前幻灯片中的所有形状
                    foreach (PowerPoint.Shape shape in slide2.Shapes)
                        {
                        // 如果形状是文本框，则进行修改
                        if (shape.HasTextFrame == MsoTriState.msoTrue && shape.TextFrame.HasText == MsoTriState.msoTrue)
                            {
                            UniPara(shape);//调用函数
                            }
                        }
                    }
                Growl.SuccessGlobal("全文页面段落统一成功");
                }
            }

        public void UniPara(PowerPoint.Shape shape)
            {
            // 如果形状是文本框，则获取其中的文本
            if (shape.HasTextFrame == MsoTriState.msoTrue && shape.TextFrame.HasText == MsoTriState.msoTrue)
                {
                PowerPoint.TextRange textRange = shape.TextFrame.TextRange;
                // 设置段落的行距、段前、段后间距和对齐方式
                if (uiCheckBox1.IsChecked == true)
                    {
                    textRange.ParagraphFormat.LineRuleWithin = MsoTriState.msoTrue;//设置为行数
                    textRange.ParagraphFormat.SpaceWithin = (float)uiTextBox1.Value;//段落间距
                    }
                if (uiCheckBox2.IsChecked == true)
                    {
                    textRange.ParagraphFormat.SpaceBefore = (float)uiTextBox2.Value;
                    }
                if (uiCheckBox3.IsChecked == true)
                    {
                    textRange.ParagraphFormat.SpaceAfter = (float)uiTextBox3.Value;
                    }
                if (uiCheckBox4.IsChecked == true)
                    {
                    if (uiComboBox1.SelectedIndex == 0)
                        {
                        textRange.ParagraphFormat.Alignment = PpParagraphAlignment.ppAlignLeft;
                        }
                    else if (uiComboBox1.SelectedIndex == 1)
                        {
                        textRange.ParagraphFormat.Alignment = PpParagraphAlignment.ppAlignCenter;
                        }
                    else if (uiComboBox1.SelectedIndex == 2)
                        {
                        textRange.ParagraphFormat.Alignment = PpParagraphAlignment.ppAlignRight;
                        }
                    else if (uiComboBox1.SelectedIndex == 3)
                        {
                        textRange.ParagraphFormat.Alignment = PpParagraphAlignment.ppAlignJustify;
                        }
                    else
                        {
                        textRange.ParagraphFormat.Alignment = PpParagraphAlignment.ppAlignDistribute;
                        }
                    }
                }
            }
        }
    }