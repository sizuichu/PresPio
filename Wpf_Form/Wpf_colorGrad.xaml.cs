using HandyControl.Controls;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using System;
using System.Linq;
using System.Windows;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
namespace PresPio
    {
    public partial class Wpf_colorGrad
        {
        public Microsoft.Office.Interop.PowerPoint.Application app;
        public double step;
        public double Gra;
        public double ColorStep;
        public Wpf_colorGrad()
            {
            InitializeComponent();
            step = Slider1.ValueEnd - Slider1.ValueStart;
            Gra = Slider2.ValueEnd - Slider2.ValueStart;
            ColorStep = Slider3.ValueEnd - Slider3.ValueStart;
            // 检查Slider1、Slider2和Slider3的值是否已经设置，如果没有设置，先设置默认值
            if (step == 0)
                {
                Slider1.ValueStart = 0;
                Slider1.ValueEnd = 100;
                step = Slider1.ValueEnd - Slider1.ValueStart;
                }

            if (Gra == 0)
                {
                Slider2.ValueStart = 0;
                Slider2.ValueEnd = 100;
                Gra = Slider2.ValueEnd - Slider2.ValueStart;
                }

            if (ColorStep == 0)
                {
                Slider3.ValueStart = 0;
                Slider3.ValueEnd = 360;
                ColorStep = Slider3.ValueEnd - Slider3.ValueStart;
                }
            }

        private void CardWindown_Loaded(object sender, RoutedEventArgs e)
            {
            app = Globals.ThisAddIn.Application;
            }

        /// <summary>
        /// 创建渐变色，输入步进值
        /// </summary>
        /// <param name="Step"></param>
        /// <returns></returns>
        public int CreatColor(double Step)
            {
            int color = 0;
            Random ran = new Random();
            int R = ran.Next(0, 256);
            int G = ran.Next(0, 256);
            int B = ran.Next(0, 256);
            //开始随机算法
            color = (int)(R + G * 255 + B * 255 * 255 - Step);
            return color;
            }
        /// <summary>
        /// 生成渐变色
        /// </summary>
        /// <param name="step">步进值</param>
        /// <param name="Gra">角度</param>
        public void CreatShap(double step, double Gra, double ColorStep)
            {
            app = Globals.ThisAddIn.Application;
            DelShpe("渐变色块");
            Slide slide = app.ActiveWindow.View.Slide;
            slide.Select();
            //新建10个形状
            Microsoft.Office.Interop.PowerPoint.Shape[] shps = new Microsoft.Office.Interop.PowerPoint.Shape[16];
            for (int i = 0 ; i < 16 ; i++)
                {
                shps[i] = slide.Shapes.AddShape(MsoAutoShapeType.msoShapeRectangle, 60 * i, 0, 50, 50);
                shps[i].Fill.TwoColorGradient(MsoGradientStyle.msoGradientHorizontal, 1);
                shps[i].Tags.Add("图形", "渐变色块");
                //shps[i].Name = "渐变色块";
                shps[i].Fill.GradientAngle = (float)Gra;
                shps[i].Fill.TwoColorGradient(MsoGradientStyle.msoGradientHorizontal, 1);
                shps[i].Line.Visible = MsoTriState.msoFalse;
                }
            for (int n = 0 ; n < 16 ; n++)
                {
                double step2 = n * step;
                double step3 = step2 + ColorStep;
                shps[n].Fill.ForeColor.RGB = CreatColor(step2);
                shps[n].Fill.BackColor.RGB = CreatColor(step3);
                shps[n].Fill.GradientAngle = (float)Gra;
                }
            }
        //删除配色色块
        /// <summary>
        /// 输入图形的名称 string Name
        /// </summary>
        /// <param name="Name">名称</param>
        public void DelShpe(string Name)
            {
            try
                {
                PowerPoint.Slide slide = app.ActiveWindow.View.Slide;

                // 使用 LINQ 查询来获取需要删除的形状
                var shapesToDelete = slide.Shapes.Cast<PowerPoint.Shape>()
                    .Where(shape => shape.Tags["图形"] != null && shape.Tags["图形"].ToString() == "渐变色块")
                    .ToList();

                // 遍历并删除形状
                foreach (PowerPoint.Shape shape in shapesToDelete)
                    {
                    shape.Delete();
                    }
                }
            catch (Exception ex)
                {
                // 处理异常，例如记录日志或者输出错误信息
                Console.WriteLine("Error: " + ex.Message);
                }
            }


        private void CreatBtn_Click(object sender, RoutedEventArgs e)
            {
            double step = Slider1.ValueEnd - Slider1.ValueStart;
            double Gra = Slider2.ValueEnd - Slider2.ValueStart;
            double ColorStep = Slider3.ValueEnd - Slider3.ValueStart;
            CreatShap(step, Gra, ColorStep);
            }

        private void DelBtn_Click(object sender, RoutedEventArgs e)
            {
            DelShpe("渐变色块");
            }

        private void EditBtn_Click(object sender, RoutedEventArgs e)
            {
            Selection sel = app.ActiveWindow.Selection;
            if (sel.Type == PpSelectionType.ppSelectionShapes)
                {
                sel.ShapeRange.Apply();
                }
            else
                {
                }
            }

        private void CopyBtn_Click(object sender, RoutedEventArgs e)
            {
            Selection sel = app.ActiveWindow.Selection;
            try
                {
                if (sel.ShapeRange.Tags["图形"] == "渐变色块" && sel.Type == PpSelectionType.ppSelectionShapes)
                    {
                    sel.ShapeRange.PickUp();
                    }
                else
                    {
                    }
                }
            catch
                {
                }
            }

        private void Slider1_ValueChanged(object sender, RoutedPropertyChangedEventArgs<HandyControl.Data.DoubleRange> e)
            {
           
            CreatShap(step, Gra, ColorStep);
            }

        private void Slider2_ValueChanged(object sender, RoutedPropertyChangedEventArgs<HandyControl.Data.DoubleRange> e)
            {
           
            CreatShap(step, Gra, ColorStep);
            }

        private void Slider3_ValueChanged(object sender, RoutedPropertyChangedEventArgs<HandyControl.Data.DoubleRange> e)
            {
            
            CreatShap(step, Gra, ColorStep);
            }
        }
    }
