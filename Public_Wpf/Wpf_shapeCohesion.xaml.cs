using System;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Media;
using System.Windows.Shapes;
using HandyControl.Controls;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace PresPio
    {
    /// <summary>
    /// Wpf_shapeCohesion.xaml 的交互逻辑
    /// </summary>
    public partial class Wpf_shapeCohesion : HandyControl.Controls.Window
        {
        public PowerPoint.Application app;
        public Wpf_shapeCohesion()
            {
            InitializeComponent();
            ToggleButton[] toggleButtons = { ToggleButton1, ToggleButton2, ToggleButton3, ToggleButton4, ToggleButton5, ToggleButton6, ToggleButton7, ToggleButton8 };

            // 在初始化代码中为每个ToggleButton添加Checked事件处理程序
            foreach (var child in toggleButtons)
                {
                if (child is ToggleButton toggleButton)
                    {
                    toggleButton.Checked += TogButton_Checked;
                    }
                }
            }

        private void ShpeWindow_Loaded(object sender, RoutedEventArgs e)
            {
            app = Globals.ThisAddIn.Application;
            }

        private void TogButton_Checked(object sender, RoutedEventArgs e)
            {
            ToggleButton clickedButton = sender as ToggleButton;
            ToggleButton[] toggleButtons = { ToggleButton1, ToggleButton2, ToggleButton3, ToggleButton4, ToggleButton5, ToggleButton6, ToggleButton7, ToggleButton8 };
            if (clickedButton.IsChecked == true)
                {
                // 取消其他ToggleButton的选中状态
                foreach (var child in toggleButtons)
                    {
                    if (child is ToggleButton toggleButton && toggleButton != clickedButton)
                        {
                        toggleButton.IsChecked = false;
                        }
                    }
                // 更新预览
                UpdatePreview(clickedButton);
                }
            else
                {
                // 清除预览
                ClearPreview();
                }
            }

        private void UpdatePreview(ToggleButton selectedButton)
            {
            // 清除现有的连接线
            ClearPreview();

            // 获取预览区域中的两个矩形
            Rectangle leftRect = PreviewCanvas.Children[0] as Rectangle;
            Rectangle rightRect = PreviewCanvas.Children[1] as Rectangle;

            // 创建连接线
            Rectangle connector = new Rectangle();
            connector.Fill = new SolidColorBrush(Color.FromRgb(33, 150, 243)); // #2196F3

            // 根据选中的按钮设置连接线的位置和样式
            if (selectedButton == ToggleButton1) // 从左到右
                {
                connector.Width = 110;
                connector.Height = 4;
                Canvas.SetLeft(connector, 140);
                Canvas.SetTop(connector, 28);
                }
            else if (selectedButton == ToggleButton2) // 从右到左
                {
                connector.Width = 110;
                connector.Height = 4;
                Canvas.SetLeft(connector, 140);
                Canvas.SetTop(connector, 28);
                }
            else if (selectedButton == ToggleButton3) // 左边垂直连接
                {
                connector.Width = 4;
                connector.Height = 40;
                Canvas.SetLeft(connector, 100);
                Canvas.SetTop(connector, 10);
                }
            else if (selectedButton == ToggleButton4) // 右边垂直连接
                {
                connector.Width = 4;
                connector.Height = 40;
                Canvas.SetLeft(connector, 290);
                Canvas.SetTop(connector, 10);
                }
            else if (selectedButton == ToggleButton5) // 左上到右下
                {
                connector.Width = 155;
                connector.Height = 4;
                Canvas.SetLeft(connector, 118);
                Canvas.SetTop(connector, 28);
                RotateTransform transform = new RotateTransform(45, 77.5, 2);
                connector.RenderTransform = transform;
                }
            else if (selectedButton == ToggleButton6) // 左下到右上
                {
                connector.Width = 155;
                connector.Height = 4;
                Canvas.SetLeft(connector, 118);
                Canvas.SetTop(connector, 28);
                RotateTransform transform = new RotateTransform(-45, 77.5, 2);
                connector.RenderTransform = transform;
                }
            else if (selectedButton == ToggleButton7) // 顶部连接
                {
                connector.Width = 110;
                connector.Height = 4;
                Canvas.SetLeft(connector, 140);
                Canvas.SetTop(connector, 10);
                }
            else if (selectedButton == ToggleButton8) // 底部连接
                {
                connector.Width = 110;
                connector.Height = 4;
                Canvas.SetLeft(connector, 140);
                Canvas.SetTop(connector, 46);
                }

            // 添加连接线到预览画布
            PreviewCanvas.Children.Add(connector);
            }

        private void ClearPreview()
            {
            // 移除所有连接线，保留两个矩形
            while (PreviewCanvas.Children.Count > 2)
                {
                PreviewCanvas.Children.RemoveAt(2);
                }
            }

        /// <summary>
        /// 形状创建函数
        /// </summary>
        /// <param name="ShapeDirection"></param>
        public void ConnectRectangleShapes(string ShapeDirection)
            {
            Selection sel = app.ActiveWindow.Selection;
            if (sel.Type == PpSelectionType.ppSelectionShapes && sel.ShapeRange.Count == 2)
                {
                float Left1, Right1, Top1, Bottom1, Left2, Right2, Top2, Bottom2;

                Left1 = sel.ShapeRange[1].Left;
                Right1 = Left1 + sel.ShapeRange[1].Width;
                Top1 = sel.ShapeRange[1].Top;
                Bottom1 = Top1 + sel.ShapeRange[1].Height;

                Left2 = sel.ShapeRange[2].Left;
                Right2 = Left2 + sel.ShapeRange[2].Width;
                Top2 = sel.ShapeRange[2].Top;
                Bottom2 = Top2 + sel.ShapeRange[2].Height;
                SlideRange myDocument = sel.SlideRange;
                switch (ShapeDirection)
                    {
                    case "1":
                            {
                                {
                                var withBlock = myDocument.Shapes.BuildFreeform(EditingType: MsoEditingType.msoEditingCorner, Left1, Top1);
                                withBlock.AddNodes(SegmentType: MsoSegmentType.msoSegmentLine, EditingType: MsoEditingType.msoEditingAuto, Left1, Bottom1);
                                withBlock.AddNodes(SegmentType: MsoSegmentType.msoSegmentLine, EditingType: MsoEditingType.msoEditingAuto, Right2, Bottom2);
                                withBlock.AddNodes(SegmentType: MsoSegmentType.msoSegmentLine, EditingType: MsoEditingType.msoEditingAuto, Right2, Top2);
                                withBlock.AddNodes(SegmentType: MsoSegmentType.msoSegmentLine, EditingType: MsoEditingType.msoEditingAuto, Left1, Top1);
                                withBlock.ConvertToShape().Line.Visible = MsoTriState.msoFalse;
                                }

                            break;
                            }
                    case "2":
                            {
                                {
                                var withBlock = myDocument.Shapes.BuildFreeform(EditingType: MsoEditingType.msoEditingCorner, Right1, Top1);
                                withBlock.AddNodes(SegmentType: MsoSegmentType.msoSegmentLine, EditingType: MsoEditingType.msoEditingAuto, Right1, Bottom1);
                                withBlock.AddNodes(SegmentType: MsoSegmentType.msoSegmentLine, EditingType: MsoEditingType.msoEditingAuto, Left2, Bottom2);
                                withBlock.AddNodes(SegmentType: MsoSegmentType.msoSegmentLine, EditingType: MsoEditingType.msoEditingAuto, Left2, Top2);
                                withBlock.AddNodes(SegmentType: MsoSegmentType.msoSegmentLine, EditingType: MsoEditingType.msoEditingAuto, Right1, Top1);
                                withBlock.ConvertToShape().Line.Visible = MsoTriState.msoFalse;
                                }

                            break;
                            }
                    case "3":
                            {
                                {
                                var withBlock = myDocument.Shapes.BuildFreeform(EditingType: MsoEditingType.msoEditingCorner, Left1, Top1);
                                withBlock.AddNodes(SegmentType: MsoSegmentType.msoSegmentLine, EditingType: MsoEditingType.msoEditingAuto, Left1, Bottom1);
                                withBlock.AddNodes(SegmentType: MsoSegmentType.msoSegmentLine, EditingType: MsoEditingType.msoEditingAuto, Left2, Bottom2);
                                withBlock.AddNodes(SegmentType: MsoSegmentType.msoSegmentLine, EditingType: MsoEditingType.msoEditingAuto, Left2, Top2);
                                withBlock.AddNodes(SegmentType: MsoSegmentType.msoSegmentLine, EditingType: MsoEditingType.msoEditingAuto, Left1, Top1);
                                withBlock.ConvertToShape().Line.Visible = MsoTriState.msoFalse;
                                }

                            break;
                            }
                    case "4":
                            {
                                {
                                var withBlock = myDocument.Shapes.BuildFreeform(EditingType: MsoEditingType.msoEditingCorner, Right1, Top1);
                                withBlock.AddNodes(SegmentType: MsoSegmentType.msoSegmentLine, EditingType: MsoEditingType.msoEditingAuto, Right1, Bottom1);
                                withBlock.AddNodes(SegmentType: MsoSegmentType.msoSegmentLine, EditingType: MsoEditingType.msoEditingAuto, Right2, Bottom2);
                                withBlock.AddNodes(SegmentType: MsoSegmentType.msoSegmentLine, EditingType: MsoEditingType.msoEditingAuto, Right2, Top2);
                                withBlock.AddNodes(SegmentType: MsoSegmentType.msoSegmentLine, EditingType: MsoEditingType.msoEditingAuto, Right1, Top1);
                                withBlock.ConvertToShape().Line.Visible = MsoTriState.msoFalse;
                                }

                            break;
                            }
                    case "5":
                            {
                                {
                                var withBlock = myDocument.Shapes.BuildFreeform(EditingType: MsoEditingType.msoEditingCorner, Left1, Top1);
                                withBlock.AddNodes(SegmentType: MsoSegmentType.msoSegmentLine, EditingType: MsoEditingType.msoEditingAuto, Left2, Bottom2);
                                withBlock.AddNodes(SegmentType: MsoSegmentType.msoSegmentLine, EditingType: MsoEditingType.msoEditingAuto, Right2, Bottom2);
                                withBlock.AddNodes(SegmentType: MsoSegmentType.msoSegmentLine, EditingType: MsoEditingType.msoEditingAuto, Right1, Top1);
                                withBlock.AddNodes(SegmentType: MsoSegmentType.msoSegmentLine, EditingType: MsoEditingType.msoEditingAuto, Left1, Top1);
                                withBlock.ConvertToShape().Line.Visible = MsoTriState.msoFalse;
                                }

                            break;
                            }
                    case "6":
                            {
                                {
                                var withBlock = myDocument.Shapes.BuildFreeform(EditingType: MsoEditingType.msoEditingCorner, Left1, Bottom1);
                                withBlock.AddNodes(SegmentType: MsoSegmentType.msoSegmentLine, EditingType: MsoEditingType.msoEditingAuto, Left2, Top2);
                                withBlock.AddNodes(SegmentType: MsoSegmentType.msoSegmentLine, EditingType: MsoEditingType.msoEditingAuto, Right2, Top2);
                                withBlock.AddNodes(SegmentType: MsoSegmentType.msoSegmentLine, EditingType: MsoEditingType.msoEditingAuto, Right1, Bottom1);
                                withBlock.AddNodes(SegmentType: MsoSegmentType.msoSegmentLine, EditingType: MsoEditingType.msoEditingAuto, Left1, Bottom1);
                                withBlock.ConvertToShape().Line.Visible = MsoTriState.msoFalse;
                                }

                            break;
                            }
                    case "7":
                            {
                                {
                                var withBlock = myDocument.Shapes.BuildFreeform(EditingType: MsoEditingType.msoEditingCorner, Left1, Top1);
                                withBlock.AddNodes(SegmentType: MsoSegmentType.msoSegmentLine, EditingType: MsoEditingType.msoEditingAuto, Left2, Top2);
                                withBlock.AddNodes(SegmentType: MsoSegmentType.msoSegmentLine, EditingType: MsoEditingType.msoEditingAuto, Right2, Top2);
                                withBlock.AddNodes(SegmentType: MsoSegmentType.msoSegmentLine, EditingType: MsoEditingType.msoEditingAuto, Right1, Top1);
                                withBlock.AddNodes(SegmentType: MsoSegmentType.msoSegmentLine, EditingType: MsoEditingType.msoEditingAuto, Left1, Top1);
                                withBlock.ConvertToShape().Line.Visible = MsoTriState.msoFalse;
                                }

                            break;
                            }
                    case "8":
                            {
                                {
                                var withBlock = myDocument.Shapes.BuildFreeform(EditingType: MsoEditingType.msoEditingCorner, Left1, Bottom1);
                                withBlock.AddNodes(SegmentType: MsoSegmentType.msoSegmentLine, EditingType: MsoEditingType.msoEditingAuto, Left2, Bottom2);
                                withBlock.AddNodes(SegmentType: MsoSegmentType.msoSegmentLine, EditingType: MsoEditingType.msoEditingAuto, Right2, Bottom2);
                                withBlock.AddNodes(SegmentType: MsoSegmentType.msoSegmentLine, EditingType: MsoEditingType.msoEditingAuto, Right1, Bottom1);
                                withBlock.AddNodes(SegmentType: MsoSegmentType.msoSegmentLine, EditingType: MsoEditingType.msoEditingAuto, Left1, Bottom1);
                                withBlock.ConvertToShape().Line.Visible = MsoTriState.msoFalse;
                                }

                            break;
                            }


                    }


                }
            else
                Growl.Warning("请选择两个形状");
            }

        private void OkBtn_Click(object sender, RoutedEventArgs e)
            {
            if (ToggleButton1.IsChecked == true)//从左到右
                {
                ConnectRectangleShapes("1");
                }
            else if (ToggleButton2.IsChecked == true)
                {
                ConnectRectangleShapes("2");
                }
            else if (ToggleButton3.IsChecked == true)
                {
                ConnectRectangleShapes("3");
                }
            else if (ToggleButton4.IsChecked == true)
                {
                ConnectRectangleShapes("4");
                }
            else if (ToggleButton5.IsChecked == true)
                {
                ConnectRectangleShapes("5");
                }
            else if (ToggleButton6.IsChecked == true)
                {
                ConnectRectangleShapes("6");
                }
            else if (ToggleButton7.IsChecked == true)
                {
                ConnectRectangleShapes("7");
                }
            else
                {
                ConnectRectangleShapes("8");
                }
            }
        }
    }
