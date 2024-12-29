using HandyControl.Controls;
using System;

using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace PresPio
    {
    public class ComHelper
        {
        public PowerPoint.Application app = Globals.ThisAddIn.Application;
        /// <summary>
        /// 根据所选内容在不同方向进行复制
        /// </summary>
        /// <param name="sel"></param>
        /// <param name="direction"></param>
        #region

        public void DuplicateShapes(string direction)
            {
            PowerPoint.Application app = Globals.ThisAddIn.Application;
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

                    switch (direction.ToLower())
                        {
                        case "left":
                            cshape.Left = shape.Left - shape.Width;
                            cshape.Top = shape.Top;
                            break;

                        case "right":
                            cshape.Left = shape.Left + shape.Width;
                            cshape.Top = shape.Top;
                            break;

                        case "up":
                            cshape.Left = shape.Left;
                            cshape.Top = shape.Top - shape.Height;
                            break;

                        case "down":
                            cshape.Left = shape.Left;
                            cshape.Top = shape.Top + shape.Height;
                            break;

                        case "leftup":
                            cshape.Left = shape.Left - shape.Width;
                            cshape.Top = shape.Top - shape.Height;
                            break;

                        case "leftdown":
                            cshape.Left = shape.Left - shape.Width;
                            cshape.Top = shape.Top + shape.Height;
                            break;

                        case "rightup":
                            cshape.Left = shape.Left + shape.Width;
                            cshape.Top = shape.Top - shape.Height;
                            break;

                        case "rightdown":
                            cshape.Left = shape.Left + shape.Width;
                            cshape.Top = shape.Top + shape.Height;
                            break;

                        default:
                            Growl.Warning("无效的方向", "错误");
                            return;
                        }

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
            }

        #endregion

        /// <summary>
        /// 矩阵排列
        /// </summary>
        /// <param name="rows"></param>
        /// <param name="cols"></param>
        #region

        public void ArrangeShapesInGrid(int rows, int cols)
            {
            PowerPoint.Application app = Globals.ThisAddIn.Application;
            PowerPoint.Selection sel = app.ActiveWindow.Selection;

            if (sel.Type == PowerPoint.PpSelectionType.ppSelectionNone)
                {
                Growl.Warning("请选择一个元素", "温馨提示");
                return;
                }

            if (sel.Type != PowerPoint.PpSelectionType.ppSelectionShapes)
                {
                Growl.Warning("请选择形状元素", "温馨提示");
                return;
                }

            PowerPoint.ShapeRange range = sel.ShapeRange;
            int count = range.Count;

            // 计算每个形状之间的水平和垂直间距
            float shapeWidth = range[1].Width;
            float shapeHeight = range[1].Height;
            float horizontalSpacing = shapeWidth * 1.5f; // 提供一些间隔，视需求而定
            float verticalSpacing = shapeHeight * 1.5f;

            int index = 1;
            for (int row = 0 ; row < rows ; row++)
                {
                for (int col = 0 ; col < cols ; col++)
                    {
                    if (index > count)
                        {
                        return; // 如果形状已排列完毕，直接返回
                        }

                    PowerPoint.Shape shape = range[index];
                    shape.Left = col * horizontalSpacing;
                    shape.Top = row * verticalSpacing;

                    index++;
                    }
                }
            }

        #endregion

        /// <summary>
        /// 环形排列,可以选择是否旋转角度
        /// </summary>
        /// <param name="numRows"></param>
        /// <param name="numCols"></param>
        #region

        public void ArrangeShapesInCircle(float radius, bool rotate)
            {
            PowerPoint.Application app = Globals.ThisAddIn.Application;
            PowerPoint.Selection sel = app.ActiveWindow.Selection;
            PowerPoint.Slide slide = app.ActiveWindow.View.Slide;

            if (sel.Type == PowerPoint.PpSelectionType.ppSelectionNone)
                {
                Growl.Warning("请选择一个元素", "温馨提示");
                return;
                }

            if (sel.Type != PowerPoint.PpSelectionType.ppSelectionShapes)
                {
                Growl.Warning("请选择形状元素", "温馨提示");
                return;
                }

            PowerPoint.ShapeRange range = sel.ShapeRange;
            int count = range.Count;

            if (count < 2)
                {
                Growl.Warning("至少选择两个形状", "温馨提示");
                return;
                }

            // 圆心坐标（幻灯片中心）
            float centerX = slide.Master.Width / 2;
            float centerY = slide.Master.Height / 2;

            // 计算每个形状的角度步进
            float angleStep = 360f / count;

            // 存储所有排列形状的名称
            string[] shapeNames = new string[count];

            for (int i = 1 ; i <= count ; i++)
                {
                PowerPoint.Shape shape = range[i];
                // 计算当前形状的角度
                float angle = i * angleStep;
                // 转换为弧度
                float radians = (float)(angle * Math.PI / 180);

                // 计算形状的位置
                float x = (float)(radius * Math.Cos(radians));
                float y = (float)(radius * Math.Sin(radians));

                // 设置形状的新位置
                shape.Left = centerX + x - shape.Width / 2;
                shape.Top = centerY + y - shape.Height / 2;

                // 如果需要旋转
                if (rotate)
                    {
                    shape.Rotation = angle; // 将角度设置为形状的旋转角度
                    }
                else
                    {
                    shape.Rotation = 0; // 不旋转形状
                    }

                // 存储形状名称
                shapeNames[i - 1] = shape.Name;
                }

            // 选择所有排列后的形状
            if (shapeNames.Length > 0)
                {
                slide.Shapes.Range(shapeNames).Select();
                }
            }

        #endregion
        }
    }