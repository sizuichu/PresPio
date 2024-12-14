using HandyControl.Controls;
using System;
using System.Collections.Generic;
using System.Drawing.Imaging;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using System.IO;
using System.Runtime.InteropServices;

namespace PresPio.Function
{
    public class ImageHelper
    {

        public void gitMessage() {

            Growl.SuccessGlobal("测试");
        
        }

        /// <summary>
        /// 图片整体透明
        /// </summary>
        /// <param name="image"></param>
        /// <param name="opacityValue"></param>
        /// <returns></returns>
        public Bitmap SetImageOverallOpacity(Image image, float opacityValue)
            {
            // 将透明度值限制在 0 到 100 之间，并将其转换为比例
            opacityValue = Math.Max(0, Math.Min(100, opacityValue));
            float actualOpacity = 1 - (opacityValue / 100); // 0 为不透明，100 为完全透明

            // 创建一个新的 Bitmap 对象
            Bitmap bmp = new Bitmap(image.Width, image.Height);

            using (Graphics g = Graphics.FromImage(bmp))
                {
                // 创建一个颜色矩阵，用于设置透明度
                ColorMatrix colorMatrix = new ColorMatrix
                    {
                    Matrix33 = actualOpacity // 设置透明度
                    };

                // 创建图像属性
                using (ImageAttributes imageAttributes = new ImageAttributes())
                    {
                    imageAttributes.SetColorMatrix(colorMatrix, ColorMatrixFlag.Default, ColorAdjustType.Bitmap);

                    // 绘制原始图像到新图像，应用透明度
                    g.DrawImage(image, new Rectangle(0, 0, bmp.Width, bmp.Height),
                                0, 0, image.Width, image.Height,
                                GraphicsUnit.Pixel, imageAttributes);
                    }
                }

            return bmp; // 返回处理后的 Bitmap
            }



        /// <summary>
        /// 图片渐变透明
        /// </summary>
        /// <param name="image"></param>
        /// <param name="gradientValue"></param>
        /// <param name="angle"></param>
        /// <returns></returns>
        public Bitmap SetImageGradientOpacity(Image image, float gradientValue, float angle)
            {
            // 创建一个新的 Bitmap 对象
            Bitmap bmp = new Bitmap(image.Width, image.Height);

            // 计算中心点
            int centerX = bmp.Width / 2;
            int centerY = bmp.Height / 2;

            // 计算角度的弧度
            float radians = (float)(angle * (Math.PI / 180));
            float cos = (float)Math.Cos(radians);
            float sin = (float)Math.Sin(radians);
            float maxDistance = (float)Math.Sqrt(centerX * centerX + centerY * centerY);

            // 锁定位图的像素数据
            BitmapData bmpData = bmp.LockBits(new Rectangle(0, 0, bmp.Width, bmp.Height), ImageLockMode.WriteOnly, bmp.PixelFormat);
            int bytesPerPixel = Image.GetPixelFormatSize(bmp.PixelFormat) / 8;

            unsafe
                {
                byte* ptr = (byte*)bmpData.Scan0;

                // 遍历每个像素
                for (int y = 0 ; y < bmp.Height ; y++)
                    {
                    for (int x = 0 ; x < bmp.Width ; x++)
                        {
                        // 计算到中心的距离
                        float deltaX = x - centerX;
                        float deltaY = y - centerY;

                        // 计算与渐变方向的投影
                        float distance = (deltaX * cos + deltaY * sin) / (float)Math.Sqrt(cos * cos + sin * sin);

                        // 计算透明度，范围从 1.0 到 0.0
                        float opacity = 1.0f - Math.Max(0, Math.Min(1, (distance / maxDistance) * (gradientValue / 100)));

                        // 确保透明度在 0 到 1 之间
                        opacity = Math.Max(0, Math.Min(1, opacity));

                        // 获取原始像素颜色
                        Color originalColor = ((Bitmap)image).GetPixel(x, y);
                        byte alpha = (byte)(originalColor.A * opacity);
                        ptr[(y * bmpData.Stride) + (x * bytesPerPixel) + 0] = originalColor.B; // Blue
                        ptr[(y * bmpData.Stride) + (x * bytesPerPixel) + 1] = originalColor.G; // Green
                        ptr[(y * bmpData.Stride) + (x * bytesPerPixel) + 2] = originalColor.R; // Red
                        ptr[(y * bmpData.Stride) + (x * bytesPerPixel) + 3] = alpha; // Alpha
                        }
                    }
                }

            // 解锁位图的像素数据
            bmp.UnlockBits(bmpData);

            return bmp; // 返回渐变透明的 Bitmap
            }





        /// <summary>
        /// 设置图片的透明度
        /// </summary>
        /// <param name="image"></param>
        /// <param name="opacity"></param>
        /// <returns></returns>
        public Bitmap SetImageOpacity(Image image, float opacity)
            {
            // 创建一个新的 Bitmap 对象
            Bitmap bmp = new Bitmap(image.Width, image.Height);

            // 创建 Graphics 对象
            using (Graphics g = Graphics.FromImage(bmp))
                {
                // 创建一个颜色矩阵来调整透明度
                ColorMatrix matrix = new ColorMatrix();
                matrix.Matrix33 = opacity; // 设置透明度

                // 创建图像属性并设置颜色矩阵
                ImageAttributes attributes = new ImageAttributes();
                attributes.SetColorMatrix(matrix, ColorMatrixFlag.Default, ColorAdjustType.Bitmap);

                // 绘制原始图像到新的 Bitmap 上
                g.DrawImage(image, new Rectangle(0, 0, bmp.Width, bmp.Height),
                            0, 0, image.Width, image.Height, GraphicsUnit.Pixel, attributes);
                }

            return bmp; // 返回调整透明度后的 Bitmap
            }

        public Bitmap ChangeImageOpacity(string imagePath, float opacity)
        {
            Bitmap originalImage = new Bitmap(imagePath);

            // 创建一个带有透明通道的新位图
            Bitmap transparentImage = new Bitmap(originalImage.Width, originalImage.Height, PixelFormat.Format32bppArgb);

            // 设置图像的分辨率
            transparentImage.SetResolution(originalImage.HorizontalResolution, originalImage.VerticalResolution);

            // 使用 Graphics 类绘制图像以应用透明度
            using (Graphics g = Graphics.FromImage(transparentImage))
            {
                // 设置透明度矩阵
                ColorMatrix matrix = new ColorMatrix();
                matrix.Matrix33 = opacity; // 0.0表示完全透明，1.0表示完全不透明

                // 创建 ImageAttributes 对象并设置颜色矩阵
                ImageAttributes imageAttributes = new ImageAttributes();
                imageAttributes.SetColorMatrix(matrix, ColorMatrixFlag.Default, ColorAdjustType.Bitmap);

                // 绘制原始图像到带有透明度的位图上
                g.DrawImage(originalImage,
                    new Rectangle(0, 0, originalImage.Width, originalImage.Height),
                    0, 0, originalImage.Width, originalImage.Height,
                    GraphicsUnit.Pixel, imageAttributes);
            }

            return transparentImage;
        }

        public void InsertImageToSlide(Bitmap image, int slideIndex)
        {
            try
            {
                // 获取当前 PowerPoint 应用程序实例
                Application pptApp = Globals.ThisAddIn.Application;

                // 获取当前活动的演示文稿
                Presentation presentation = pptApp.ActivePresentation;

                // 获取要插入图片的幻灯片
                Slide slide = presentation.Slides[slideIndex];

                // 将 Bitmap 转换为 byte[]
                byte[] imageBytes;
                using (MemoryStream stream = new MemoryStream())
                {
                    image.Save(stream, ImageFormat.Png); // 可以根据需要修改保存格式
                    imageBytes = stream.ToArray();
                }

                // 将图片插入到幻灯片中
                slide.Shapes.AddPicture(
                    FileName: "",
                    LinkToFile: MsoTriState.msoFalse,
                    SaveWithDocument: MsoTriState.msoCTrue,
                    Left: 0, Top: 0, Width: image.Width, Height: image.Height
                    ); // 作为第一个形状插入

                // 清理资源
                Marshal.ReleaseComObject(slide);
                Marshal.ReleaseComObject(presentation);
            }
            catch (Exception ex)
            {
                Console.WriteLine("插入图片时发生错误：" + ex.Message);
                // 这里可以添加处理异常的代码
            }
        }

    }
}
