using System;
using System.Drawing;
using System.Drawing.Text;
using System.IO;
using System.Runtime.InteropServices;

namespace PresPio
    {
    public class RibbonIcoHelper
        {
        private static Font segoeFluentFont;
        // 从应用程序设置中获取图标颜色
        public Color iconColor = Properties.Settings.Default.iconColor;

        /// <summary>
        /// 加载字体
        /// </summary>
        private static void LoadFont(float fontSize)
            {
            try
                {
                // 从资源文件加载字体
                byte[] fontData = Properties.Resources.FluentIcons;
                if (fontData == null || fontData.Length == 0)
                    {
                    throw new ArgumentNullException(nameof(fontData), "字体数据不能为空，请确保在资源文件中添加了名为'FluentIcons'的字体资源");
                    }

                PrivateFontCollection fontCollection = new PrivateFontCollection();
                IntPtr fontPtr = Marshal.AllocCoTaskMem(fontData.Length);
                Marshal.Copy(fontData, 0, fontPtr, fontData.Length);
                fontCollection.AddMemoryFont(fontPtr, fontData.Length);
                Marshal.FreeCoTaskMem(fontPtr);

                if (fontCollection.Families.Length == 0)
                    {
                    throw new InvalidOperationException("无法加载字体文件");
                    }

                segoeFluentFont = new Font(fontCollection.Families[0], fontSize);
                }
            catch (Exception ex)
                {
                throw new InvalidOperationException($"加载字体时出错: {ex.Message}", ex);
                }
            }

        /// <summary>
        /// 绘制 Segoe Fluent 图标并生成图像
        /// </summary>
        private static Bitmap GetIconImage(string iconUnicode, int imageWidth, int imageHeight, Color fontColor, Color backgroundColor)
            {
            if (segoeFluentFont == null)
                {
                throw new InvalidOperationException("字体未加载，请先调用 LoadFont 方法");
                }

            using (Bitmap iconImage = new Bitmap(imageWidth, imageHeight))
                {
                using (Graphics g = Graphics.FromImage(iconImage))
                    {
                    // 设置抗锯齿
                    g.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.AntiAlias;
                    g.TextRenderingHint = System.Drawing.Text.TextRenderingHint.ClearTypeGridFit;

                    // 填充背景
                    using (Brush backgroundBrush = new SolidBrush(backgroundColor))
                        {
                        g.FillRectangle(backgroundBrush, 0, 0, imageWidth, imageHeight);
                        }

                    // 计算图标位置
                    SizeF iconSize = g.MeasureString(iconUnicode, segoeFluentFont);
                    float x = (imageWidth - iconSize.Width) / 2;
                    float y = (imageHeight - iconSize.Height) / 2;

                    // 绘制图标
                    using (Brush fontBrush = new SolidBrush(fontColor))
                        {
                        g.DrawString(iconUnicode, segoeFluentFont, fontBrush, new PointF(x, y));
                        }
                    }
                return (Bitmap)iconImage.Clone();
                }
            }

        // 根据应用程序名称获取图标
        public void GetIcons(string appName)
            {
            // 获取功能区的实例
            MyRibbon myRibbon = Globals.Ribbons.Ribbon1;

            try
                {
                if (appName == "PPT")
                    {
                    // 获取屏幕的 DPI
                    using (Graphics g = Graphics.FromHwnd(IntPtr.Zero))
                        {
                        float dpiX = g.DpiX;
                        float dpiY = g.DpiY;

                        // 根据 DPI 设置字体大小和图标尺寸
                        float baseFontSize = 72f;
                        float fontSize = baseFontSize * (dpiX / 96f); // 96 DPI 是标准 DPI
                        int baseImageSize = 120;
                        int imageWidth = (int)(baseImageSize * (dpiX / 96f));
                        int imageHeight = (int)(baseImageSize * (dpiY / 96f));

                        // 加载字体
                        LoadFont(fontSize);

                        // 设置字体颜色和背景颜色
                        Color fontColor = Color.FromArgb(iconColor.A, iconColor.R, iconColor.G, iconColor.B);
                        Color backgroundColor = Color.Transparent;

                        // 设置功能区中的图标
                        myRibbon.menu1.Image = GetIconImage("\ue9cf", imageWidth, imageHeight, fontColor, backgroundColor);
                        myRibbon.menu2.Image = GetIconImage("\uf18e", imageWidth, imageHeight, fontColor, backgroundColor);
                        myRibbon.menu3.Image = GetIconImage("\uF0B6", imageWidth, imageHeight, fontColor, backgroundColor);
                        myRibbon.menu4.Image = GetIconImage("\uee41", imageWidth, imageHeight, fontColor, backgroundColor);
                        myRibbon.menu5.Image = GetIconImage("\uebb8", imageWidth, imageHeight, fontColor, backgroundColor);
                        myRibbon.menu6.Image = GetIconImage("\uea4e", imageWidth, imageHeight, fontColor, backgroundColor);
                        myRibbon.menu7.Image = GetIconImage("\ue579", imageWidth, imageHeight, fontColor, backgroundColor);
                        myRibbon.menu8.Image = GetIconImage("\uee79", imageWidth, imageHeight, fontColor, backgroundColor);
                        myRibbon.menu9.Image = GetIconImage("\uf197", imageWidth, imageHeight, fontColor, backgroundColor);
                        myRibbon.menu10.Image = GetIconImage("\ue7b7", imageWidth, imageHeight, fontColor, backgroundColor);
                        myRibbon.menu11.Image = GetIconImage("\uee69", imageWidth, imageHeight, fontColor, backgroundColor);
                        myRibbon.menu19.Image = GetIconImage("\ueabb", imageWidth, imageHeight, fontColor, backgroundColor);
                        myRibbon.menu21.Image = GetIconImage("\ue706", imageWidth, imageHeight, fontColor, backgroundColor);
                        myRibbon.button3.Image = GetIconImage("\uefbb", imageWidth, imageHeight, fontColor, backgroundColor);
                        myRibbon.button4.Image = GetIconImage("\uf22c", imageWidth, imageHeight, fontColor, backgroundColor);
                        myRibbon.button23.Image = GetIconImage("\ue7f8", imageWidth, imageHeight, fontColor, backgroundColor);
                        myRibbon.button73.Image = GetIconImage("\ue7c4", imageWidth, imageHeight, fontColor, backgroundColor);
                        myRibbon.button92.Image = GetIconImage("\uee4b", imageWidth, imageHeight, fontColor, backgroundColor);
                        myRibbon.button94.Image = GetIconImage("\ue478", imageWidth, imageHeight, fontColor, backgroundColor);
                        myRibbon.button143.Image = GetIconImage("\uFA8D", imageWidth, imageHeight, fontColor, backgroundColor);
                        myRibbon.button144.Image = GetIconImage("\uEE0B", imageWidth, imageHeight, fontColor, backgroundColor);
                        myRibbon.button145.Image = GetIconImage("\uEBC7", imageWidth, imageHeight, fontColor, backgroundColor);
                        myRibbon.button146.Image = GetIconImage("\uFA8F", imageWidth, imageHeight, fontColor, backgroundColor);
                        myRibbon.toggleButton4.Image = GetIconImage("\uf138", imageWidth, imageHeight, fontColor, backgroundColor);
                        myRibbon.splitButton15.Image = GetIconImage("\uf0d0", imageWidth, imageHeight, fontColor, backgroundColor);
                        myRibbon.splitButton1.Image = GetIconImage("\ue82b", imageWidth, imageHeight, fontColor, backgroundColor);
                        myRibbon.splitButton2.Image = GetIconImage("\uf021", imageWidth, imageHeight, fontColor, backgroundColor);
                        myRibbon.splitButton3.Image = GetIconImage("\ue5e7", imageWidth, imageHeight, fontColor, backgroundColor);
                        myRibbon.splitButton4.Image = GetIconImage("\ue334", imageWidth, imageHeight, fontColor, backgroundColor);
                        myRibbon.splitButton5.Image = GetIconImage("\uea1e", imageWidth, imageHeight, fontColor, backgroundColor);
                        myRibbon.splitButton6.Image = GetIconImage("\uef87", imageWidth, imageHeight, fontColor, backgroundColor);
                        myRibbon.splitButton7.Image = GetIconImage("\ue958", imageWidth, imageHeight, fontColor, backgroundColor);
                        myRibbon.splitButton8.Image = GetIconImage("\ue817", imageWidth, imageHeight, fontColor, backgroundColor);
                        myRibbon.splitButton13.Image = GetIconImage("\ue7ab", imageWidth, imageHeight, fontColor, backgroundColor);
                        myRibbon.splitButton12.Image = GetIconImage("\uEC14", imageWidth, imageHeight, fontColor, backgroundColor);
                        myRibbon.splitButton16.Image = GetIconImage("\ue434", imageWidth, imageHeight, fontColor, backgroundColor);
                        }
                    }
                }
            catch (Exception ex)
                {
                System.Windows.Forms.MessageBox.Show($"加载图标时出错: {ex.Message}", "错误", 
                    System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
                }
            }
        }
    }