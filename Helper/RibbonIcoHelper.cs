using System;
using System.Drawing;
using System.IO;
using DocumentFormat.OpenXml.Office2016.Drawing.Command;
using IconUtility;
using Microsoft.Office.Interop.Word;
using Svg;

namespace PresPio
    {
    public class RibbonIcoHelper
        {
        // 从应用程序设置中获取图标颜色
        public Color iconColor = Properties.Settings.Default.iconColor;

        // 根据应用程序名称获取图标
        public void GetIcons(string appName)
            {
            // 获取功能区的实例
            MyRibbon myRibbon = Globals.Ribbons.Ribbon1;

            // 假设 Segoe Fluent ttf 文件名为 "SegoeFluent.ttf"，并已添加到资源文件中
            byte[] fontData = Properties.Resources.Fluents; // "SegoeIcons" 是资源文件中 ttf 文件的名称

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
                    int baseImageSize =120;
                    int imageWidth = (int)(baseImageSize * (dpiX / 96f));
                    int imageHeight = (int)(baseImageSize * (dpiY / 96f));

                    // 加载字体
                    IconHelper.LoadFont(fontData, fontSize);

                    // 设置字体颜色和背景颜色
                    Color fontColor = Color.FromArgb(iconColor.A, iconColor.R, iconColor.G, iconColor.B);
                    Color backgroundColor = Color.Transparent;

                    // 设置功能区中的图标
                    myRibbon.menu1.Image = IconHelper.GetIconImage("\ue9cf", imageWidth, imageHeight, fontColor, backgroundColor);
                    myRibbon.menu2.Image = IconHelper.GetIconImage("\uf18e", imageWidth, imageHeight, fontColor, backgroundColor);
                    myRibbon.menu3.Image = IconHelper.GetIconImage("\uF0B6", imageWidth, imageHeight, fontColor, backgroundColor);
                    myRibbon.menu4.Image = IconHelper.GetIconImage("\uee41", imageWidth, imageHeight, fontColor, backgroundColor);
                    myRibbon.menu5.Image = IconHelper.GetIconImage("\uebb8", imageWidth, imageHeight, fontColor, backgroundColor);
                    myRibbon.menu6.Image = IconHelper.GetIconImage("\uea4e", imageWidth, imageHeight, fontColor, backgroundColor);
                    myRibbon.menu7.Image = IconHelper.GetIconImage("\ue579", imageWidth, imageHeight, fontColor, backgroundColor);
                    myRibbon.menu8.Image = IconHelper.GetIconImage("\uee79", imageWidth, imageHeight, fontColor, backgroundColor);
                    myRibbon.menu9.Image = IconHelper.GetIconImage("\uf197", imageWidth, imageHeight, fontColor, backgroundColor);
                    myRibbon.menu10.Image = IconHelper.GetIconImage("\ue7b7", imageWidth, imageHeight, fontColor, backgroundColor);
                    myRibbon.menu11.Image = IconHelper.GetIconImage("\uee69", imageWidth, imageHeight, fontColor, backgroundColor);
                    myRibbon.menu19.Image = IconHelper.GetIconImage("\ueabb", imageWidth, imageHeight, fontColor, backgroundColor);
                    myRibbon.menu21.Image = IconHelper.GetIconImage("\ue706", imageWidth, imageHeight, fontColor, backgroundColor);
                     myRibbon.button3.Image = IconHelper.GetIconImage("\uefbb", imageWidth, imageHeight, fontColor, backgroundColor);
                    myRibbon.button4.Image = IconHelper.GetIconImage("\uf22c", imageWidth, imageHeight, fontColor, backgroundColor);
                    myRibbon.button23.Image = IconHelper.GetIconImage("\ue7f8", imageWidth, imageHeight, fontColor, backgroundColor);
                    myRibbon.button73.Image = IconHelper.GetIconImage("\ue7c4", imageWidth, imageHeight, fontColor, backgroundColor);
                    myRibbon.button92.Image = IconHelper.GetIconImage("\uee4b", imageWidth, imageHeight, fontColor, backgroundColor);
                    myRibbon.button94.Image = IconHelper.GetIconImage("\ue478", imageWidth, imageHeight, fontColor, backgroundColor);
                    myRibbon.button143.Image = IconHelper.GetIconImage("\uFA8D", imageWidth, imageHeight, fontColor, backgroundColor);
                    myRibbon.button144.Image = IconHelper.GetIconImage("\uEE0B", imageWidth, imageHeight, fontColor, backgroundColor);
                    myRibbon.button145.Image = IconHelper.GetIconImage("\uEBC7", imageWidth, imageHeight, fontColor, backgroundColor);
                    myRibbon.button146.Image = IconHelper.GetIconImage("\uFA8F", imageWidth, imageHeight, fontColor, backgroundColor);
                      myRibbon.toggleButton4.Image = IconHelper.GetIconImage("\uf138", imageWidth, imageHeight, fontColor, backgroundColor);
                    myRibbon.splitButton15.Image = IconHelper.GetIconImage("\uf0d0", imageWidth, imageHeight, fontColor, backgroundColor);
                    myRibbon.splitButton1.Image = IconHelper.GetIconImage("\ue82b", imageWidth, imageHeight, fontColor, backgroundColor);
                    myRibbon.splitButton2.Image = IconHelper.GetIconImage("\uf021", imageWidth, imageHeight, fontColor, backgroundColor);
                    myRibbon.splitButton3.Image = IconHelper.GetIconImage("\ue5e7", imageWidth, imageHeight, fontColor, backgroundColor);
                    myRibbon.splitButton4.Image = IconHelper.GetIconImage("\ue334", imageWidth, imageHeight, fontColor, backgroundColor);
                    myRibbon.splitButton5.Image = IconHelper.GetIconImage("\uea1e", imageWidth, imageHeight, fontColor, backgroundColor);
                    myRibbon.splitButton6.Image = IconHelper.GetIconImage("\uef87", imageWidth, imageHeight, fontColor, backgroundColor);
                    myRibbon.splitButton7.Image = IconHelper.GetIconImage("\ue958", imageWidth, imageHeight, fontColor, backgroundColor);
                    myRibbon.splitButton8.Image = IconHelper.GetIconImage("\ue817", imageWidth, imageHeight, fontColor, backgroundColor);
                       myRibbon.splitButton13.Image = IconHelper.GetIconImage("\ue7ab", imageWidth, imageHeight, fontColor, backgroundColor);
                    myRibbon.splitButton12.Image = IconHelper.GetIconImage("\uEC14", imageWidth, imageHeight, fontColor, backgroundColor);
                      myRibbon.splitButton16.Image = IconHelper.GetIconImage("\ue434", imageWidth, imageHeight, fontColor, backgroundColor);
                 
                    }
                }
            }

        }


}
