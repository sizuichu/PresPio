using HandyControl.Controls;
using Microsoft.Office.Interop.PowerPoint;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Forms;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Media.Media3D;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.Xml.Linq;
using static System.Net.WebRequestMethods;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
namespace PresPio
    {
    /// <summary>
    /// Wpf_superGuide.xaml 的交互逻辑
    /// </summary>
    public partial class Wpf_superGuide
        {
        public Wpf_superGuide()
            {
            InitializeComponent();
            imposValue(); //设置按钮的最大最小值
            AttachValueChangedEventHandlers();//注册按钮事件
            }

        public PowerPoint.Application app;



        private void GrideWindow_Loaded(object sender, RoutedEventArgs e)
            {
            app = Globals.ThisAddIn.Application;
           LoadXml(0);
            }

        /// <summary>
        /// 初始化数据
        /// </summary>
        public void LoadXml(int index)
            {
            var app = Globals.ThisAddIn.Application;
            Assembly assembly = Assembly.GetExecutingAssembly();
            string resourceName = "PresPio.UserData.GuideData.xml";

            try
                {
                using (Stream stream = assembly.GetManifestResourceStream(resourceName))
                    {
                    if (stream == null)
                        {
                        Growl.ErrorGlobal("无法找到嵌入资源，请检查资源名称是否正确。");
                        return;
                        }

                    XDocument xdoc = XDocument.Load(stream);
                    var schemes = xdoc.Descendants("Scheme").ToList();

                    // 清空 GrideListBox 中现有的项
                    GrideListBox.Items.Clear();

                    // 加载所有方案的名称到 GrideListBox 中
                    foreach (var scheme in schemes)
                        {
                        string name = scheme.Element("Name")?.Value;
                        GrideListBox.Items.Add(name); // 将方案名称添加到 GrideListBox
                        }

                    // 设置选中的项（确保索引有效）
                    if (index >= 0 && index < GrideListBox.Items.Count)
                        {
                        GrideListBox.SelectedIndex = index;
                        string selectedSchemeName = GrideListBox.SelectedItem as string;
                        LoadXmlAndUpdateByName(selectedSchemeName); // 根据选择的名称更新 UI
                        }

                   // Growl.SuccessGlobal("XML 文件已成功加载并处理。");
                    }
                }
            catch (Exception ex)
                {
                Growl.ErrorGlobal($"处理 XML 文件时发生错误：{ex.Message}");
                }
            }

        /// <summary>
        /// 根据名称查询并更新数据
        /// </summary>
        public void LoadXmlAndUpdateByName(string searchName)
            {
            var app = Globals.ThisAddIn.Application;
            Assembly assembly = Assembly.GetExecutingAssembly();
            string resourceName = "PresPio.UserData.GuideData.xml";

            try
                {
                using (Stream stream = assembly.GetManifestResourceStream(resourceName))
                    {
                    if (stream == null)
                        {
                        Growl.ErrorGlobal("无法找到嵌入资源，请检查资源名称是否正确。");
                        return;
                        }

                    XDocument xdoc = XDocument.Load(stream);
                    var schemes = xdoc.Descendants("Scheme").ToList();

                    bool found = false; // 用来标识是否找到匹配项

                    // 查找匹配的 Name
                    foreach (var scheme in schemes)
                        {
                        string name = scheme.Element("Name")?.Value;

                        if (name == searchName)
                            {
                            // 获取对应的其他数据
                            int topDistance = GetIntElementValue(scheme, "TopDistance");
                            int bottomDistance = GetIntElementValue(scheme, "BottomDistance");
                            int leftDistance = GetIntElementValue(scheme, "LeftDistance");
                            int rightDistance = GetIntElementValue(scheme, "RightDistance");
                            int rowCount = GetIntElementValue(scheme, "RowCount");
                            int columnCount = GetIntElementValue(scheme, "ColumnCount");

                            // 更新 UI 组件
                            UpdateUI(name, topDistance, bottomDistance, leftDistance, rightDistance, rowCount, columnCount);

                            // 标记已找到匹配项
                            found = true;
                            break; // 找到后退出循环
                            }
                        }

                    if (!found)
                        {
                        Growl.ErrorGlobal($"未找到名称为 '{searchName}' 的方案。");
                        }

                  //  Growl.SuccessGlobal("XML 文件已成功加载并处理。");
                    }
                }
            catch (Exception ex)
                {
                Growl.ErrorGlobal($"处理 XML 文件时发生错误：{ex.Message}");
                }
            }

        /// <summary>
        /// 获取 XML 元素值的辅助方法，确保返回合理的整数值
        /// </summary>
        private int GetIntElementValue(XElement element, string name)
            {
            int result = 0;
            var el = element.Element(name);
            if (el != null && int.TryParse(el.Value, out result))
                {
                return result;
                }
            return result; // 返回默认值 0
            }

        /// <summary>
        /// 处理 GrideListBox 选中项变化
        /// </summary>
        private void GrideListBox_SelectedIndexChanged(object sender, SelectionChangedEventArgs e)
            {
            string scheme = GrideListBox.SelectedItem as string;
            if (!string.IsNullOrEmpty(scheme))
                {
              //  Growl.Success($"选择了方案: {scheme}");
                LoadXmlAndUpdateByName(scheme);  // 根据选中的名称更新数据
                }
            else
                {
                Growl.ErrorGlobal("未选择有效的方案。");
                }
            }

        /// <summary>
        /// 更新 NumericUpDown 控件的函数
        /// </summary>
        private void UpdateUI(string name, int topDistance, int bottomDistance, int leftDistance, int rightDistance, int rowCount, int columnCount)
            {
            // 更新 NumericUpDown 控件的值
            NumericUpDown1.Value = topDistance;
            NumericUpDown2.Value = bottomDistance;
            NumericUpDown3.Value = leftDistance;
            NumericUpDown4.Value = rightDistance;
            NumericUpDown5.Value = rowCount;
            NumericUpDown6.Value = columnCount;
            }


        //注册事件
        private void AttachValueChangedEventHandlers()
            {
            NumericUpDown1.ValueChanged += NumericUpDown_ValueChanged;
            NumericUpDown2.ValueChanged += NumericUpDown_ValueChanged;
            NumericUpDown3.ValueChanged += NumericUpDown_ValueChanged;
            NumericUpDown4.ValueChanged += NumericUpDown_ValueChanged;
            NumericUpDown5.ValueChanged += NumericUpDown_ValueChanged;
            NumericUpDown6.ValueChanged += NumericUpDown_ValueChanged;

            }
        private void NumericUpDown_ValueChanged(object sender, HandyControl.Data.FunctionEventArgs<double> e)
            {
            getGrid();
            }

        //色彩更换
        private void toggleBlock_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
            {
            colorPicker.Visibility = Visibility.Visible;//显示色卡
            }

        private void colorPicker_Canceled(object sender, EventArgs e)
            {
            colorPicker.Visibility = Visibility.Collapsed;
            }

        private void colorPicker_Confirmed(object sender, HandyControl.Data.FunctionEventArgs<Color> e)
            {
            //colorPicker.Visibility = Visibility.Collapsed;
            //toggleBlock.Background=colorPicker.SelectedBrush;
            }

        private void toggleBlock_MouseDoubleClick(object sender, MouseButtonEventArgs e)
            {
            colorPicker.Visibility = Visibility.Visible;//显示色卡
            }

        private void toggleBlock_MouseEnter(object sender, System.Windows.Forms.MouseEventArgs e)
            {
            colorPicker.Visibility = Visibility.Visible;//显示色卡
            }

        private void toggleBlock_MouseLeave(object sender, System.Windows.Forms.MouseEventArgs e)
            {
            //  colorPicker.Visibility = Visibility.Collapsed;//显示色卡
            }
        /// <summary>
        /// 创建外部参考线
        /// </summary>
        public void getGrid()
            {

            CreatGrid creatGrid = new CreatGrid();
            // 获取 NumericUpdown 控件的值并转换为浮点数
            float hTop = NumericUpDown1 != null ? Convert.ToSingle(NumericUpDown1.Value) : 10;
            float hBottom = NumericUpDown2 != null ? Convert.ToSingle(NumericUpDown2.Value) : 10;
            float hLeft = NumericUpDown3 != null ? Convert.ToSingle(NumericUpDown3.Value) : 10;
            float hRight = NumericUpDown4 != null ? Convert.ToSingle(NumericUpDown4.Value) : 10;
            int hNum = (int)NumericUpDown5.Value;
            int vNum = (int)NumericUpDown6.Value;

            creatGrid.xGrid(hTop, hBottom, hLeft, hRight, hNum, vNum);
            }
        /// <summary>

        private void Button_Click(object sender, RoutedEventArgs e)
            {
            app = Globals.ThisAddIn.Application;
            //删除参考线
            int i = app.ActivePresentation.Guides.Count;
            if (i > 0)
                {
                for (int j = i ; j > 0 ; j--)
                    {
                    Guide guide1 = app.ActivePresentation.Guides[j];
                    guide1.Delete();
                    }
                }
            //删除参考线
            int k = app.ActivePresentation.SlideMaster.Guides.Count;
            if (k > 0)
                {
                for (int j = k ; j > 0 ; j--)
                    {
                    Guide guide2 = app.ActivePresentation.SlideMaster.Guides[j];
                    guide2.Delete();
                    }
                }
            }

        private void Button_Click_1(object sender, RoutedEventArgs e)
            {
            app = Globals.ThisAddIn.Application;
            Presentation pre = app.ActivePresentation;
            Slide slide = app.ActiveWindow.View.Slide;
            //删除参考线
            int i = app.ActivePresentation.SlideMaster.Guides.Count;
            if (i > 0)
                {
                for (int j = i ; j > 0 ; j--)
                    {
                    Guide guide1 = app.ActivePresentation.SlideMaster.Guides[j];
                    guide1.Delete();
                    }
                }
            //添加参考线
            float hTop = (float)NumericUpDown1.Value;
            float hBottom = (float)NumericUpDown2.Value;
            float hLeft = (float)NumericUpDown3.Value;
            float hRight = (float)NumericUpDown4.Value;
            float PageWidth = pre.PageSetup.SlideWidth;
            float PageHeight = pre.PageSetup.SlideHeight;
            float newBottom = PageHeight - hBottom;
            float NewRight = PageWidth - hRight;
            int vNum = (int)NumericUpDown5.Value;
            int hNum = (int)NumericUpDown6.Value;
            //水平参考线的位置
            Guide GuideTop1 = pre.SlideMaster.Guides.Add(PpGuideOrientation.ppHorizontalGuide, hTop);
            Guide GuideTop2 = pre.SlideMaster.Guides.Add(PpGuideOrientation.ppHorizontalGuide, hTop + 30);
            Guide GuideBottom = pre.SlideMaster.Guides.Add(PpGuideOrientation.ppHorizontalGuide, newBottom);

            ////垂直参考线的位置
            Guide GuideLeft = pre.SlideMaster.Guides.Add(PpGuideOrientation.ppVerticalGuide, hLeft);
            Guide GuideRight = pre.SlideMaster.Guides.Add(PpGuideOrientation.ppVerticalGuide, NewRight);
            float vInterval = PageWidth / (vNum + 1); // 垂直间距

            for (int j = 1 ; j <= vNum ; j++)
                {
                var GuideV = pre.SlideMaster.Guides.Add(PpGuideOrientation.ppVerticalGuide, vInterval * j);
                }

            // 添加水平参考线和注释形状
            float hInterval = PageHeight / (hNum + 1); // 水平间距
            for (int j = 1 ; j <= hNum ; j++)
                {
                var GuideH = pre.SlideMaster.Guides.Add(PpGuideOrientation.ppHorizontalGuide, hInterval * j);

                }

            }


        /// <summary>
        /// 设置初始值
        /// </summary>
        public void imposValue()
            {
            app = Globals.ThisAddIn.Application;
            Presentation pre = app.ActivePresentation;
            //添加参考线
            float PageWidth = pre.PageSetup.SlideWidth;
            float PageHeight = pre.PageSetup.SlideHeight;
            NumericUpDown1.Maximum = PageHeight;
            NumericUpDown2.Maximum = PageHeight;
            NumericUpDown3.Maximum = PageWidth;
            NumericUpDown4.Maximum = PageWidth;
            NumericUpDown1.Minimum = 0;
            NumericUpDown2.Minimum = 0;
            NumericUpDown3.Minimum = 0;
            NumericUpDown4.Minimum = 0;
            NumericUpDown5.Minimum = 0;
            NumericUpDown6.Minimum = 0;

            }



        public class CreatGrid
            {
            public PowerPoint.Application app = Globals.ThisAddIn.Application;
            public float xTop { get; set; }
            public float xBottom { get; set; }
            public float xLeft { get; set; }
            public float xRight { get; set; }


            /// <summary>
            /// 创建边缘参考线
            /// </summary>
            /// <param name="hTop"></param>
            /// <param name="hBottom"></param>
            /// <param name="hLeft"></param>
            /// <param name="hRight"></param>
            public void xGrid(float hTop, float hBottom, float hLeft, float hRight, int hNum, int vNum)
                {

                app = Globals.ThisAddIn.Application;
                Presentation pre = app.ActivePresentation;
                Slide slide = app.ActiveWindow.View.Slide;
                //删除参考线
                int i = app.ActivePresentation.Guides.Count;
                if (i > 0)
                    {
                    for (int j = i ; j > 0 ; j--)
                        {
                        Guide guide1 = app.ActivePresentation.Guides[j];
                        guide1.Delete();
                        }
                    }
                //添加参考线
                float PageWidth = pre.PageSetup.SlideWidth;
                float PageHeight = pre.PageSetup.SlideHeight;
                float newBottom = PageHeight - hBottom;
                float NewRight = PageWidth - hRight;

                //水平参考线的位置
                Guide GuideTop1 = pre.Guides.Add(PpGuideOrientation.ppHorizontalGuide, hTop);
                Guide GuideTop2 = pre.Guides.Add(PpGuideOrientation.ppHorizontalGuide, hTop + 30);
                Guide GuideBottom = pre.Guides.Add(PpGuideOrientation.ppHorizontalGuide, newBottom);

                ////垂直参考线的位置
                Guide GuideLeft = pre.Guides.Add(PpGuideOrientation.ppVerticalGuide, hLeft);
                Guide GuideRight = pre.Guides.Add(PpGuideOrientation.ppVerticalGuide, NewRight);
                float vInterval = PageWidth / (vNum + 1); // 垂直间距

                for (int j = 1 ; j <= vNum ; j++)
                    {
                    var GuideV = pre.Guides.Add(PpGuideOrientation.ppVerticalGuide, vInterval * j);
                    }

                // 添加水平参考线和注释形状
                float hInterval = PageHeight / (hNum + 1); // 水平间距
                for (int j = 1 ; j <= hNum ; j++)
                    {
                    var GuideH = pre.Guides.Add(PpGuideOrientation.ppHorizontalGuide, hInterval * j);

                    }
                }
            }

        }
    }


