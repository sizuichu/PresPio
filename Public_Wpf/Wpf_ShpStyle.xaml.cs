using Microsoft.Office.Core;
using System.Windows;
using System.Windows.Controls;

namespace PresPio
    {
    /// <summary>
    /// Wpf_ShpStyle.xaml 的交互逻辑
    /// </summary>
    public partial class Wpf_ShpStyle
        {
        public Wpf_ShpStyle()
            {
            InitializeComponent();
            loadRadion();
            shpRadio1.Click += RadioButton_Click;
            shpRadio2.Click += RadioButton_Click;
            shpRadio3.Click += RadioButton_Click;
            shpRadio4.Click += RadioButton_Click;
            shpRadio5.Click += RadioButton_Click;
            shpRadio6.Click += RadioButton_Click;
            shpRadio7.Click += RadioButton_Click;
            shpRadio8.Click += RadioButton_Click;
            shpRadio9.Click += RadioButton_Click;
            }

        public void loadRadion()
            {
            var item = Properties.Settings.Default.Shape_Style;
            if (item == MsoAutoShapeType.msoShapeRectangle)
                {
                shpRadio1.IsChecked = true;
                }
            else if (item == MsoAutoShapeType.msoShapeSnip2DiagRectangle)
                {
                shpRadio2.IsChecked = true;
                }
            else if (item == MsoAutoShapeType.msoShapeRound1Rectangle)
                {
                shpRadio3.IsChecked = true;
                }
            else if (item == MsoAutoShapeType.msoShapeRoundedRectangle)
                {
                shpRadio4.IsChecked = true;
                }
            else if (item == MsoAutoShapeType.msoShapeSnip2DiagRectangle)
                {
                shpRadio5.IsChecked = true;
                }
            else if (item == MsoAutoShapeType.msoShapeRound2SameRectangle)
                {
                shpRadio6.IsChecked = true;
                }
            else if (item == MsoAutoShapeType.msoShapeSnip1Rectangle)
                {
                shpRadio7.IsChecked = true;
                }
            else if (item == MsoAutoShapeType.msoShapeSnipRoundRectangle)
                {
                shpRadio8.IsChecked = true;
                }
            else if (item == MsoAutoShapeType.msoShapeRound2DiagRectangle)
                {
                shpRadio9.IsChecked = true;
                }
            }

        private void RadioButton_Click(object sender, RoutedEventArgs e)
            {
            RadioButton radioButton = sender as RadioButton;
            if (radioButton != null)
                {
                string buttonName = radioButton.Name;
                switch (buttonName)
                    {
                    case "shpRadio1":
                        Properties.Settings.Default.Shape_Style = MsoAutoShapeType.msoShapeRectangle;
                        break;

                    case "shpRadio2":
                        Properties.Settings.Default.Shape_Style = MsoAutoShapeType.msoShapeSnip2DiagRectangle;
                        break;

                    case "shpRadio3":
                        Properties.Settings.Default.Shape_Style = MsoAutoShapeType.msoShapeRound1Rectangle;
                        break;

                    case "shpRadio4":
                        Properties.Settings.Default.Shape_Style = MsoAutoShapeType.msoShapeRoundedRectangle;
                        break;

                    case "shpRadio5":
                        Properties.Settings.Default.Shape_Style = MsoAutoShapeType.msoShapeSnip2DiagRectangle;
                        break;

                    case "shpRadio6":
                        Properties.Settings.Default.Shape_Style = MsoAutoShapeType.msoShapeRound2SameRectangle;
                        break;

                    case "shpRadio7":
                        Properties.Settings.Default.Shape_Style = MsoAutoShapeType.msoShapeSnip1Rectangle;
                        break;

                    case "shpRadio8":
                        Properties.Settings.Default.Shape_Style = MsoAutoShapeType.msoShapeSnipRoundRectangle;
                        break;

                    case "shpRadio9":
                        Properties.Settings.Default.Shape_Style = MsoAutoShapeType.msoShapeRound2DiagRectangle;

                        break;
                    }
                }
            Properties.Settings.Default.Save();
            }
        }
    }