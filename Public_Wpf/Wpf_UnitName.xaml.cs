using System;
using System.Collections.ObjectModel;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Threading;
using Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace PresPio.Public_Wpf
    {
    public partial class Wpf_UnitName : Window
        {
        private readonly PowerPoint.Application ppApp;
        private ObservableCollection<ShapeItem> ShapesList { get; set; }
        private ObservableCollection<ShapeItem> FilteredShapesList { get; set; }
        private readonly DispatcherTimer selectionTimer;

        public Wpf_UnitName()
            {
            InitializeComponent();
            ppApp = Globals.ThisAddIn.Application;
            ShapesList = new ObservableCollection<ShapeItem>();
            FilteredShapesList = new ObservableCollection<ShapeItem>();
            ShapesListView.ItemsSource = FilteredShapesList;

            // 初始化选择延迟计时器
            selectionTimer = new DispatcherTimer
                {
                Interval = TimeSpan.FromMilliseconds(200)
                };
            selectionTimer.Tick += SelectionTimer_Tick;

            LoadShapes();
            }

        private void LoadShapes()
            {
            try
                {
                ShapesList.Clear();
                var currentSlide = ppApp.ActiveWindow.View.Slide;
                var shapes = currentSlide.Shapes;
                int index = 1;

                foreach (PowerPoint.Shape shape in shapes)
                    {
                    ShapesList.Add(new ShapeItem
                        {
                        Index = index++,
                        Name = shape.Name,
                        ShapeType = GetMainShapeType(shape),
                        DetailType = GetDetailShapeType(shape),
                        ZOrder = shape.ZOrderPosition,
                        IsSelected = false
                        });
                    }

                ApplyFilter();
                UpdateSelectionInfo();
                }
            catch (Exception ex)
                {
                MessageBox.Show($"加载形状时出错：{ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }

        private string GetMainShapeType(PowerPoint.Shape shape)
            {
            if (shape.Type == MsoShapeType.msoPicture) return "图片";
            if (shape.Type == MsoShapeType.msoTextBox) return "文本框";
            if (shape.Type == MsoShapeType.msoAutoShape) return "自选图形";
            if (shape.Type == MsoShapeType.msoGroup) return "组合";
            if (shape.Type == MsoShapeType.msoTable) return "表格";
            if (shape.Type == MsoShapeType.msoEmbeddedOLEObject) return "OLE对象";
            if (shape.Type == MsoShapeType.msoChart) return "图表";
            if (shape.Type == MsoShapeType.msoSmartArt) return "SmartArt";
            if (shape.Type == MsoShapeType.msoMedia) return "媒体";
            if (shape.Type == MsoShapeType.msoTextEffect) return "艺术字";
            if (shape.Type == MsoShapeType.msoFreeform) return "任意多边形";
            if (shape.Type == MsoShapeType.msoLine) return "直线";
            if (shape.Type == MsoShapeType.msoPlaceholder) return "占位符";
            return "其他";
            }

        private string GetDetailShapeType(PowerPoint.Shape shape)
            {
            if (shape.Type == MsoShapeType.msoAutoShape)
                {
                if (shape.AutoShapeType == MsoAutoShapeType.msoShapeRectangle) return "矩形";
                if (shape.AutoShapeType == MsoAutoShapeType.msoShapeOval) return "椭圆";
                if (shape.AutoShapeType == MsoAutoShapeType.msoShapeRoundedRectangle) return "圆角矩形";
                if (shape.AutoShapeType == MsoAutoShapeType.msoShapeIsoscelesTriangle) return "等腰三角形";
                if (shape.AutoShapeType == MsoAutoShapeType.msoShapeRightTriangle) return "直角三角形";
                if (shape.AutoShapeType == MsoAutoShapeType.msoShape5pointStar) return "五角星";
                if (shape.AutoShapeType == MsoAutoShapeType.msoShape8pointStar) return "八角星";
                if (shape.AutoShapeType == MsoAutoShapeType.msoShapeHeart) return "心形";
                if (shape.AutoShapeType == MsoAutoShapeType.msoShapeDiamond) return "菱形";
                if (shape.AutoShapeType == MsoAutoShapeType.msoShapeParallelogram) return "平行四边形";
                if (shape.AutoShapeType == MsoAutoShapeType.msoShapeTrapezoid) return "梯形";
                if (shape.AutoShapeType == MsoAutoShapeType.msoShapeHexagon) return "六边形";
                if (shape.AutoShapeType == MsoAutoShapeType.msoShapeOctagon) return "八边形";
                if (shape.AutoShapeType == MsoAutoShapeType.msoShapeCross) return "十字形";
                if (shape.AutoShapeType == MsoAutoShapeType.msoShapeRightArrow) return "右箭头";
                if (shape.AutoShapeType == MsoAutoShapeType.msoShapeLeftArrow) return "左箭头";
                if (shape.AutoShapeType == MsoAutoShapeType.msoShapeUpArrow) return "上箭头";
                if (shape.AutoShapeType == MsoAutoShapeType.msoShapeDownArrow) return "下箭头";
                if (shape.AutoShapeType == MsoAutoShapeType.msoShapeFlowchartProcess) return "流程图-过程";
                if (shape.AutoShapeType == MsoAutoShapeType.msoShapeFlowchartDecision) return "流程图-判断";
                return "其他自选图形";
                }
            return string.Empty;
            }

        private void ApplyFilter()
            {
            var selectedType = (FilterTypeComboBox.SelectedItem as ComboBoxItem)?.Content.ToString() ?? "全部";
            var searchText = SearchTextBox.Text.ToLower();

            var filteredShapes = ShapesList.Where(shape =>
                (selectedType == "全部" || shape.ShapeType == selectedType) &&
                (string.IsNullOrEmpty(searchText) ||
                 shape.Name.ToLower().Contains(searchText) ||
                 shape.ShapeType.ToLower().Contains(searchText) ||
                 shape.DetailType.ToLower().Contains(searchText))
            ).ToList();

            FilteredShapesList.Clear();
            foreach (var shape in filteredShapes)
                {
                FilteredShapesList.Add(shape);
                }

            UpdateSelectionInfo();
            }

        private void FilterTypeComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
            {
            ApplyFilter();
            }

        private void SearchTextBox_TextChanged(object sender, TextChangedEventArgs e)
            {
            ApplyFilter();
            }

        private void SelectAllCheckBox_Click(object sender, RoutedEventArgs e)
            {
            var isChecked = SelectAllCheckBox.IsChecked ?? false;
            foreach (var item in FilteredShapesList)
                {
                item.IsSelected = isChecked;
                }
            UpdateSelectionInfo();
            UpdatePPTSelection();
            }

        private void ItemCheckBox_Click(object sender, RoutedEventArgs e)
            {
            UpdateSelectAllCheckBoxState();
            UpdateSelectionInfo();
            UpdatePPTSelection();
            }

        private void UpdateSelectAllCheckBoxState()
            {
            bool allSelected = true;
            bool anySelected = false;

            foreach (var item in FilteredShapesList)
                {
                if (item.IsSelected)
                    {
                    anySelected = true;
                    }
                else
                    {
                    allSelected = false;
                    }
                }

            if (allSelected)
                {
                SelectAllCheckBox.IsChecked = true;
                }
            else if (anySelected)
                {
                SelectAllCheckBox.IsChecked = null;
                }
            else
                {
                SelectAllCheckBox.IsChecked = false;
                }
            }

        private void UpdateSelectionInfo()
            {
            int selectedCount = FilteredShapesList.Count(item => item.IsSelected);
            SelectionInfoText.Text = $"已选择：{selectedCount} 个形状";
            }

        private void SelectionTimer_Tick(object sender, EventArgs e)
            {
            selectionTimer.Stop();
            UpdatePPTSelection();
            }

        private void ShapesListView_SelectionChanged(object sender, SelectionChangedEventArgs e)
            {
            if (e.AddedItems != null && e.AddedItems.Count > 0)
                {
                foreach (ShapeItem item in e.AddedItems)
                    {
                    item.IsSelected = true;
                    }
                }
            if (e.RemovedItems != null && e.RemovedItems.Count > 0)
                {
                foreach (ShapeItem item in e.RemovedItems)
                    {
                    item.IsSelected = false;
                    }
                }
            UpdateSelectAllCheckBoxState();
            UpdateSelectionInfo();
            selectionTimer.Stop();
            selectionTimer.Start();
            }

        private void UpdatePPTSelection()
            {
            try
                {
                var currentSlide = ppApp.ActiveWindow.View.Slide;
                var selectedShapes = new System.Collections.Generic.List<PowerPoint.Shape>();

                foreach (PowerPoint.Shape shape in currentSlide.Shapes)
                    {
                    var shapeItem = FilteredShapesList.FirstOrDefault(item => item.Name == shape.Name);
                    if (shapeItem != null && shapeItem.IsSelected)
                        {
                        selectedShapes.Add(shape);
                        }
                    }

                if (selectedShapes.Count > 0)
                    {
                    currentSlide.Shapes.Range(selectedShapes.Select(s => s.Name).ToArray()).Select();
                    }
                else
                    {
                    currentSlide.Shapes.Range().Select();
                    ppApp.ActiveWindow.Selection.Unselect();
                    }
                }
            catch (Exception ex)
                {
                MessageBox.Show($"更新选择状态时出错：{ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }

        private void RenameBtn_Click(object sender, RoutedEventArgs e)
            {
            try
                {
                var selectedItems = FilteredShapesList.Where(item => item.IsSelected).ToList();
                if (!selectedItems.Any())
                    {
                    MessageBox.Show("请先选择要重命名的形状", "提示", MessageBoxButton.OK, MessageBoxImage.Information);
                    return;
                    }

                var newName = NewNameTextBox.Text.Trim();
                if (string.IsNullOrEmpty(newName))
                    {
                    MessageBox.Show("请输入新名称", "提示", MessageBoxButton.OK, MessageBoxImage.Information);
                    return;
                    }

                var currentSlide = ppApp.ActiveWindow.View.Slide;
                if (selectedItems.Count == 1)
                    {
                    string targetName = selectedItems[0].Name;
                    PowerPoint.Shape shape = null;
                    foreach (PowerPoint.Shape s in currentSlide.Shapes)
                        {
                        if (s.Name == targetName)
                            {
                            shape = s;
                            break;
                            }
                        }
                    if (shape != null)
                        {
                        shape.Name = newName;
                        selectedItems[0].Name = newName;
                        }
                    }
                else
                    {
                    int index = 1;
                    foreach (var item in selectedItems)
                        {
                        string targetName = item.Name;
                        PowerPoint.Shape shape = null;
                        foreach (PowerPoint.Shape s in currentSlide.Shapes)
                            {
                            if (s.Name == targetName)
                                {
                                shape = s;
                                break;
                                }
                            }
                        if (shape != null)
                            {
                            shape.Name = $"{newName}_{index:D2}";
                            item.Name = $"{newName}_{index:D2}";
                            index++;
                            }
                        }
                    }
                }
            catch (Exception ex)
                {
                MessageBox.Show($"重命名时出错：{ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }

        private void BatchRenameBtn_Click(object sender, RoutedEventArgs e)
            {
            try
                {
                var selectedItems = FilteredShapesList.Where(item => item.IsSelected).ToList();
                if (!selectedItems.Any())
                    {
                    MessageBox.Show("请先选择要批量重命名的形状", "提示", MessageBoxButton.OK, MessageBoxImage.Information);
                    return;
                    }

                var prefix = BatchPrefixTextBox.Text.Trim();
                if (string.IsNullOrEmpty(prefix))
                    {
                    MessageBox.Show("请输入批量命名前缀", "提示", MessageBoxButton.OK, MessageBoxImage.Information);
                    return;
                    }

                if (!int.TryParse(StartNumberTextBox.Text, out int startNumber))
                    {
                    MessageBox.Show("请输入有效的起始序号", "提示", MessageBoxButton.OK, MessageBoxImage.Information);
                    return;
                    }

                var currentSlide = ppApp.ActiveWindow.View.Slide;
                foreach (var item in selectedItems)
                    {
                    string targetName = item.Name;
                    PowerPoint.Shape shape = null;
                    foreach (PowerPoint.Shape s in currentSlide.Shapes)
                        {
                        if (s.Name == targetName)
                            {
                            shape = s;
                            break;
                            }
                        }
                    if (shape != null)
                        {
                        shape.Name = $"{prefix}_{startNumber:D2}";
                        item.Name = $"{prefix}_{startNumber:D2}";
                        startNumber++;
                        }
                    }
                }
            catch (Exception ex)
                {
                MessageBox.Show($"批量重命名时出错：{ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }

        private void BringToFrontBtn_Click(object sender, RoutedEventArgs e)
            {
            try
                {
                var currentSlide = ppApp.ActiveWindow.View.Slide;
                foreach (PowerPoint.Shape shape in currentSlide.Shapes)
                    {
                    if (FilteredShapesList.Any(item => item.IsSelected && item.Name == shape.Name))
                        {
                        shape.ZOrder(MsoZOrderCmd.msoBringToFront);
                        }
                    }
                LoadShapes();
                }
            catch (Exception ex)
                {
                MessageBox.Show($"调整图层时出错：{ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }

        private void SendToBackBtn_Click(object sender, RoutedEventArgs e)
            {
            try
                {
                var currentSlide = ppApp.ActiveWindow.View.Slide;
                foreach (PowerPoint.Shape shape in currentSlide.Shapes)
                    {
                    if (FilteredShapesList.Any(item => item.IsSelected && item.Name == shape.Name))
                        {
                        shape.ZOrder(MsoZOrderCmd.msoSendToBack);
                        }
                    }
                LoadShapes();
                }
            catch (Exception ex)
                {
                MessageBox.Show($"调整图层时出错：{ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }

        private void BringForwardBtn_Click(object sender, RoutedEventArgs e)
            {
            try
                {
                var currentSlide = ppApp.ActiveWindow.View.Slide;
                foreach (PowerPoint.Shape shape in currentSlide.Shapes)
                    {
                    if (FilteredShapesList.Any(item => item.IsSelected && item.Name == shape.Name))
                        {
                        shape.ZOrder(MsoZOrderCmd.msoBringForward);
                        }
                    }
                LoadShapes();
                }
            catch (Exception ex)
                {
                MessageBox.Show($"调整图层时出错：{ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }

        private void SendBackwardBtn_Click(object sender, RoutedEventArgs e)
            {
            try
                {
                var currentSlide = ppApp.ActiveWindow.View.Slide;
                foreach (PowerPoint.Shape shape in currentSlide.Shapes)
                    {
                    if (FilteredShapesList.Any(item => item.IsSelected && item.Name == shape.Name))
                        {
                        shape.ZOrder(MsoZOrderCmd.msoSendBackward);
                        }
                    }
                LoadShapes();
                }
            catch (Exception ex)
                {
                MessageBox.Show($"调整图层时出错：{ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
        }

    public class ShapeItem : System.ComponentModel.INotifyPropertyChanged
        {
        public int Index { get; set; }
        public string Name { get; set; }
        public string ShapeType { get; set; }
        public string DetailType { get; set; }
        public int ZOrder { get; set; }

        private bool isSelected;

        public bool IsSelected
            {
            get => isSelected;
            set
                {
                if (isSelected != value)
                    {
                    isSelected = value;
                    PropertyChanged?.Invoke(this, new System.ComponentModel.PropertyChangedEventArgs(nameof(IsSelected)));
                    }
                }
            }

        public event System.ComponentModel.PropertyChangedEventHandler PropertyChanged;
        }
    }