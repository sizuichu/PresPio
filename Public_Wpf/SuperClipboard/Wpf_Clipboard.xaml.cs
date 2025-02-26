using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Windows.Controls;
using System.Windows.Documents;
using System.Windows.Media.Imaging;

namespace PresPio
    {
    public partial class Wpf_Clipboard
        {
        public ObservableCollection<TabData> Items { get; set; }
        public object SelectedItemContent { get; set; } // 用于绑定到显示内容

        public Wpf_Clipboard()
            {
            InitializeComponent();
            LoadData();
            DataContext = this; // 设置数据上下文
            }

        /// <summary>
        /// 获取所选的内容
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void dataListViewAll_SelectionChanged(object sender, SelectionChangedEventArgs e)
            {
            var listView = sender as System.Windows.Controls.ListView;
            if (listView == null) return;

            var selectedItem = listView.SelectedItem as ListData;
            if (selectedItem != null)
                {
                // 更新所选项内容
                SelectedItemContent = selectedItem.Content;

                // 清空 RichTextBox 的内容
                richTextBox.Document.Blocks.Clear();

                // 根据内容类型设置 RichTextBox 的内容
                if (selectedItem.Content is string textContent)
                    {
                    richTextBox.Document.Blocks.Add(new Paragraph(new Run(textContent)));
                    }
                else if (selectedItem.Content is BitmapImage imageContent)
                    {
                    var image = new System.Windows.Controls.Image { Source = imageContent };
                    var block = new BlockUIContainer(image);
                    richTextBox.Document.Blocks.Add(block);
                    }

                drawer.Dock = Dock.Right;
                drawer.IsOpen = true;
                }
            }

        private void LoadData()
            {
            Items = new ObservableCollection<TabData>
    {
        new TabData
        {
            ItemTab = "全部",
            ListItems = new ObservableCollection<ListData>()
        }
    };

            // 生成 1000 条数据
            for (int i = 1 ; i <= 1000 ; i++)
                {
                Items[0].ListItems.Add(new ListData
                    {
                    Index = i,
                    Name = $"Item {i}",
                    Content = $"这是内容 {i}", // 仅使用文本内容
                    Remark = $"Remark {i}"
                    });
                }

            // 添加其他标签的数据
            Items.Add(new TabData
                {
                ItemTab = "金句",
                ListItems = new ObservableCollection<ListData>
        {
            new ListData { Index = 1, Name = "Text Item 1", Content = "文本内容 1", Remark = "Text Remark 1" },
            new ListData { Index = 2, Name = "Text Item 2", Content = "文本内容 2", Remark = "Text Remark 2" }
        }
                });

            Items.Add(new TabData
                {
                ItemTab = "常用",
                ListItems = new ObservableCollection<ListData>
        {
            new ListData { Index = 1, Name = "常用项 1", Content = "常用内容 1", Remark = "常用备注 1" },
            new ListData { Index = 2, Name = "常用项 2", Content = "常用内容 2", Remark = "常用备注 2" }
        }
                });
            }

        public class TabData
            {
            public string ItemTab { get; set; }
            public ObservableCollection<ListData> ListItems { get; set; }
            }

        public class ListData : INotifyPropertyChanged
            {
            private int _index;      // 序号
            private string _name;    // 名称
            private string _remark;   // 标签
            private object _content;  // 内容，可以是任意类型

            public object Content
                {
                get { return _content; }
                set
                    {
                    _content = value;
                    OnPropertyChanged(nameof(Content));
                    }
                }

            public int Index
                {
                get { return _index; }
                set
                    {
                    _index = value;
                    OnPropertyChanged(nameof(Index));
                    }
                }

            public string Name
                {
                get { return _name; }
                set
                    {
                    _name = value;
                    OnPropertyChanged(nameof(Name));
                    }
                }

            public string Remark
                {
                get { return _remark; }
                set
                    {
                    _remark = value;
                    OnPropertyChanged(nameof(Remark));
                    }
                }

            public event PropertyChangedEventHandler PropertyChanged;

            protected virtual void OnPropertyChanged(string propertyName)
                {
                PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
                }
            }
        }
    }