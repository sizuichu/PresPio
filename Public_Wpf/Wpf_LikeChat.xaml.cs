using System.Collections.ObjectModel;
using System.Windows;
using System.Windows.Data;

namespace PresPio
{
    public partial class Wpf_LikeChat : Window
    {
        public ObservableCollection<ChatMessage> ChatMessages { get; set; }

        public Wpf_LikeChat()
        {
            InitializeComponent();
            ChatMessages = new ObservableCollection<ChatMessage>();
            ChatListBox.ItemsSource = ChatMessages;
        }

        private void SendButton_Click(object sender, RoutedEventArgs e)
        {
            if (!string.IsNullOrWhiteSpace(MessageTextBox.Text))
            {
                ChatMessages.Add(new ChatMessage
                {
                    Sender = "User",
                    Message = MessageTextBox.Text,
                    IsSenderAI = false,
                    IsSenderUser = true
                });
                MessageTextBox.Clear();
            }
        }
    }

    public class ChatMessage
    {
        public string Sender { get; set; }
        public string Message { get; set; }
        public bool IsSenderAI { get; set; }
        public bool IsSenderUser { get; set; }
    }

    public class BooleanToVisibilityConverter : IValueConverter
    {
        public object Convert(object value, System.Type targetType, object parameter, System.Globalization.CultureInfo culture)
        {
            return (bool)value ? Visibility.Visible : Visibility.Collapsed;
        }

        public object ConvertBack(object value, System.Type targetType, object parameter, System.Globalization.CultureInfo culture)
        {
            throw new System.NotImplementedException();
        }
    }
}