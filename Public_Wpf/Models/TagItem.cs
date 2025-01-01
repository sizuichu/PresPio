using System.ComponentModel;
using System.Windows.Media;

namespace PresPio.Public_Wpf.Models
{
    public class TagItem : INotifyPropertyChanged
    {
        private string name;
        private SolidColorBrush color;

        public string Name
        {
            get => name;
            set
            {
                name = value;
                OnPropertyChanged(nameof(Name));
            }
        }

        public SolidColorBrush Color
        {
            get => color;
            set
            {
                color = value;
                OnPropertyChanged(nameof(Color));
            }
        }

        public event PropertyChangedEventHandler PropertyChanged;

        protected virtual void OnPropertyChanged(string propertyName)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }
} 