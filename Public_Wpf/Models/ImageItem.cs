using System;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Windows.Media.Imaging;

namespace PresPio.Public_Wpf.Models
{
    public class ImageItem : INotifyPropertyChanged
    {
        private string filePath;
        private string fileName;
        private string fileSize;
        private DateTime creationTime;
        private DateTime modificationTime;
        private int width;
        private int height;
        private BitmapImage thumbnail;
        private ObservableCollection<TagItem> tags;
        private bool isSelected;

        public string FilePath
        {
            get => filePath;
            set
            {
                filePath = value;
                OnPropertyChanged(nameof(FilePath));
            }
        }

        public string FileName
        {
            get => fileName;
            set
            {
                fileName = value;
                OnPropertyChanged(nameof(FileName));
            }
        }

        public string FileSize
        {
            get => fileSize;
            set
            {
                fileSize = value;
                OnPropertyChanged(nameof(FileSize));
            }
        }

        public DateTime CreationTime
        {
            get => creationTime;
            set
            {
                creationTime = value;
                OnPropertyChanged(nameof(CreationTime));
            }
        }

        public DateTime ModificationTime
        {
            get => modificationTime;
            set
            {
                modificationTime = value;
                OnPropertyChanged(nameof(ModificationTime));
            }
        }

        public int Width
        {
            get => width;
            set
            {
                width = value;
                OnPropertyChanged(nameof(Width));
            }
        }

        public int Height
        {
            get => height;
            set
            {
                height = value;
                OnPropertyChanged(nameof(Height));
            }
        }

        public BitmapImage Thumbnail
        {
            get => thumbnail;
            set
            {
                thumbnail = value;
                OnPropertyChanged(nameof(Thumbnail));
            }
        }

        public ObservableCollection<TagItem> Tags
        {
            get => tags;
            set
            {
                tags = value;
                OnPropertyChanged(nameof(Tags));
            }
        }

        public bool IsSelected
        {
            get => isSelected;
            set
            {
                isSelected = value;
                OnPropertyChanged(nameof(IsSelected));
            }
        }

        public event PropertyChangedEventHandler PropertyChanged;

        protected virtual void OnPropertyChanged(string propertyName)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }
} 