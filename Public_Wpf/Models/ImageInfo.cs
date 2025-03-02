using System;
using System.Collections.Generic;
using LiteDB;

namespace PresPio.Public_Wpf.Models
    {
    public class ImageInfo
        {
        [BsonId]
        public ObjectId Id { get; set; }

        public string FilePath { get; set; }
        public string FileName { get; set; }
        public string FileSize { get; set; }
        public int Width { get; set; }
        public int Height { get; set; }
        public DateTime CreationTime { get; set; }
        public DateTime ModificationTime { get; set; }
        public List<string> Tags { get; set; } = new List<string>();
        public string Category { get; set; }
        public List<ColorInfo> DominantColors { get; set; } = new List<ColorInfo>();
        public DateTime LastAccessTime { get; set; }
        public DateTime ImportTime { get; set; }
        }

    public class ColorInfo
        {
        public string ColorHex { get; set; }
        public double Percentage { get; set; }
        public (double H, double S, double L) Hsl { get; set; }
        }

    public class CategoryInfo
        {
        [BsonId]
        public ObjectId Id { get; set; }

        public string Name { get; set; }
        public string Path { get; set; }
        public int ImageCount { get; set; }
        public DateTime CreationTime { get; set; }
        }

    public class TagInfo
        {
        [BsonId]
        public ObjectId Id { get; set; }

        public string Name { get; set; }
        public string ColorHex { get; set; }
        public int ImageCount { get; set; }
        public DateTime CreationTime { get; set; }
        }
    }