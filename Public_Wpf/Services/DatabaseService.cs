using System;
using System.Collections.Generic;
using LiteDB;
using PresPio.Public_Wpf.Models;

namespace PresPio.Public_Wpf.Services
    {
    public class DatabaseService : IDisposable
        {
        private readonly string _dbPath;
        private readonly LiteDatabase _db;
        private readonly ILiteCollection<ImageInfo> _images;
        private readonly ILiteCollection<TagInfo> _tags;
        private readonly ILiteCollection<CategoryInfo> _categories;
        private readonly ILiteCollection<SettingInfo> _settings;

        public DatabaseService(string dbPath)
            {
            _dbPath = dbPath;
            _db = new LiteDatabase(dbPath);
            _images = _db.GetCollection<ImageInfo>("images");
            _tags = _db.GetCollection<TagInfo>("tags");
            _categories = _db.GetCollection<CategoryInfo>("categories");
            _settings = _db.GetCollection<SettingInfo>("settings");

            // 创建索引
            _images.EnsureIndex(x => x.FilePath, true);
            _tags.EnsureIndex(x => x.Name, true);
            _categories.EnsureIndex(x => x.Name, true);
            _settings.EnsureIndex(x => x.Key, true);
            }

        public IEnumerable<ImageInfo> GetAllImages()
            {
            return _images.FindAll();
            }

        public void DeleteImage(string filePath)
            {
            _images.DeleteMany(x => x.FilePath == filePath);
            }

        public ImageInfo GetImageByPath(string filePath)
            {
            return _images.FindOne(x => x.FilePath == filePath);
            }

        public void UpsertImage(ImageInfo image)
            {
            var existing = GetImageByPath(image.FilePath);
            if (existing != null)
                {
                image.Id = existing.Id;
                }
            _images.Upsert(image);
            }

        public IEnumerable<ImageInfo> GetImagesByTag(string tagName)
            {
            return _images.Find(x => x.Tags.Contains(tagName));
            }

        public IEnumerable<ImageInfo> GetImagesByCategory(string categoryName)
            {
            return _images.Find(x => x.Category == categoryName);
            }

        public void UpsertTag(TagInfo tag)
            {
            var existing = GetTag(tag.Name);
            if (existing != null)
                {
                tag.Id = existing.Id;
                }
            _tags.Upsert(tag);
            }

        public TagInfo GetTag(string name)
            {
            return _tags.FindOne(x => x.Name == name);
            }

        public IEnumerable<TagInfo> GetAllTags()
            {
            return _tags.FindAll();
            }

        public void UpdateTagImageCount(string tagName)
            {
            var tag = GetTag(tagName);
            if (tag != null)
                {
                tag.ImageCount = _images.Count(x => x.Tags.Contains(tagName));
                _tags.Update(tag);
                }
            }

        public void UpsertCategory(CategoryInfo category)
            {
            var existing = GetCategory(category.Name);
            if (existing != null)
                {
                category.Id = existing.Id;
                }
            _categories.Upsert(category);
            }

        public CategoryInfo GetCategory(string name)
            {
            return _categories.FindOne(x => x.Name == name);
            }

        public IEnumerable<CategoryInfo> GetAllCategories()
            {
            return _categories.FindAll();
            }

        public void UpdateCategoryImageCount(string categoryName)
            {
            var category = GetCategory(categoryName);
            if (category != null)
                {
                category.ImageCount = _images.Count(x => x.Category == categoryName);
                _categories.Update(category);
                }
            }

        public void DeleteCategory(string categoryName)
            {
            // 删除分类
            _categories.DeleteMany(x => x.Name == categoryName);

            // 获取该分类下的所有图片
            var images = _images.Find(x => x.Category == categoryName);

            // 清除这些图片的分类信息
            foreach (var image in images)
                {
                image.Category = null;
                _images.Update(image);
                }
            }

        public void SaveSetting(string key, string value)
            {
            var setting = _settings.FindOne(x => x.Key == key);
            if (setting == null)
                {
                setting = new SettingInfo { Key = key, Value = value };
                _settings.Insert(setting);
                }
            else
                {
                setting.Value = value;
                setting.LastModified = DateTime.Now;
                _settings.Update(setting);
                }
            }

        public string GetSetting(string key)
            {
            var setting = _settings.FindOne(x => x.Key == key);
            return setting?.Value;
            }

        public void Dispose()
            {
            _db?.Dispose();
            }
        }
    }