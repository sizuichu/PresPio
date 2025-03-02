using System;

namespace PresPio.Public_Wpf.Models
    {
    public class SettingInfo
        {
        public int Id { get; set; }
        public string Key { get; set; }
        public string Value { get; set; }
        public DateTime LastModified { get; set; } = DateTime.Now;
        }
    }