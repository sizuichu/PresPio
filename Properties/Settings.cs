using System;
using System.Configuration;

namespace PresPio.Properties
{
    public sealed class Settings : ApplicationSettingsBase
    {
        private static Settings defaultInstance = ((Settings)(ApplicationSettingsBase.Synchronized(new Settings())));

        public static Settings Default
        {
            get
            {
                return defaultInstance;
            }
        }

        [UserScopedSetting()]
        [DefaultSettingValue("")]
        public string PicPath
        {
            get
            {
                return ((string)(this["PicPath"]));
            }
            set
            {
                this["PicPath"] = value;
            }
        }
    }
} 