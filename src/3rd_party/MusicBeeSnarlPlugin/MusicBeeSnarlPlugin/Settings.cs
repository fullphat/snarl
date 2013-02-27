using System;
using System.IO;
using System.Xml.Serialization;

namespace MusicBeeSnarlPlugin
{
    public sealed class Settings
    {
        private const String SnarlSettingsSubPath = "\\Snarl\\";
        private const String SnarlSettingsFile = "Settings.xml";

        [XmlElement]
        public String Headline { get; set; }
        [XmlElement]
        public String Text { get; set; }
        [XmlElement]
        public String DefaultIconPath { get; set; }


        public Settings()
        {
            Reset();
        }

        internal void Reset()
        {
            Headline = "%artist%";
            Text = "%track_title%%newline%%album%";
            DefaultIconPath = @"%musicbee_path%\Plugins\MusicBee_Logo.png";
        }

        internal static Settings Load(String dataPath)
        {
            Settings settings = null;
            String path = dataPath + SnarlSettingsSubPath + SnarlSettingsFile;

            if (!File.Exists(path))
                return new Settings();

            using (StreamReader stream = new StreamReader(path))
            {
                XmlSerializer ser = new XmlSerializer(typeof(Settings));
                settings = ser.Deserialize(stream) as Settings;
            }
            return settings;
        }

        internal void Save(String dataPath)
        {
            if (!Directory.Exists(dataPath + SnarlSettingsSubPath))
                Directory.CreateDirectory(dataPath + SnarlSettingsSubPath);

            String path = dataPath + SnarlSettingsSubPath + SnarlSettingsFile;

            using (StreamWriter stream = new StreamWriter(path))
            {
                XmlSerializer ser = new XmlSerializer(typeof(Settings));
                ser.Serialize(stream, this);
            }
        }

        public void DeleteSettings(String dataPath)
        {
            String path = dataPath + SnarlSettingsSubPath + SnarlSettingsFile;
            if (File.Exists(path))
            {
                File.Delete(path);
                Directory.Delete(dataPath + SnarlSettingsSubPath);
            }
        }
    }
}
