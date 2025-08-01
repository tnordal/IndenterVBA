using System;
using System.IO;
using System.Xml.Serialization;

namespace IndenterVBA
{
    [Serializable]
    public class IndenterSettings
    {
        // Default values
        private const int DefaultIndentSpaces = 4;
        private const bool DefaultIndentDeclarations = false;
        private const bool DefaultUseLogging = true;

        // Properties with default values
        public int IndentSpaces { get; set; } = DefaultIndentSpaces;
        public bool IndentDeclarations { get; set; } = DefaultIndentDeclarations;
        public bool UseLogging { get; set; } = DefaultUseLogging;

        // Singleton instance
        private static IndenterSettings _instance;
        public static IndenterSettings Instance
        {
            get
            {
                if (_instance == null)
                {
                    _instance = Load();
                }
                return _instance;
            }
        }

        // Get paths for settings and logs
        public static string AppDataFolder
        {
            get
            {
                string path = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments),
                    "IndentVBA");
                
                if (!Directory.Exists(path))
                {
                    Directory.CreateDirectory(path);
                }
                return path;
            }
        }

        public static string LogsFolder
        {
            get
            {
                string path = Path.Combine(AppDataFolder, "logs");
                if (!Directory.Exists(path))
                {
                    Directory.CreateDirectory(path);
                }
                return path;
            }
        }

        private static string SettingsFilePath => Path.Combine(AppDataFolder, "settings.xml");

        // Save settings to file
        public void Save()
        {
            try
            {
                var serializer = new XmlSerializer(typeof(IndenterSettings));
                using (var stream = new FileStream(SettingsFilePath, FileMode.Create))
                {
                    serializer.Serialize(stream, this);
                }
                _instance = this;
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show($"Error saving settings: {ex.Message}");
            }
        }

        // Load settings from file or create default
        private static IndenterSettings Load()
        {
            try
            {
                if (File.Exists(SettingsFilePath))
                {
                    var serializer = new XmlSerializer(typeof(IndenterSettings));
                    using (var stream = new FileStream(SettingsFilePath, FileMode.Open))
                    {
                        return (IndenterSettings)serializer.Deserialize(stream);
                    }
                }
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show($"Error loading settings: {ex.Message}");
            }

            // Return default settings if loading fails
            return new IndenterSettings();
        }

        // Reset settings to default
        public void ResetToDefault()
        {
            IndentSpaces = DefaultIndentSpaces;
            IndentDeclarations = DefaultIndentDeclarations;
            UseLogging = DefaultUseLogging;
        }
    }
}