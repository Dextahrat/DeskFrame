using System;
using System.IO;
using System.Text.Json;

namespace DeskFrame
{
    /// <summary>
    /// Uygulama genelinde kullan?lan ayarlar? tutar ve JSON dosyas?na kaydeder.
    /// Bu ayarlar hem mevcut aç?k olan hem de yeni olu?turulacak tüm frame'lere uygulan?r.
    /// </summary>
    public class AppSettings
    {
        private static readonly string SettingsFilePath = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
            "DeskFrame",
            "app_settings.json"
        );

        /// <summary>
        /// Frame'lerin varsay?lan opacity de?eri (0.0 - 1.0)
        /// Bu de?er de?i?ti?inde TÜM AÇIK FRAME'LER güncellenir!
        /// </summary>
        public double DefaultOpacity { get; set; } = 0.85;

        /// <summary>
        /// Frame'lerin varsay?lan arkaplan rengi (HEX format: #AARRGGBB)
        /// Bu de?er de?i?ti?inde TÜM AÇIK FRAME'LER güncellenir!
        /// </summary>
        public string DefaultBackgroundColor { get; set; } = "#0C000000";

        /// <summary>
        /// Ayarlar? dosyadan yükler. Dosya yoksa varsay?lan de?erlerle yeni bir instance döner.
        /// </summary>
        public static AppSettings Load()
        {
            try
            {
                if (File.Exists(SettingsFilePath))
                {
                    string json = File.ReadAllText(SettingsFilePath);
                    var settings = JsonSerializer.Deserialize<AppSettings>(json);
                    return settings ?? new AppSettings();
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error loading settings: {ex.Message}");
            }

            return new AppSettings();
        }

        /// <summary>
        /// Ayarlar? dosyaya kaydeder.
        /// </summary>
        public void Save()
        {
            try
            {
                // Klasörü olu?tur
                string directory = Path.GetDirectoryName(SettingsFilePath);
                if (!Directory.Exists(directory))
                {
                    Directory.CreateDirectory(directory);
                }

                // JSON'a serialize et ve kaydet
                var options = new JsonSerializerOptions 
                { 
                    WriteIndented = true 
                };
                string json = JsonSerializer.Serialize(this, options);
                File.WriteAllText(SettingsFilePath, json);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error saving settings: {ex.Message}");
            }
        }

        /// <summary>
        /// Opacity de?erini 0-100 aras? bir de?ere çevirir (UI için)
        /// </summary>
        public int GetOpacityPercentage()
        {
            return (int)(DefaultOpacity * 100);
        }

        /// <summary>
        /// 0-100 aras? bir de?eri 0.0-1.0 aras? opacity'ye çevirir
        /// </summary>
        public void SetOpacityFromPercentage(int percentage)
        {
            DefaultOpacity = Math.Clamp(percentage / 100.0, 0.0, 1.0);
        }
    }
}
