using System.Diagnostics;
using Wpf.Ui.Controls;
using System.Windows;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using Microsoft.Win32;
using Color = System.Windows.Media.Color;
using ColorConverter = System.Windows.Media.ColorConverter;
using Application = System.Windows.Application;

namespace DeskFrame
{
    public partial class SettingsWindow : FluentWindow
    {
        private AppSettings _settings;
        private InstanceController _controller;
        private MainWindow _window;

        public SettingsWindow(InstanceController controller, MainWindow window)
        {
            InitializeComponent();
            _window = window;
            _controller = controller;
            
            // Ayarları yükle
            _settings = AppSettings.Load();
            
            // UI'ı ayarlarla doldur
            LoadSettingsToUI();
        }

        private void LoadSettingsToUI()
        {
            // Opacity slider'ı ayarla
            OpacitySlider.Value = _settings.GetOpacityPercentage();
            OpacityLabel.Text = $"{_settings.GetOpacityPercentage()}%";
            
            // Arkaplan rengini ayarla
            BackgroundColorTextBox.Text = _settings.DefaultBackgroundColor;
            UpdateColorPreview(_settings.DefaultBackgroundColor);
            UpdatePreview();
        }

        private void OpacitySlider_ValueChanged(object sender, RoutedPropertyChangedEventArgs<double> e)
        {
            if (OpacityLabel == null) return;
            
            int value = (int)OpacitySlider.Value;
            OpacityLabel.Text = $"{value}%";
            
            UpdatePreview();
        }

        private void BackgroundColorTextBox_TextChanged(object sender, System.Windows.Controls.TextChangedEventArgs e)
        {
            string colorText = BackgroundColorTextBox.Text;
            UpdateColorPreview(colorText);
            UpdatePreview();
        }

        private void UpdateColorPreview(string colorHex)
        {
            try
            {
                if (ColorPreview == null) return;
                
                var color = (Color)ColorConverter.ConvertFromString(colorHex);
                ColorPreview.Background = new SolidColorBrush(color);
            }
            catch
            {
                // Geçersiz renk formatı
                if (ColorPreview != null)
                {
                    ColorPreview.Background = new SolidColorBrush(Colors.Red);
                }
            }
        }

        private void UpdatePreview()
        {
            if (PreviewBorder == null) return;
            
            try
            {
                string colorHex = BackgroundColorTextBox.Text;
                var color = (Color)ColorConverter.ConvertFromString(colorHex);
                PreviewBorder.Background = new SolidColorBrush(color);
                
                double opacity = OpacitySlider.Value / 100.0;
                PreviewBorder.Opacity = opacity;
            }
            catch
            {
                // Geçersiz renk durumunda preview'ı güncelleme
            }
        }

        private void PickColorButton_Click(object sender, RoutedEventArgs e)
        {
            // Windows Forms ColorDialog kullan
            var colorDialog = new System.Windows.Forms.ColorDialog
            {
                AllowFullOpen = true,
                FullOpen = true,
                AnyColor = true
            };

            if (colorDialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                var color = Color.FromArgb(
                    12, // Varsayılan alpha
                    colorDialog.Color.R,
                    colorDialog.Color.G,
                    colorDialog.Color.B
                );
                
                string hexColor = $"#{color.A:X2}{color.R:X2}{color.G:X2}{color.B:X2}";
                BackgroundColorTextBox.Text = hexColor;
            }
        }

        private void PresetColor_Click(object sender, RoutedEventArgs e)
        {
            if (sender is System.Windows.Controls.Button button && button.Tag is string colorHex)
            {
                BackgroundColorTextBox.Text = colorHex;
            }
        }

        private void SaveButton_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                // Yeni ayarları al
                int newOpacityValue = (int)((OpacitySlider.Value / 100.0) * 255);
                string newBackgroundColor = BackgroundColorTextBox.Text;
                
                // Ayarları güncelle
                _settings.SetOpacityFromPercentage((int)OpacitySlider.Value);
                _settings.DefaultBackgroundColor = newBackgroundColor;
                
                // Kaydet
                _settings.Save();
                
                // TÜM AÇIK FRAME'LERİ GÜNCELLE
                UpdateAllFrames(newOpacityValue, newBackgroundColor);
                
                // Kullanıcıya bilgi ver (pencere açık kalacak)
                var messageBox = new Wpf.Ui.Controls.MessageBox
                {
                    Title = "Settings Saved",
                    Content = "Settings have been saved and applied to all frames successfully!",
                    CloseButtonText = "OK"
                };
                messageBox.ShowDialogAsync();
            }
            catch (Exception ex)
            {
                var messageBox = new Wpf.Ui.Controls.MessageBox
                {
                    Title = "Error",
                    Content = $"Failed to save settings: {ex.Message}",
                    CloseButtonText = "OK"
                };
                messageBox.ShowDialogAsync();
            }
        }

        /// <summary>
        /// Tüm açık frame'lerin görünümünü günceller
        /// </summary>
        private void UpdateAllFrames(int opacityValue, string backgroundColor)
        {
            foreach (var subWindow in _controller._subWindows)
            {
                // Instance ayarlarını güncelle
                subWindow.Instance.Opacity = opacityValue;
                subWindow.Instance.ListViewBackgroundColor = backgroundColor;
                
                // UI'ı hemen güncelle
                Application.Current.Dispatcher.Invoke(() =>
                {
                    subWindow.ChangeBackgroundOpacity(opacityValue);
                });
            }
        }

        private void ResetButton_Click(object sender, RoutedEventArgs e)
        {
            // Varsayılan değerlere dön
            var defaultSettings = new AppSettings();
            OpacitySlider.Value = defaultSettings.GetOpacityPercentage();
            BackgroundColorTextBox.Text = defaultSettings.DefaultBackgroundColor;
        }

        private void FluentWindow_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            // Pencere kapatılırken yapılacak işlemler
        }
    }
}
