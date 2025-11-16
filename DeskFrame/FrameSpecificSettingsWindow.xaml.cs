using System.Windows;
using System.Windows.Media;
using Wpf.Ui.Controls;
using Color = System.Windows.Media.Color;
using ColorConverter = System.Windows.Media.ColorConverter;

namespace DeskFrame
{
    public partial class FrameSpecificSettingsWindow : FluentWindow
    {
        private DeskFrameWindow _frame;
        private Instance _instance;

        public FrameSpecificSettingsWindow(DeskFrameWindow frame)
        {
            InitializeComponent();
            _frame = frame;
            _instance = frame.Instance;

            // Load current settings
            OpacitySlider.Value = (_instance.Opacity / 255.0) * 100;
            BackgroundColorTextBox.Text = _instance.ListViewBackgroundColor;
            
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
            UpdateColorPreview(BackgroundColorTextBox.Text);
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
                // Invalid color
            }
        }

        private void PickColorButton_Click(object sender, RoutedEventArgs e)
        {
            var colorDialog = new System.Windows.Forms.ColorDialog
            {
                AllowFullOpen = true,
                FullOpen = true,
                AnyColor = true
            };

            if (colorDialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                var color = Color.FromArgb(
                    12,
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

        private void ApplyButton_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                // Apply settings to this frame only
                int newOpacityValue = (int)((OpacitySlider.Value / 100.0) * 255);
                string newBackgroundColor = BackgroundColorTextBox.Text;
                
                // Update instance properties FIRST
                _instance.Opacity = newOpacityValue;
                _instance.ListViewBackgroundColor = newBackgroundColor;
                
                // Force UI update immediately
                System.Windows.Application.Current.Dispatcher.Invoke(() =>
                {
                    try
                    {
                        // Parse the new color
                        var color = (Color)ColorConverter.ConvertFromString(newBackgroundColor);
                        
                        // Apply with new opacity
                        _frame.WindowBackground.Background = new SolidColorBrush(
                            Color.FromArgb((byte)newOpacityValue, color.R, color.G, color.B)
                        );
                    }
                    catch (Exception ex)
                    {
                        System.Diagnostics.Debug.WriteLine($"Failed to update UI: {ex.Message}");
                    }
                });
                
                var messageBox = new Wpf.Ui.Controls.MessageBox
                {
                    Title = "Settings Applied",
                    Content = "Settings have been applied to this frame!",
                    CloseButtonText = "OK"
                };
                messageBox.ShowDialogAsync();
            }
            catch (Exception ex)
            {
                var messageBox = new Wpf.Ui.Controls.MessageBox
                {
                    Title = "Error",
                    Content = $"Failed to apply settings: {ex.Message}",
                    CloseButtonText = "OK"
                };
                messageBox.ShowDialogAsync();
            }
        }

        private void ResetButton_Click(object sender, RoutedEventArgs e)
        {
            // Load from global settings
            var appSettings = AppSettings.Load();
            OpacitySlider.Value = appSettings.GetOpacityPercentage();
            BackgroundColorTextBox.Text = appSettings.DefaultBackgroundColor;
        }
    }
}
