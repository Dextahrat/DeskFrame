using DeskFrame.Properties;
using Microsoft.Toolkit.Uwp.Notifications;
using Microsoft.Win32;
using System.Diagnostics;
using System.Windows;
using System.Windows.Threading;
using Application = System.Windows.Application;

namespace DeskFrame
{
    /// <summary>
    /// Interaction logic for App.xaml
    /// </summary>
    public partial class App : Application
    {
        public RegistryHelper reg = new RegistryHelper("DeskFrame");
        protected override void OnStartup(StartupEventArgs e)
        {
#if !DEBUG
            if (Process.GetProcessesByName(Process.GetCurrentProcess().ProcessName).Length > 1)
            {

                var dialog = new Wpf.Ui.Controls.MessageBox
                {
                    Title = "DeskFrame",
                    Content = Lang.DeskFrame_AlreadyRunning,
                };

                var result = dialog.ShowDialogAsync();

                if (result.Result == Wpf.Ui.Controls.MessageBoxResult.None)
                {
                    Application.Current.Shutdown();
                }
            }
#endif
            PresentationTraceSources.DataBindingSource.Switch.Level = SourceLevels.Critical;
            base.OnStartup(e);
        }
    }

}
