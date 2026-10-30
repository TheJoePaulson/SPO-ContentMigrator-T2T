using System.Windows;
using System.Windows.Threading;

namespace ContentMigrator;

public partial class App : Application
{
    protected override void OnStartup(StartupEventArgs e)
    {
        // Last-resort guard: surface a dialog instead of a silent crash (same intent as the
        // script's "GUI failed to start" message box).
        DispatcherUnhandledException += OnDispatcherUnhandledException;
        base.OnStartup(e);
    }

    private static void OnDispatcherUnhandledException(object sender, DispatcherUnhandledExceptionEventArgs e)
    {
        MessageBox.Show($"Unexpected error:\n\n{e.Exception.Message}", "Content Migrator T2T - Community Edition",
            MessageBoxButton.OK, MessageBoxImage.Error);
        e.Handled = true;
    }
}
