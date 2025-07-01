using System.Diagnostics;
using System.Reflection;
using System.Windows;
using System.Windows.Navigation;

namespace BatchConvertToCompressedFile;

public partial class AboutWindow
{
    public AboutWindow()
    {
        InitializeComponent();

        AppVersionTextBlock.Text = $"Version: {GetApplicationVersion()}";
        DescriptionTextBlock.Text = "A utility for batch compressing files to .7z or .zip and for verifying the integrity of compressed files.";
    }

    private void CloseButton_Click(object sender, RoutedEventArgs e)
    {
        Close();
    }

    private void Hyperlink_RequestNavigate(object sender, RequestNavigateEventArgs e)
    {
        try
        {
            using var process = Process.Start(new ProcessStartInfo
            {
                FileName = e.Uri.AbsoluteUri,
                UseShellExecute = true
            });
        }
        catch (Exception ex)
        {
            if (App.SharedBugReportService != null)
            {
                _ = App.SharedBugReportService.SendBugReportAsync($"Error opening URL: {e.Uri.AbsoluteUri}. Exception: {ex.Message}");
            }

            MessageBox.Show($"Unable to open link: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
        }
        
        e.Handled = true;
    }

    private string GetApplicationVersion()
    {
        var version = Assembly.GetExecutingAssembly().GetName().Version;
        return version?.ToString() ?? "Unknown";
    }
}