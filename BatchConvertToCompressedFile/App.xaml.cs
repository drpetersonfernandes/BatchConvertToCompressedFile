using System.Globalization;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows;
using System.Windows.Threading;
using SevenZip;

namespace BatchConvertToCompressedFile;

public partial class App : IDisposable
{
    private readonly BugReportService? _bugReportService;

    public static BugReportService? SharedBugReportService { get; private set; }

    public App()
    {
        BugReportService? bugReportService = null;
        try
        {
            bugReportService = new BugReportService(AppConfig.BugReportApiUrl, AppConfig.BugReportApiKey,
                AppConfig.ApplicationName);
            SharedBugReportService = bugReportService;
            _bugReportService = SharedBugReportService;

            AppDomain.CurrentDomain.UnhandledException += CurrentDomain_UnhandledException;
            DispatcherUnhandledException += App_DispatcherUnhandledException;
            TaskScheduler.UnobservedTaskException += TaskScheduler_UnobservedTaskException;
        }
        catch
        {
            // If initialization fails, ensure proper cleanup
            bugReportService?.Dispose();
            SharedBugReportService = null;
            throw;
        }

        // Initialize SevenZipSharp library path
        InitializeSevenZipSharp();

        // Register the Exit event handler
        Exit += App_Exit;
    }

    private void App_Exit(object sender, ExitEventArgs e)
    {
        // Dispose of the shared BugReportService instance
        _bugReportService?.Dispose();
        SharedBugReportService = null;

        // Unregister event handlers to prevent memory leaks
        AppDomain.CurrentDomain.UnhandledException -= CurrentDomain_UnhandledException;
        DispatcherUnhandledException -= App_DispatcherUnhandledException;
        TaskScheduler.UnobservedTaskException -= TaskScheduler_UnobservedTaskException;
    }

    private void InitializeSevenZipSharp()
    {
        try
        {
            // Determine the path to the 7z dll based on the process architecture.
            string dllName;

            if (RuntimeInformation.ProcessArchitecture == Architecture.Arm64)
            {
                dllName = "7z_arm64.dll";
            }
            else
            {
                dllName = Environment.Is64BitProcess ? "7z_x64.dll" : "7z_x86.dll";
            }

            var dllPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, dllName);

            if (File.Exists(dllPath))
            {
                SevenZipBase.SetLibraryPath(dllPath);
            }
            else
            {
                // Notify developer
                // If the specific DLL is not found, log an error. Extraction will likely fail.
                var errorMessage =
                    $"Could not find the required 7-Zip library: {dllName} in {AppDomain.CurrentDomain.BaseDirectory}";
                if (_bugReportService != null)
                {
                    _ = _bugReportService.SendBugReportAsync(errorMessage);
                }
            }
        }
        catch (Exception ex)
        {
            // Notify developer
            if (_bugReportService != null)
            {
                _ = _bugReportService.SendBugReportAsync(ex.Message);
            }
        }
    }

    private void CurrentDomain_UnhandledException(object sender, UnhandledExceptionEventArgs e)
    {
        if (e.ExceptionObject is Exception exception)
        {
            ReportException(exception, "AppDomain.UnhandledException");
        }
    }

    private void App_DispatcherUnhandledException(object sender, DispatcherUnhandledExceptionEventArgs e)
    {
        ReportException(e.Exception, "Application.DispatcherUnhandledException");
        e.Handled = true;
    }

    private void TaskScheduler_UnobservedTaskException(object? sender, UnobservedTaskExceptionEventArgs e)
    {
        ReportException(e.Exception, "TaskScheduler.UnobservedTaskException");
        e.SetObserved();
    }

    private async void ReportException(Exception exception, string source)
    {
        try
        {
            var message = BuildExceptionReport(exception, source);
            if (_bugReportService != null)
            {
                await _bugReportService.SendBugReportAsync(message);
            }
        }
        catch
        {
            // Silently ignore
        }
    }

    internal static string BuildExceptionReport(Exception exception, string source)
    {
        var sb = new StringBuilder();
        sb.AppendLine(CultureInfo.InvariantCulture, $"Error Source: {source}");
        sb.AppendLine(CultureInfo.InvariantCulture, $"Date and Time: {DateTime.Now}");
        sb.AppendLine(CultureInfo.InvariantCulture, $"OS Version: {Environment.OSVersion}");
        sb.AppendLine(CultureInfo.InvariantCulture, $".NET Version: {Environment.Version}");
        sb.AppendLine();
        sb.AppendLine("Exception Details:");
        AppendExceptionDetails(sb, exception);
        return sb.ToString();
    }

    internal static void AppendExceptionDetails(StringBuilder sb, Exception exception, int level = 0)
    {
        while (true)
        {
            var indent = new string(' ', level * 2);
            sb.AppendLine(CultureInfo.InvariantCulture, $"{indent}Type: {exception.GetType().FullName}");
            sb.AppendLine(CultureInfo.InvariantCulture, $"{indent}Message: {exception.Message}");
            sb.AppendLine(CultureInfo.InvariantCulture, $"{indent}Source: {exception.Source}");
            sb.AppendLine(CultureInfo.InvariantCulture, $"{indent}StackTrace:");
            sb.AppendLine(CultureInfo.InvariantCulture, $"{indent}{exception.StackTrace}");

            if (exception.InnerException == null) break;

            sb.AppendLine(CultureInfo.InvariantCulture, $"{indent}Inner Exception:");
            exception = exception.InnerException;
            level += 1;
        }
    }

    public void Dispose()
    {
        _bugReportService?.Dispose();
        SharedBugReportService = null;

        AppDomain.CurrentDomain.UnhandledException -= CurrentDomain_UnhandledException;
        DispatcherUnhandledException -= App_DispatcherUnhandledException;
        TaskScheduler.UnobservedTaskException -= TaskScheduler_UnobservedTaskException;

        GC.SuppressFinalize(this);
    }
}