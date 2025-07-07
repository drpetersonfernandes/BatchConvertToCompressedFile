using System.ComponentModel;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Controls;
using BatchConvertToCompressedFile.models;
using Microsoft.Win32;
using SevenZip;

namespace BatchConvertToCompressedFile;

public partial class MainWindow : IDisposable
{
    private CancellationTokenSource _cts;
    private static readonly string[] AllSupportedVerificationExtensions = { ".zip", ".7z", ".rar" };

    // Statistics
    private int _totalFilesProcessed;
    private int _processedOkCount;
    private int _failedCount;

    private readonly Stopwatch _operationTimer = new();

    // For Write Speed Calculation
    private const int WriteSpeedUpdateIntervalMs = 1000;
    private readonly Stopwatch _writeSpeedTimer = new();
    private long _bytesWrittenSinceLastUpdate;
    private readonly Lock _writeSpeedLock = new();
    private Timer? _writeSpeedUpdateTimer;

    public MainWindow()
    {
        InitializeComponent();
        _cts = new CancellationTokenSource();

        DisplayCompressionInstructionsInLog();
        ResetOperationStats();
        SetVersionInStatusBar();
    }

    private void SetVersionInStatusBar()
    {
        var version = System.Reflection.Assembly.GetExecutingAssembly().GetName().Version;
        VersionText.Text = $"Version: {version?.ToString() ?? "Unknown"}";
    }

    private void UpdateStatus(string message)
    {
        Application.Current.Dispatcher.InvokeAsync(() => { StatusText.Text = message; });
    }

    private void Window_Loaded(object sender, RoutedEventArgs e)
    {
        // Check for updates on startup. This is silent if no update is found.
        _ = CheckForUpdatesAsync(false);
    }

    private void CheckForUpdatesMenuItem_Click(object sender, RoutedEventArgs e)
    {
        // Manually check for updates and show a message regardless of the outcome.
        _ = CheckForUpdatesAsync(true);
    }

    private async Task CheckForUpdatesAsync(bool isManualCheck)
    {
        LogMessage("Checking for updates...");
        UpdateStatus("Checking for updates...");

        try
        {
            using var updateChecker = new UpdateChecker(AppConfig.UpdateCheckUrl);
            var updateResult = await updateChecker.CheckForUpdateAsync();

            if (updateResult is { IsNewVersionAvailable: true, ReleaseUrl: not null })
            {
                var message = $"A new version ({updateResult.LatestVersion}) is available. Would you like to go to the download page?";
                var result = MessageBox.Show(this, message, "Update Available", MessageBoxButton.YesNo, MessageBoxImage.Information);

                if (result == MessageBoxResult.Yes)
                {
                    try
                    {
                        // Open the release URL in the default browser
                        Process.Start(new ProcessStartInfo
                        {
                            FileName = updateResult.ReleaseUrl,
                            UseShellExecute = true
                        });
                    }
                    catch (Exception ex)
                    {
                        var errorMsg = $"Could not open the download page: {ex.Message}";
                        LogMessage(errorMsg);
                        ShowError(errorMsg);
                    }
                }
            }
            else
            {
                LogMessage("Application is up-to-date.");
                if (isManualCheck)
                {
                    MessageBox.Show(this, "You are using the latest version.", "No Updates Found", MessageBoxButton.OK, MessageBoxImage.Information);
                }
            }
        }
        catch (Exception ex)
        {
            var errorMsg = $"Could not check for updates: {ex.Message}";
            LogMessage(errorMsg);
            if (isManualCheck)
            {
                ShowError(errorMsg);
            }

            // Silently fail on automatic check, but log the error.
            _ = ReportBugAsync("Failed to check for updates", ex);
        }
        finally
        {
            UpdateStatus("Ready");
        }
    }

    private void StartWriteSpeedTracking()
    {
        lock (_writeSpeedLock)
        {
            _bytesWrittenSinceLastUpdate = 0;
            _writeSpeedTimer.Restart();

            _writeSpeedUpdateTimer?.Dispose();
            _writeSpeedUpdateTimer = new Timer(UpdateWriteSpeedCallback, null,
                WriteSpeedUpdateIntervalMs, WriteSpeedUpdateIntervalMs);
        }
    }

    private void StopWriteSpeedTracking()
    {
        lock (_writeSpeedLock)
        {
            _writeSpeedUpdateTimer?.Dispose();
            _writeSpeedUpdateTimer = null;
            _writeSpeedTimer.Stop();
            UpdateWriteSpeedDisplay(0);
        }
    }

    private void UpdateWriteSpeedCallback(object? state)
    {
        lock (_writeSpeedLock)
        {
            if (_writeSpeedTimer.ElapsedMilliseconds == 0) return;

            var elapsedSeconds = _writeSpeedTimer.Elapsed.TotalSeconds;
            var speedInBytesPerSecond = elapsedSeconds > 0 ? _bytesWrittenSinceLastUpdate / elapsedSeconds : 0;
            var speedInMBps = speedInBytesPerSecond / (1024.0 * 1024.0);

            UpdateWriteSpeedDisplay(speedInMBps);

            // Reset for next interval
            _bytesWrittenSinceLastUpdate = 0;
            _writeSpeedTimer.Restart();
        }
    }

    private void AddBytesWritten(long bytes)
    {
        lock (_writeSpeedLock)
        {
            _bytesWrittenSinceLastUpdate += bytes;
        }
    }

    private static string SanitizeFileName(string name)
    {
        var invalidChars = Regex.Escape(new string(Path.GetInvalidFileNameChars()));
        var invalidRegStr = string.Format(CultureInfo.InvariantCulture, @"([{0}]*\.+$)|([{0}]+)", invalidChars);
        return Regex.Replace(name, invalidRegStr, "_");
    }

    private void DisplayCompressionInstructionsInLog()
    {
        LogMessage($"Welcome to {AppConfig.ApplicationName}. (Compression Mode)");
        LogMessage("");
        LogMessage("This program will compress files into individual .7z or .zip archives.");
        LogMessage("");
        LogMessage("Please follow these steps for compression:");
        LogMessage("1. Select the input folder containing files to compress.");
        LogMessage("2. Select the output folder where compressed files will be saved.");
        LogMessage("3. Choose the output format (.7z or .zip).");
        LogMessage("4. Choose whether to delete original files after compression.");
        LogMessage("5. Click 'Start Compression' to begin the process.");
        LogMessage("");
        LogMessage("--- Ready for Compression ---");
        LogMessage("");
    }

    private async Task<bool> CompressWithLibraryAsync(string inputFile, string outputFile, string format,
        CancellationToken token)
    {
        try
        {
            var compressor = new SevenZipCompressor();
            // Configure compression settings
            compressor.CompressionMode = CompressionMode.Create;
            compressor.ArchiveFormat = format == ".7z" ? OutArchiveFormat.SevenZip : OutArchiveFormat.Zip;
            compressor.CompressionMethod = CompressionMethod.Lzma2;
            compressor.CompressionLevel = CompressionLevel.Ultra;
            compressor.FastCompression = false;

            await Task.Run(() =>
            {
                token.ThrowIfCancellationRequested();
                compressor.CompressFiles(outputFile, inputFile);
            }, token);

            // Get output file size and add to bytes written tracking
            if (!File.Exists(outputFile)) return true;

            var outputFileInfo = new FileInfo(outputFile);
            AddBytesWritten(outputFileInfo.Length);

            return true;
        }
        catch (OperationCanceledException)
        {
            throw;
        }
        catch (Exception ex)
        {
            LogMessage($"Error creating {format} file: {ex.Message}");
            return false;
        }
    }

    private void DisplayVerificationInstructionsInLog()
    {
        LogMessage($"Welcome to {AppConfig.ApplicationName}. (Verification Mode)");
        LogMessage("");
        LogMessage("This program will verify the integrity of .7z, .zip, and .rar archives.");
        LogMessage("");
        LogMessage("Please follow these steps for verification:");
        LogMessage("1. Select the folder containing archives to verify.");
        LogMessage("2. Choose whether to include subfolders in the search.");
        LogMessage("3. Optionally, move successfully/failed tested files to other folders.");
        LogMessage("4. Click 'Start Verification' to begin the process.");
        LogMessage("");
        LogMessage("--- Ready for Verification ---");
    }

    private void MainTabControl_SelectionChanged(object sender, SelectionChangedEventArgs e)
    {
        if (e.Source is not TabControl tabControl) return;
        if (!StartCompressionButton.IsEnabled || !StartVerificationButton.IsEnabled) return;

        Application.Current.Dispatcher.InvokeAsync(() => LogViewer.Clear());
        if (tabControl.SelectedItem is TabItem selectedTab)
        {
            switch (selectedTab.Name)
            {
                case "CompressTab":
                    DisplayCompressionInstructionsInLog();
                    break;
                case "VerifyTab":
                    DisplayVerificationInstructionsInLog();
                    break;
            }
        }

        UpdateStatus("Ready");
        UpdateWriteSpeedDisplay(0);
    }

    private void MoveSuccessFilesCheckBox_CheckedUnchecked(object sender, RoutedEventArgs e)
    {
        SuccessFolderPanel.Visibility =
            MoveSuccessFilesCheckBox.IsChecked == true ? Visibility.Visible : Visibility.Collapsed;
        SetControlsState(StartCompressionButton.IsEnabled);
    }

    private void MoveFailedFilesCheckBox_CheckedUnchecked(object sender, RoutedEventArgs e)
    {
        FailedFolderPanel.Visibility =
            MoveFailedFilesCheckBox.IsChecked == true ? Visibility.Visible : Visibility.Collapsed;
        SetControlsState(StartCompressionButton.IsEnabled);
    }

    private void BrowseSuccessFolderButton_Click(object sender, RoutedEventArgs e)
    {
        var folder = SelectFolder("Select folder for successfully verified archives");
        if (!string.IsNullOrEmpty(folder))
        {
            SuccessFolderTextBox.Text = folder;
        }
    }

    private void BrowseFailedFolderButton_Click(object sender, RoutedEventArgs e)
    {
        var folder = SelectFolder("Select folder for failed archives");
        if (!string.IsNullOrEmpty(folder))
        {
            FailedFolderTextBox.Text = folder;
        }
    }

    private void Window_Closing(object sender, CancelEventArgs e)
    {
        _cts?.Cancel();
        _cts?.Dispose();
        StopWriteSpeedTracking();
    }

    private void LogMessage(string message)
    {
        var timestampedMessage = $"[{DateTime.Now:HH:mm:ss.fff}] {message}";
        Application.Current.Dispatcher.InvokeAsync(() =>
        {
            LogViewer.AppendText($"{timestampedMessage}{Environment.NewLine}");
            LogViewer.ScrollToEnd();
        });
    }

    private void BrowseCompressionInputButton_Click(object sender, RoutedEventArgs e)
    {
        var folder = SelectFolder("Select the folder containing files to compress");
        if (string.IsNullOrEmpty(folder)) return;

        CompressionInputFolderTextBox.Text = folder;
        LogMessage($"Compression input folder selected: {folder}");
    }

    private void BrowseCompressionOutputButton_Click(object sender, RoutedEventArgs e)
    {
        var folder = SelectFolder("Select the output folder for compressed files");
        if (string.IsNullOrEmpty(folder)) return;

        CompressionOutputFolderTextBox.Text = folder;
        LogMessage($"Compression output folder selected: {folder}");
    }

    private void BrowseVerificationInputButton_Click(object sender, RoutedEventArgs e)
    {
        var folder = SelectFolder("Select the folder containing archives to verify");
        if (string.IsNullOrEmpty(folder)) return;

        VerificationInputFolderTextBox.Text = folder;
        LogMessage($"Verification input folder selected: {folder}");
    }

    private async void StartCompressionButton_Click(object sender, RoutedEventArgs e)
    {
        try
        {
            await Application.Current.Dispatcher.InvokeAsync(() => LogViewer.Clear());
            DisplayCompressionInstructionsInLog();
            var inputFolder = CompressionInputFolderTextBox.Text;
            var outputFolder = CompressionOutputFolderTextBox.Text;
            var deleteFiles = DeleteOriginalsCheckBox.IsChecked ?? false;
            var outputFormat = Format7ZRadioButton.IsChecked == true ? ".7z" : ".zip";
            if (string.IsNullOrEmpty(inputFolder) || string.IsNullOrEmpty(outputFolder))
            {
                ShowError("Please select both input and output folders for compression.");
                return;
            }

            if (inputFolder.Equals(outputFolder, StringComparison.OrdinalIgnoreCase))
            {
                ShowError("Input and output folders must be different for compression.");
                return;
            }

            // 1. Disable controls immediately to prevent re-entry.
            SetControlsState(false);

            // 2. Safely recreate the CancellationTokenSource.
            _cts?.Dispose();
            _cts = new CancellationTokenSource();

            // 3. Reset stats and start the operation.
            ResetOperationStats();
            _operationTimer.Restart();

            StartWriteSpeedTracking();
            UpdateStatus("Starting batch compression...");
            LogMessage("--- Starting batch compression process... ---");
            LogMessage($"Input folder: {inputFolder}");
            LogMessage($"Output folder: {outputFolder}");
            LogMessage($"Output format: {outputFormat}");
            LogMessage($"Delete original files: {deleteFiles}");
            try
            {
                await PerformBatchCompressionAsync(inputFolder, outputFolder, deleteFiles, outputFormat, _cts.Token);
            }
            catch (OperationCanceledException)
            {
                LogMessage("Compression operation was canceled by user.");
                UpdateStatus("Operation canceled.");
            }
            catch (Exception ex)
            {
                LogMessage($"Error during batch compression: {ex.Message}");
                UpdateStatus("Error during compression. See log for details.");
                _ = ReportBugAsync("Error during batch compression process", ex);
            }
            finally
            {
                _operationTimer.Stop();
                StopWriteSpeedTracking();
                UpdateProcessingTimeDisplay();
                SetControlsState(true);
                LogOperationSummary("Compression");
            }
        }
        catch (Exception ex)
        {
            _ = ReportBugAsync("Error during batch compression process", ex);
        }
    }

    private async void StartVerificationButton_Click(object sender, RoutedEventArgs e)
    {
        try
        {
            await Application.Current.Dispatcher.InvokeAsync(() => LogViewer.Clear());
            DisplayVerificationInstructionsInLog();
            var inputFolder = VerificationInputFolderTextBox.Text;
            var includeSubfolders = VerificationIncludeSubfoldersCheckBox.IsChecked ?? false;
            var moveSuccess = MoveSuccessFilesCheckBox.IsChecked == true;
            var successFolder = SuccessFolderTextBox.Text;
            var moveFailed = MoveFailedFilesCheckBox.IsChecked == true;
            var failedFolder = FailedFolderTextBox.Text;
            if (string.IsNullOrEmpty(inputFolder))
            {
                ShowError("Please select the input folder containing archives to verify.");
                return;
            }

            if (moveSuccess && string.IsNullOrEmpty(successFolder))
            {
                ShowError("Please select a Success Folder or uncheck the option to move successful files.");
                return;
            }

            if (moveFailed && string.IsNullOrEmpty(failedFolder))
            {
                ShowError("Please select a Failed Folder or uncheck the option to move failed files.");
                return;
            }

            if (moveSuccess && moveFailed && !string.IsNullOrEmpty(successFolder) &&
                successFolder.Equals(failedFolder, StringComparison.OrdinalIgnoreCase))
            {
                ShowError("Please select different folders for successful and failed files.");
                return;
            }

            if ((moveSuccess && !string.IsNullOrEmpty(successFolder) &&
                 successFolder.Equals(inputFolder, StringComparison.OrdinalIgnoreCase)) ||
                (moveFailed && !string.IsNullOrEmpty(failedFolder) &&
                 failedFolder.Equals(inputFolder, StringComparison.OrdinalIgnoreCase)))
            {
                ShowError("Please select Success/Failed folders that are different from the Input folder.");
                return;
            }

            if (_cts.IsCancellationRequested) _cts.Dispose();
            _cts = new CancellationTokenSource();
            ResetOperationStats();
            SetControlsState(false);
            _operationTimer.Restart();
            StartWriteSpeedTracking();
            UpdateStatus("Starting batch verification...");
            LogMessage("--- Starting batch verification process... ---");
            LogMessage($"Input folder: {inputFolder}");
            LogMessage($"Include subfolders: {includeSubfolders}");
            if (moveSuccess) LogMessage($"Moving successful files to: {successFolder}");
            if (moveFailed) LogMessage($"Moving failed files to: {failedFolder}");
            try
            {
                await PerformBatchVerificationAsync(inputFolder, includeSubfolders, moveSuccess, successFolder,
                    moveFailed, failedFolder, _cts.Token);
            }
            catch (OperationCanceledException)
            {
                LogMessage("Verification operation was canceled by user.");
                UpdateStatus("Operation canceled.");
            }
            catch (Exception ex)
            {
                LogMessage($"Error during batch verification: {ex.Message}");
                UpdateStatus("Error during verification. See log for details.");
                _ = ReportBugAsync("Error during batch verification process", ex);
            }
            finally
            {
                _operationTimer.Stop();
                StopWriteSpeedTracking();
                UpdateProcessingTimeDisplay();
                SetControlsState(true);
                LogOperationSummary("Verification");
            }
        }
        catch (Exception ex)
        {
            _ = ReportBugAsync("Error during batch verification process", ex);
        }
    }

    private void CancelButton_Click(object sender, RoutedEventArgs e)
    {
        _cts.Cancel();
        LogMessage("Cancellation requested. Waiting for current operation(s) to complete...");
        UpdateStatus("Cancellation requested...");
    }

    private void SetControlsState(bool enabled)
    {
        CompressionInputFolderTextBox.IsEnabled = enabled;
        BrowseCompressionInputButton.IsEnabled = enabled;
        CompressionOutputFolderTextBox.IsEnabled = enabled;
        BrowseCompressionOutputButton.IsEnabled = enabled;
        Format7ZRadioButton.IsEnabled = enabled;
        FormatZipRadioButton.IsEnabled = enabled;
        DeleteOriginalsCheckBox.IsEnabled = enabled;
        StartCompressionButton.IsEnabled = enabled;
        VerificationInputFolderTextBox.IsEnabled = enabled;
        BrowseVerificationInputButton.IsEnabled = enabled;
        VerificationIncludeSubfoldersCheckBox.IsEnabled = enabled;
        StartVerificationButton.IsEnabled = enabled;
        MoveSuccessFilesCheckBox.IsEnabled = enabled;
        SuccessFolderTextBox.IsEnabled = enabled && (MoveSuccessFilesCheckBox.IsChecked == true);
        BrowseSuccessFolderButton.IsEnabled = enabled && (MoveSuccessFilesCheckBox.IsChecked == true);
        MoveFailedFilesCheckBox.IsEnabled = enabled;
        FailedFolderTextBox.IsEnabled = enabled && (MoveFailedFilesCheckBox.IsChecked == true);
        BrowseFailedFolderButton.IsEnabled = enabled && (MoveFailedFilesCheckBox.IsChecked == true);
        SuccessFolderPanel.IsEnabled = enabled;
        FailedFolderPanel.IsEnabled = enabled;
        MainTabControl.IsEnabled = enabled;
        ProgressText.Visibility = enabled ? Visibility.Collapsed : Visibility.Visible;
        ProgressBar.Visibility = enabled ? Visibility.Collapsed : Visibility.Visible;
        CancelButton.Visibility = enabled ? Visibility.Collapsed : Visibility.Visible;
        if (!enabled) return;

        ClearProgressDisplay();
        UpdateWriteSpeedDisplay(0);
        UpdateStatus("Ready");
    }

    private static string? SelectFolder(string description)
    {
        var dialog = new OpenFolderDialog { Title = description };
        return dialog.ShowDialog() == true ? dialog.FolderName : null;
    }

    private async Task PerformBatchCompressionAsync(string inputFolder, string outputFolder, bool deleteFiles,
        string outputFormat, CancellationToken token)
    {
        // Validate input folder exists
        if (!await Task.Run(() => Directory.Exists(inputFolder), token))
        {
            LogMessage($"Error: Input folder does not exist: {inputFolder}");
            throw new DirectoryNotFoundException($"Input folder not found: {inputFolder}");
        }

        // Validate or create output folder
        if (!await Task.Run(() => Directory.Exists(outputFolder), token))
        {
            try
            {
                await Task.Run(() => Directory.CreateDirectory(outputFolder), token);
                LogMessage($"Created output folder: {outputFolder}");
            }
            catch (Exception ex)
            {
                LogMessage($"Error: Cannot create output folder: {outputFolder}. {ex.Message}");
                throw new DirectoryNotFoundException($"Cannot create output folder: {outputFolder}", ex);
            }
        }

        // Check folder permissions by attempting to access
        try
        {
            await Task.Run(() => Directory.EnumerateFiles(inputFolder).Take(1).ToList(), token);
        }
        catch (UnauthorizedAccessException ex)
        {
            LogMessage($"Error: Access denied to input folder: {inputFolder}");
            throw new UnauthorizedAccessException($"Access denied to input folder: {inputFolder}", ex);
        }

        var filesToCompress =
            await Task.Run(() => Directory.GetFiles(inputFolder, "*.*", SearchOption.TopDirectoryOnly), token);
        token.ThrowIfCancellationRequested();
        _totalFilesProcessed = filesToCompress.Length;
        UpdateStatsDisplay();
        LogMessage($"Found {_totalFilesProcessed} files to process for compression.");
        if (_totalFilesProcessed == 0)
        {
            LogMessage("No files found in the input folder for compression.");
            return;
        }

        ProgressBar.Maximum = _totalFilesProcessed;
        token.ThrowIfCancellationRequested();

        var filesActuallyProcessedCount = 0;
        foreach (var inputFile in filesToCompress)
        {
            token.ThrowIfCancellationRequested();
            var fileName = Path.GetFileName(inputFile);
            UpdateProgressDisplay(filesActuallyProcessedCount + 1, _totalFilesProcessed, fileName, "Compressing");

            var status =
                await ProcessSingleFileForCompressionAsync(inputFile, outputFolder, deleteFiles, outputFormat, token);

            switch (status)
            {
                case ProcessStatus.Success:
                    _processedOkCount++;
                    break;
                case ProcessStatus.Failed:
                    _failedCount++;
                    break;
                case ProcessStatus.Skipped:
                    // Do nothing, or you could add a "Skipped" counter.
                    break;
            }

            filesActuallyProcessedCount++;
            UpdateStatsDisplay();
            UpdateProcessingTimeDisplay();
        }
    }

    private async Task<ProcessStatus> ProcessSingleFileForCompressionAsync(string inputFile, string outputFolder,
        bool deleteOriginal, string outputFormat, CancellationToken token)
    {
        var originalFileName = Path.GetFileName(inputFile);
        LogMessage($"Starting to compress: {originalFileName}");
        var baseName = Path.GetFileNameWithoutExtension(inputFile);
        var outputFilePath = Path.Combine(outputFolder, SanitizeFileName(baseName) + outputFormat);
        try
        {
            token.ThrowIfCancellationRequested();
            if (await Task.Run(() => File.Exists(outputFilePath), token))
            {
                LogMessage(
                    $"Skipping {originalFileName}: Output file {Path.GetFileName(outputFilePath)} already exists.");
                return ProcessStatus.Skipped;
            }

            var compressionSuccessful = await CompressWithLibraryAsync(inputFile, outputFilePath, outputFormat, token);
            token.ThrowIfCancellationRequested();
            if (compressionSuccessful)
            {
                LogMessage($"Successfully compressed {originalFileName} to {Path.GetFileName(outputFilePath)}");
                if (deleteOriginal)
                {
                    await TryDeleteFileAsync(inputFile, "original file", token);
                }

                return ProcessStatus.Success;
            }
            else
            {
                LogMessage($"Failed to compress {originalFileName}.");
                await TryDeleteFileAsync(outputFilePath, "partially created archive", CancellationToken.None);
                return ProcessStatus.Failed;
            }
        }
        catch (OperationCanceledException)
        {
            LogMessage($"Compression cancelled for {originalFileName}.");
            await TryDeleteFileAsync(outputFilePath, "partially created archive", CancellationToken.None);
            throw;
        }
        catch (Exception ex)
        {
            LogMessage($"Error compressing file {originalFileName}: {ex.Message}");
            _ = ReportBugAsync($"Error compressing file: {originalFileName}", ex);
            await TryDeleteFileAsync(outputFilePath, "partially created archive", CancellationToken.None);
            return ProcessStatus.Failed;
        }
    }

    private async Task PerformBatchVerificationAsync(string inputFolder, bool includeSubfolders, bool moveSuccess,
        string successFolder, bool moveFailed, string failedFolder, CancellationToken token)
    {
        // Validate input folder exists
        if (!await Task.Run(() => Directory.Exists(inputFolder), token))
        {
            LogMessage($"Error: Input folder does not exist: {inputFolder}");
            throw new DirectoryNotFoundException($"Input folder not found: {inputFolder}");
        }

        // Validate or create success folder if needed
        if (moveSuccess && !string.IsNullOrEmpty(successFolder))
        {
            if (!await Task.Run(() => Directory.Exists(successFolder), token))
            {
                try
                {
                    await Task.Run(() => Directory.CreateDirectory(successFolder), token);
                    LogMessage($"Created success folder: {successFolder}");
                }
                catch (Exception ex)
                {
                    LogMessage($"Error: Cannot create success folder: {successFolder}. {ex.Message}");
                    throw new DirectoryNotFoundException($"Cannot create success folder: {successFolder}", ex);
                }
            }
        }

        // Validate or create failed folder if needed
        if (moveFailed && !string.IsNullOrEmpty(failedFolder))
        {
            if (!await Task.Run(() => Directory.Exists(failedFolder), token))
            {
                try
                {
                    await Task.Run(() => Directory.CreateDirectory(failedFolder), token);
                    LogMessage($"Created failed folder: {failedFolder}");
                }
                catch (Exception ex)
                {
                    LogMessage($"Error: Cannot create failed folder: {failedFolder}. {ex.Message}");
                    throw new DirectoryNotFoundException($"Cannot create failed folder: {failedFolder}", ex);
                }
            }
        }

        // Check folder permissions
        try
        {
            await Task.Run(() => Directory.EnumerateFiles(inputFolder).Take(1).ToList(), token);
        }
        catch (UnauthorizedAccessException ex)
        {
            LogMessage($"Error: Access denied to input folder: {inputFolder}");
            throw new UnauthorizedAccessException($"Access denied to input folder: {inputFolder}", ex);
        }

        var searchOption = includeSubfolders ? SearchOption.AllDirectories : SearchOption.TopDirectoryOnly;
        var filesToVerify = await Task.Run(() =>
            Directory.GetFiles(inputFolder, "*.*", searchOption)
                .Where(file => AllSupportedVerificationExtensions.Contains(Path.GetExtension(file).ToLowerInvariant()))
                .ToArray(), token);
        token.ThrowIfCancellationRequested();
        _totalFilesProcessed = filesToVerify.Length;
        UpdateStatsDisplay();
        LogMessage($"Found {_totalFilesProcessed} archives to verify.");
        if (_totalFilesProcessed == 0)
        {
            LogMessage("No supported archives found in the specified folder for verification.");
            return;
        }

        ProgressBar.Maximum = _totalFilesProcessed;
        var filesActuallyProcessedCount = 0;
        foreach (var archiveFile in filesToVerify)
        {
            token.ThrowIfCancellationRequested();
            var fileName = Path.GetFileName(archiveFile);
            UpdateProgressDisplay(filesActuallyProcessedCount + 1, _totalFilesProcessed, fileName, "Verifying");
            var isValid = await VerifyArchiveAsync(archiveFile, token);
            token.ThrowIfCancellationRequested();
            if (isValid)
            {
                LogMessage($"✓ Verification successful: {fileName}");
                _processedOkCount++;
                if (moveSuccess && !string.IsNullOrEmpty(successFolder))
                {
                    await MoveVerifiedFileAsync(archiveFile, successFolder, inputFolder, includeSubfolders,
                        "successfully verified", token);

                    // Track bytes moved for write speed calculation
                    if (File.Exists(archiveFile))
                    {
                        var fileInfo = new FileInfo(archiveFile);
                        AddBytesWritten(fileInfo.Length);
                    }
                }
            }
            else
            {
                LogMessage($"✗ Verification failed: {fileName}");
                _failedCount++;
                if (moveFailed && !string.IsNullOrEmpty(failedFolder))
                {
                    await MoveVerifiedFileAsync(archiveFile, failedFolder, inputFolder, includeSubfolders,
                        "failed verification", token);

                    // Track bytes moved for write speed calculation
                    if (File.Exists(archiveFile))
                    {
                        var fileInfo = new FileInfo(archiveFile);
                        AddBytesWritten(fileInfo.Length);
                    }
                }
            }

            token.ThrowIfCancellationRequested();
            filesActuallyProcessedCount++;
            UpdateStatsDisplay();
            UpdateProcessingTimeDisplay();
        }
    }

    private async Task MoveVerifiedFileAsync(string sourceFile, string destinationParentFolder, string baseInputFolder,
        bool maintainSubfolders, string moveReason, CancellationToken token)
    {
        var fileName = Path.GetFileName(sourceFile);
        string destinationFile;
        try
        {
            token.ThrowIfCancellationRequested();
            string targetDir;
            if (maintainSubfolders && !string.IsNullOrEmpty(Path.GetDirectoryName(sourceFile)) &&
                Path.GetDirectoryName(sourceFile) != baseInputFolder)
            {
                var relativeDir =
                    Path.GetRelativePath(baseInputFolder, Path.GetDirectoryName(sourceFile) ?? string.Empty);
                targetDir = Path.Combine(destinationParentFolder, relativeDir);
            }
            else
            {
                targetDir = destinationParentFolder;
            }

            destinationFile = Path.Combine(targetDir, fileName);
            if (!await Task.Run(() => Directory.Exists(targetDir), token))
            {
                await Task.Run(() => Directory.CreateDirectory(targetDir), token);
            }

            token.ThrowIfCancellationRequested();
            if (await Task.Run(() => File.Exists(destinationFile), token))
            {
                LogMessage(
                    $"Cannot move {fileName}: Destination file already exists at {destinationFile}. Skipping move.");
                return;
            }

            token.ThrowIfCancellationRequested();
            await Task.Run(() => File.Move(sourceFile, destinationFile), token);
            LogMessage($"Moved {fileName} ({moveReason}) to {destinationFile}");
        }
        catch (OperationCanceledException)
        {
            LogMessage($"Move operation for {fileName} cancelled.");
            throw;
        }
        catch (Exception ex)
        {
            LogMessage($"Error moving {fileName} to {destinationParentFolder}: {ex.Message}");
            _ = ReportBugAsync($"Error moving verified file {fileName}", ex);
        }
    }

    private async Task<bool> VerifyArchiveAsync(string archivePath, CancellationToken token)
    {
        try
        {
            // Use a reasonable timeout to prevent indefinite hanging
            using var timeoutCts = new CancellationTokenSource(TimeSpan.FromMinutes(10)); // 10-minute timeout
            using var combinedCts = CancellationTokenSource.CreateLinkedTokenSource(token, timeoutCts.Token);
            return await Task.Run(() =>
            {
                combinedCts.Token.ThrowIfCancellationRequested();
                using var extractor = new SevenZipExtractor(archivePath);
                // Log start of verification for large files
                var fileInfo = new FileInfo(archivePath);
                if (fileInfo.Length > 100 * 1024 * 1024) // Log for files > 100MB
                {
                    LogMessage(
                        $"Verifying large archive: {Path.GetFileName(archivePath)} ({fileInfo.Length / (1024.0 * 1024):F1} MB)");
                }

                return extractor.Check();
            }, combinedCts.Token);
        }
        catch (OperationCanceledException) when (token.IsCancellationRequested)
        {
            LogMessage($"Verification cancelled for {Path.GetFileName(archivePath)}.");
            throw;
        }
        catch (OperationCanceledException)
        {
            LogMessage($"Verification timeout for {Path.GetFileName(archivePath)}.");
            return false;
        }
        catch (SevenZipException ex)
        {
            LogMessage($"Archive '{Path.GetFileName(archivePath)}' appears to be corrupt or invalid: {ex.Message}");
            return false;
        }
        catch (FileNotFoundException)
        {
            LogMessage($"Archive '{Path.GetFileName(archivePath)}' not found. It may have been moved or deleted.");
            return false;
        }
        catch (UnauthorizedAccessException)
        {
            LogMessage(
                $"Permission denied for archive '{Path.GetFileName(archivePath)}'. Please check file permissions.");
            return false;
        }
        catch (IOException ex) when ((uint)ex.HResult == 0x80070020) // ERROR_SHARING_VIOLATION
        {
            LogMessage($"Archive '{Path.GetFileName(archivePath)}' is locked by another process.");
            return false;
        }
        catch (IOException ex)
        {
            LogMessage($"An I/O error occurred while verifying '{Path.GetFileName(archivePath)}': {ex.Message}");
            return false;
        }
        catch (Exception ex)
        {
            LogMessage(
                $"An unexpected error occurred during verification of '{Path.GetFileName(archivePath)}': {ex.Message}");
            return false;
        }
    }

    private void ResetOperationStats()
    {
        _totalFilesProcessed = 0;
        _processedOkCount = 0;
        _failedCount = 0;
        _operationTimer.Reset();
        UpdateStatsDisplay();
        UpdateProcessingTimeDisplay();
        UpdateWriteSpeedDisplay(0);
        ClearProgressDisplay();
        UpdateStatus("Ready");
    }

    private void UpdateStatsDisplay()
    {
        Application.Current.Dispatcher.InvokeAsync(() =>
        {
            TotalFilesValue.Text = _totalFilesProcessed.ToString(CultureInfo.InvariantCulture);
            SuccessValue.Text = _processedOkCount.ToString(CultureInfo.InvariantCulture);
            FailedValue.Text = _failedCount.ToString(CultureInfo.InvariantCulture);
        });
    }

    private void UpdateProcessingTimeDisplay()
    {
        var elapsed = _operationTimer.Elapsed;
        Application.Current.Dispatcher.InvokeAsync(() => { ProcessingTimeValue.Text = $@"{elapsed:hh\:mm\:ss}"; });
    }

    private void UpdateWriteSpeedDisplay(double speedInMBps)
    {
        Application.Current.Dispatcher.InvokeAsync(() => { WriteSpeedValue.Text = $"{speedInMBps:F1} MB/s"; });
    }

    private void UpdateProgressDisplay(int current, int total, string currentFileName, string operationVerb)
    {
        var percentage = total == 0 ? 0 : (double)current / total * 100;
        Application.Current.Dispatcher.InvokeAsync(() =>
        {
            ProgressText.Text = $"{operationVerb} file {current} of {total}: {currentFileName} ({percentage:F1}%)";
            ProgressBar.Value = current;
            ProgressBar.Maximum = total > 0 ? total : 1;
            ProgressText.Visibility = Visibility.Visible;
            ProgressBar.Visibility = Visibility.Visible;
        });
    }

    private void ClearProgressDisplay()
    {
        Application.Current.Dispatcher.InvokeAsync(() =>
        {
            ProgressBar.Value = 0;
            ProgressBar.Visibility = Visibility.Collapsed;
            ProgressText.Text = string.Empty;
            ProgressText.Visibility = Visibility.Collapsed;
        });
    }

    private async Task TryDeleteFileAsync(string filePath, string description, CancellationToken token)
    {
        try
        {
            if (token != CancellationToken.None) token.ThrowIfCancellationRequested();
            if (!await Task.Run(() => File.Exists(filePath),
                    token == CancellationToken.None
                        ? new CancellationTokenSource(TimeSpan.FromSeconds(5)).Token
                        : token)) return;

            if (token != CancellationToken.None) token.ThrowIfCancellationRequested();
            await Task.Run(() => File.Delete(filePath),
                token == CancellationToken.None ? new CancellationTokenSource(TimeSpan.FromSeconds(5)).Token : token);
            LogMessage($"Deleted {description}: {Path.GetFileName(filePath)}");
        }
        catch (OperationCanceledException) when (token != CancellationToken.None)
        {
            LogMessage("Process was canceled by user.");
            throw;
        }
        catch (Exception ex)
        {
            LogMessage($"Failed to delete {description} {Path.GetFileName(filePath)}: {ex.Message}");
        }
    }

    private void LogOperationSummary(string operationType)
    {
        LogMessage("");
        LogMessage($"--- Batch {operationType.ToLowerInvariant()} completed. ---");
        LogMessage($"Total files processed: {_totalFilesProcessed}");
        LogMessage($"Successfully {GetPastTense(operationType)}: {_processedOkCount} files");
        if (_failedCount > 0) LogMessage($"Failed to {operationType.ToLowerInvariant()}: {_failedCount} files");
        Application.Current.Dispatcher.InvokeAsync(() =>
            ShowMessageBox($"Batch {operationType.ToLowerInvariant()} completed.\n\n" +
                           $"Total files processed: {_totalFilesProcessed}\n" +
                           $"Successfully {GetPastTense(operationType)}: {_processedOkCount} files\n" +
                           $"Failed: {_failedCount} files",
                $"{operationType} Complete", MessageBoxButton.OK,
                _failedCount > 0 ? MessageBoxImage.Warning : MessageBoxImage.Information));
    }

    private string GetPastTense(string verb)
    {
        return verb.ToLowerInvariant() switch
        {
            "compression" => "compressed",
            "verification" => "verified",
            _ => verb.ToLowerInvariant() + "ed"
        };
    }

    private void ShowMessageBox(string message, string title, MessageBoxButton buttons, MessageBoxImage icon)
    {
        MessageBox.Show(this, message, title, buttons, icon);
    }

    private void ShowError(string message)
    {
        Application.Current.Dispatcher.InvokeAsync(() =>
            ShowMessageBox(message, "Error", MessageBoxButton.OK, MessageBoxImage.Error));
    }

    private async Task ReportBugAsync(string message, Exception? exception = null)
    {
        try
        {
            var fullReport = new StringBuilder();
            fullReport.AppendLine("=== Bug Report ===");
            fullReport.AppendLine($"Application: {AppConfig.ApplicationName}");
            fullReport.AppendLine(CultureInfo.InvariantCulture, $"Version: {GetType().Assembly.GetName().Version}");
            fullReport.AppendLine(CultureInfo.InvariantCulture, $"OS: {Environment.OSVersion}");
            fullReport.AppendLine(CultureInfo.InvariantCulture, $".NET Version: {Environment.Version}");
            fullReport.AppendLine(CultureInfo.InvariantCulture, $"Date/Time: {DateTime.Now}");
            fullReport.AppendLine().AppendLine("=== Error Message ===").AppendLine(message).AppendLine();
            if (exception != null)
            {
                fullReport.AppendLine("=== Exception Details ===");
                App.AppendExceptionDetails(fullReport, exception);
            }

            if (LogViewer != null)
            {
                var logContent = await Application.Current.Dispatcher.InvokeAsync(() => LogViewer.Text);
                if (!string.IsNullOrEmpty(logContent))
                    fullReport.AppendLine().AppendLine("=== Application Log ===").Append(logContent);
            }

            if (App.SharedBugReportService != null)
            {
                await App.SharedBugReportService.SendBugReportAsync(fullReport.ToString());
            }
        }
        catch
        {
            /* Silently fail reporting */
        }
    }

    private void ExitMenuItem_Click(object sender, RoutedEventArgs e)
    {
        Close();
    }

    private void AboutMenuItem_Click(object sender, RoutedEventArgs e)
    {
        try
        {
            new AboutWindow { Owner = this }.ShowDialog();
        }
        catch (Exception ex)
        {
            LogMessage($"Error opening About window: {ex.Message}");
            _ = ReportBugAsync("Error opening About window", ex);
        }
    }

    public void Dispose()
    {
        _cts?.Cancel();
        _cts?.Dispose();
        _writeSpeedUpdateTimer?.Dispose();
        _operationTimer?.Stop();
        GC.SuppressFinalize(this);
    }
}
