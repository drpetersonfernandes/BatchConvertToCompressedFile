using System.ComponentModel;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Controls;
using Microsoft.Win32;
using SevenZipExtractor;

namespace BatchConvertToCompressedFile;

public partial class MainWindow : IDisposable
{
    private CancellationTokenSource _cts;

    private readonly string _sevenZipPath;
    private readonly bool _isSevenZipAvailable;

    private static readonly string[] AllSupportedVerificationExtensions = { ".zip", ".7z", ".rar" };

    private int _currentDegreeOfParallelismForFiles = 1;

    // Statistics
    private int _totalFilesProcessed;
    private int _processedOkCount;
    private int _failedCount;
    private readonly Stopwatch _operationTimer = new();

    // For Write Speed Calculation
    private const int WriteSpeedUpdateIntervalMs = 1000;

    public MainWindow()
    {
        InitializeComponent();
        _cts = new CancellationTokenSource();

        var appDirectory = AppDomain.CurrentDomain.BaseDirectory;
        _sevenZipPath = Path.Combine(appDirectory, "7z.exe");
        _isSevenZipAvailable = File.Exists(_sevenZipPath);

        DisplayCompressionInstructionsInLog();
        ResetOperationStats();
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
        LogMessage("5. Optionally, enable parallel processing for faster compression.");
        LogMessage("6. Click 'Start Compression' to begin the process.");
        LogMessage("");

        if (_isSevenZipAvailable)
        {
            LogMessage("7z.exe found. .7z compression is available.");
        }
        else
        {
            LogMessage("WARNING: 7z.exe not found in the application directory!");
            LogMessage(".7z compression will be disabled.");
            Format7zRadioButton.IsEnabled = false;
            FormatZipRadioButton.IsChecked = true;
            Task.Run(async () => await ReportBugAsync("7z.exe not found on startup. This will prevent .7z compression."));
        }
        LogMessage("Using System.IO.Compression for .zip files.");
        LogMessage("--- Ready for Compression ---");
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
        LogMessage("Using SevenZipExtractor library for verification.");
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
        UpdateWriteSpeedDisplay(0);
    }

    private void MoveSuccessFilesCheckBox_CheckedUnchecked(object sender, RoutedEventArgs e)
    {
        SuccessFolderPanel.Visibility = MoveSuccessFilesCheckBox.IsChecked == true ? Visibility.Visible : Visibility.Collapsed;
        SetControlsState(StartCompressionButton.IsEnabled);
    }

    private void MoveFailedFilesCheckBox_CheckedUnchecked(object sender, RoutedEventArgs e)
    {
        FailedFolderPanel.Visibility = MoveFailedFilesCheckBox.IsChecked == true ? Visibility.Visible : Visibility.Collapsed;
        SetControlsState(StartCompressionButton.IsEnabled);
    }

    private void BrowseSuccessFolderButton_Click(object sender, RoutedEventArgs e)
    {
        var folder = SelectFolder("Select folder for successfully verified archives");
        if (!string.IsNullOrEmpty(folder)) SuccessFolderTextBox.Text = folder;
    }

    private void BrowseFailedFolderButton_Click(object sender, RoutedEventArgs e)
    {
        var folder = SelectFolder("Select folder for failed archives");
        if (!string.IsNullOrEmpty(folder)) FailedFolderTextBox.Text = folder;
    }

    private void Window_Closing(object sender, CancelEventArgs e) => _cts.Cancel();

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
        await Application.Current.Dispatcher.InvokeAsync(() => LogViewer.Clear());
        DisplayCompressionInstructionsInLog();

        var inputFolder = CompressionInputFolderTextBox.Text;
        var outputFolder = CompressionOutputFolderTextBox.Text;
        var deleteFiles = DeleteOriginalsCheckBox.IsChecked ?? false;
        var useParallelFileProcessing = ParallelProcessingCheckBox.IsChecked ?? false;
        var outputFormat = Format7zRadioButton.IsChecked == true ? ".7z" : ".zip";

        if (outputFormat == ".7z" && !_isSevenZipAvailable)
        {
            ShowError("7z.exe is missing. Cannot create .7z archives.");
            return;
        }

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

        if (_cts.IsCancellationRequested) _cts.Dispose();
        _cts = new CancellationTokenSource();

        _currentDegreeOfParallelismForFiles = useParallelFileProcessing ? 3 : 1;

        ResetOperationStats();
        SetControlsState(false);
        _operationTimer.Restart();

        LogMessage("--- Starting batch compression process... ---");
        LogMessage($"Input folder: {inputFolder}");
        LogMessage($"Output folder: {outputFolder}");
        LogMessage($"Output format: {outputFormat}");
        LogMessage($"Delete original files: {deleteFiles}");
        LogMessage($"Parallel file processing: {useParallelFileProcessing} (Max concurrency: {_currentDegreeOfParallelismForFiles})");

        try
        {
            await PerformBatchCompressionAsync(inputFolder, outputFolder, deleteFiles, useParallelFileProcessing, outputFormat, _cts.Token);
        }
        catch (OperationCanceledException)
        {
            LogMessage("Compression operation was canceled by user.");
        }
        catch (Exception ex)
        {
            LogMessage($"Error during batch compression: {ex.Message}");
            await ReportBugAsync("Error during batch compression process", ex);
        }
        finally
        {
            _operationTimer.Stop();
            UpdateProcessingTimeDisplay();
            UpdateWriteSpeedDisplay(0);
            SetControlsState(true);
            LogOperationSummary("Compression");
        }
    }

    private async void StartVerificationButton_Click(object sender, RoutedEventArgs e)
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
        if (moveSuccess && moveFailed && !string.IsNullOrEmpty(successFolder) && successFolder.Equals(failedFolder, StringComparison.OrdinalIgnoreCase))
        {
            ShowError("Please select different folders for successful and failed files.");
            return;
        }
        if ((moveSuccess && !string.IsNullOrEmpty(successFolder) && successFolder.Equals(inputFolder, StringComparison.OrdinalIgnoreCase)) ||
            (moveFailed && !string.IsNullOrEmpty(failedFolder) && failedFolder.Equals(inputFolder, StringComparison.OrdinalIgnoreCase)))
        {
            ShowError("Please select Success/Failed folders that are different from the Input folder.");
            return;
        }

        if (_cts.IsCancellationRequested) _cts.Dispose();
        _cts = new CancellationTokenSource();

        ResetOperationStats();
        SetControlsState(false);
        _operationTimer.Restart();

        LogMessage("--- Starting batch verification process... ---");
        LogMessage($"Input folder: {inputFolder}");
        LogMessage($"Include subfolders: {includeSubfolders}");
        if (moveSuccess) LogMessage($"Moving successful files to: {successFolder}");
        if (moveFailed) LogMessage($"Moving failed files to: {failedFolder}");

        try
        {
            await PerformBatchVerificationAsync(inputFolder, includeSubfolders, moveSuccess, successFolder, moveFailed, failedFolder, _cts.Token);
        }
        catch (OperationCanceledException)
        {
            LogMessage("Verification operation was canceled by user.");
        }
        catch (Exception ex)
        {
            LogMessage($"Error during batch verification: {ex.Message}");
            await ReportBugAsync("Error during batch verification process", ex);
        }
        finally
        {
            _operationTimer.Stop();
            UpdateProcessingTimeDisplay();
            UpdateWriteSpeedDisplay(0);
            SetControlsState(true);
            LogOperationSummary("Verification");
        }
    }

    private void CancelButton_Click(object sender, RoutedEventArgs e)
    {
        _cts.Cancel();
        LogMessage("Cancellation requested. Waiting for current operation(s) to complete...");
    }

    private void SetControlsState(bool enabled)
    {
        CompressionInputFolderTextBox.IsEnabled = enabled;
        BrowseCompressionInputButton.IsEnabled = enabled;
        CompressionOutputFolderTextBox.IsEnabled = enabled;
        BrowseCompressionOutputButton.IsEnabled = enabled;
        Format7zRadioButton.IsEnabled = enabled && _isSevenZipAvailable;
        FormatZipRadioButton.IsEnabled = enabled;
        DeleteOriginalsCheckBox.IsEnabled = enabled;
        ParallelProcessingCheckBox.IsEnabled = enabled;
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
    }

    private static string? SelectFolder(string description)
    {
        var dialog = new OpenFolderDialog { Title = description };
        return dialog.ShowDialog() == true ? dialog.FolderName : null;
    }

    private async Task PerformBatchCompressionAsync(string inputFolder, string outputFolder, bool deleteFiles, bool useParallelFileProcessing, string outputFormat, CancellationToken token)
    {
        var filesToCompress = await Task.Run(() => Directory.GetFiles(inputFolder, "*.*", SearchOption.TopDirectoryOnly), token);
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
        var filesActuallyProcessedCount = 0;

        var parallelOptions = new ParallelOptions { MaxDegreeOfParallelism = useParallelFileProcessing ? _currentDegreeOfParallelismForFiles : 1, CancellationToken = token };
        await Parallel.ForEachAsync(filesToCompress, parallelOptions, async (currentFile, ct) =>
        {
            var success = await ProcessSingleFileForCompressionAsync(currentFile, outputFolder, deleteFiles, outputFormat, ct);
            if (success) Interlocked.Increment(ref _processedOkCount);
            else Interlocked.Increment(ref _failedCount);

            var processedSoFar = Interlocked.Increment(ref filesActuallyProcessedCount);
            UpdateProgressDisplay(processedSoFar, _totalFilesProcessed, Path.GetFileName(currentFile), "Compressing");
            UpdateStatsDisplay();
            UpdateProcessingTimeDisplay();
        });

        token.ThrowIfCancellationRequested();
        UpdateWriteSpeedDisplay(0);
    }

    private async Task<bool> ProcessSingleFileForCompressionAsync(string inputFile, string outputFolder, bool deleteOriginal, string outputFormat, CancellationToken token)
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
                LogMessage($"Skipping {originalFileName}: Output file {Path.GetFileName(outputFilePath)} already exists.");
                return false;
            }

            bool compressionSuccessful;
            if (outputFormat == ".zip")
            {
                compressionSuccessful = await CompressToZipAsync(inputFile, outputFilePath, token);
            }
            else
            {
                compressionSuccessful = await CompressTo7zAsync(inputFile, outputFilePath, token);
            }

            token.ThrowIfCancellationRequested();

            if (compressionSuccessful)
            {
                LogMessage($"Successfully compressed {originalFileName} to {Path.GetFileName(outputFilePath)}");
                if (deleteOriginal)
                {
                    await TryDeleteFileAsync(inputFile, "original file", token);
                }
                return true;
            }
            else
            {
                LogMessage($"Failed to compress {originalFileName}.");
                await TryDeleteFileAsync(outputFilePath, "partially created archive", CancellationToken.None);
                return false;
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
            await ReportBugAsync($"Error compressing file: {originalFileName}", ex);
            await TryDeleteFileAsync(outputFilePath, "partially created archive", CancellationToken.None);
            return false;
        }
        finally
        {
            UpdateWriteSpeedDisplay(0);
        }
    }

    private async Task PerformBatchVerificationAsync(string inputFolder, bool includeSubfolders, bool moveSuccess, string successFolder, bool moveFailed, string failedFolder, CancellationToken token)
    {
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
                    await MoveVerifiedFileAsync(archiveFile, successFolder, inputFolder, includeSubfolders, "successfully verified", token);
                }
            }
            else
            {
                LogMessage($"✗ Verification failed: {fileName}");
                _failedCount++;
                if (moveFailed && !string.IsNullOrEmpty(failedFolder))
                {
                    await MoveVerifiedFileAsync(archiveFile, failedFolder, inputFolder, includeSubfolders, "failed verification", token);
                }
            }

            token.ThrowIfCancellationRequested();
            filesActuallyProcessedCount++;
            UpdateStatsDisplay();
            UpdateProcessingTimeDisplay();
        }
    }

    private async Task MoveVerifiedFileAsync(string sourceFile, string destinationParentFolder, string baseInputFolder, bool maintainSubfolders, string moveReason, CancellationToken token)
    {
        var fileName = Path.GetFileName(sourceFile);
        string destinationFile;

        try
        {
            token.ThrowIfCancellationRequested();
            string targetDir;
            if (maintainSubfolders && !string.IsNullOrEmpty(Path.GetDirectoryName(sourceFile)) && Path.GetDirectoryName(sourceFile) != baseInputFolder)
            {
                var relativeDir = Path.GetRelativePath(baseInputFolder, Path.GetDirectoryName(sourceFile) ?? string.Empty);
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
                LogMessage($"Cannot move {fileName}: Destination file already exists at {destinationFile}. Skipping move.");
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
            await ReportBugAsync($"Error moving verified file {fileName}", ex);
        }
    }

    private async Task<bool> CompressToZipAsync(string inputFile, string outputZipFile, CancellationToken token)
    {
        try
        {
            await Task.Run(() =>
            {
                token.ThrowIfCancellationRequested();
                using var archive = ZipFile.Open(outputZipFile, ZipArchiveMode.Create);
                token.ThrowIfCancellationRequested();
                archive.CreateEntryFromFile(inputFile, Path.GetFileName(inputFile), CompressionLevel.Optimal);
            }, token);
            return await Task.Run(() => File.Exists(outputZipFile), token);
        }
        catch (Exception ex)
        {
            LogMessage($"Error creating zip file {Path.GetFileName(outputZipFile)}: {ex.Message}");
            return false;
        }
    }

    private async Task<bool> CompressTo7zAsync(string inputFile, string output7zFile, CancellationToken token)
    {
        using var process = new Process();
        try
        {
            token.ThrowIfCancellationRequested();
            var arguments = $"a -t7z \"{output7zFile}\" \"{inputFile}\" -y";

            process.StartInfo = new ProcessStartInfo
            {
                FileName = _sevenZipPath,
                Arguments = arguments,
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                UseShellExecute = false,
                CreateNoWindow = true
            };
            process.EnableRaisingEvents = true;

            process.OutputDataReceived += (_, args) => { if (!string.IsNullOrEmpty(args.Data)) LogMessage($"[7-ZIP STDOUT] {args.Data}"); };
            process.ErrorDataReceived += (_, args) => { if (!string.IsNullOrEmpty(args.Data)) LogMessage($"[7-ZIP STDERR] {args.Data}"); };

            process.Start();
            process.BeginOutputReadLine();
            process.BeginErrorReadLine();

            var lastSpeedCheckTime = DateTime.UtcNow;
            long lastFileSize = 0;

            while (!process.HasExited)
            {
                if (token.IsCancellationRequested)
                {
                    try { if (!process.HasExited) process.Kill(true); } catch { /* ignore */ }
                    token.ThrowIfCancellationRequested();
                }

                await Task.Delay(WriteSpeedUpdateIntervalMs, token);
                if (process.HasExited || token.IsCancellationRequested) break;

                try
                {
                    if (await Task.Run(() => File.Exists(output7zFile), token))
                    {
                        var currentFileSize = await Task.Run(() => new FileInfo(output7zFile).Length, token);
                        var currentTime = DateTime.UtcNow;
                        var timeDelta = currentTime - lastSpeedCheckTime;

                        if (timeDelta.TotalSeconds > 0)
                        {
                            var bytesDelta = currentFileSize - lastFileSize;
                            var speed = (bytesDelta / timeDelta.TotalSeconds) / (1024.0 * 1024.0);
                            UpdateWriteSpeedDisplay(speed);
                        }

                        lastFileSize = currentFileSize;
                        lastSpeedCheckTime = currentTime;
                    }
                }
                catch (FileNotFoundException) { /* File might not be created yet */ }
                catch (Exception ex) { LogMessage($"Write speed monitoring error: {ex.Message}"); }
            }

            await process.WaitForExitAsync(token);
            return process.ExitCode == 0;
        }
        catch (OperationCanceledException)
        {
            LogMessage("7z compression operation was canceled by user.");
            try { if (!process.HasExited) process.Kill(true); } catch { /* ignore */ }
            throw;
        }
        catch (Exception ex)
        {
            LogMessage($"Error running 7z.exe for {Path.GetFileName(inputFile)}: {ex.Message}");
            await ReportBugAsync($"Error running 7z.exe for {Path.GetFileName(inputFile)}", ex);
            return false;
        }
    }

    private async Task<bool> VerifyArchiveAsync(string archivePath, CancellationToken token)
    {
        try
        {
            return await Task.Run(() =>
            {
                token.ThrowIfCancellationRequested();
                using var archiveFile = new ArchiveFile(archivePath);
                return archiveFile.Check();
            }, token);
        }
        catch (OperationCanceledException)
        {
            LogMessage($"Verification cancelled for {Path.GetFileName(archivePath)}.");
            throw;
        }
        catch (Exception ex)
        {
            LogMessage($"Error verifying archive {Path.GetFileName(archivePath)}: {ex.Message}");
            await ReportBugAsync($"Error verifying archive: {Path.GetFileName(archivePath)}", ex);
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
            if (!await Task.Run(() => File.Exists(filePath), token == CancellationToken.None ? new CancellationTokenSource(TimeSpan.FromSeconds(5)).Token : token)) return;
            if (token != CancellationToken.None) token.ThrowIfCancellationRequested();
            await Task.Run(() => File.Delete(filePath), token == CancellationToken.None ? new CancellationTokenSource(TimeSpan.FromSeconds(5)).Token : token);
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
        Application.Current.Dispatcher.InvokeAsync(() => ShowMessageBox(message, "Error", MessageBoxButton.OK, MessageBoxImage.Error));
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
                if (!string.IsNullOrEmpty(logContent)) fullReport.AppendLine().AppendLine("=== Application Log ===").Append(logContent);
            }

            if (App.SharedBugReportService != null)
            {
                await App.SharedBugReportService.SendBugReportAsync(fullReport.ToString());
            }
        }
        catch { /* Silently fail reporting */ }
    }

    private void ExitMenuItem_Click(object sender, RoutedEventArgs e) => Close();

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
        _operationTimer?.Stop();
        GC.SuppressFinalize(this);
    }
}
