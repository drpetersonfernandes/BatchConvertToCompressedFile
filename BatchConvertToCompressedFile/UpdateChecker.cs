using System.Net.Http;
using System.Net.Http.Headers;
using System.Net.Http.Json;
using System.Reflection;
using System.Runtime.InteropServices;
using BatchConvertToCompressedFile.models;

namespace BatchConvertToCompressedFile;

/// <summary>
/// Checks for new application versions from a GitHub repository.
/// </summary>
public class UpdateChecker : IDisposable
{
    private readonly HttpClient _httpClient;
    private readonly string _updateUrl;

    /// <summary>
    /// Initializes a new instance of the <see cref="UpdateChecker"/> class.
    /// </summary>
    /// <param name="updateUrl">The GitHub API URL for the latest release.</param>
    public UpdateChecker(string updateUrl)
    {
        _updateUrl = updateUrl;
        _httpClient = new HttpClient();
        // GitHub API requires a User-Agent header.
        _httpClient.DefaultRequestHeaders.UserAgent.Add(
            new ProductInfoHeaderValue(AppConfig.ApplicationName, GetApplicationVersion()));
    }

    /// <summary>
    /// Checks for a new version of the application.
    /// </summary>
    /// <returns>An <see cref="UpdateCheckResult"/> containing update information, or null if no update is available or an error occurs.</returns>
    public async Task<UpdateCheckResult?> CheckForUpdateAsync()
    {
        try
        {
            var latestRelease = await _httpClient.GetFromJsonAsync<GitHubRelease>(_updateUrl);
            if (latestRelease?.TagName == null) return null;

            // Parse version from tag_name (e.g., "release_1.1.0")
            var latestVersionStr = latestRelease.TagName.Split('_').LastOrDefault();
            if (string.IsNullOrEmpty(latestVersionStr) || !Version.TryParse(latestVersionStr, out var latestVersion))
            {
                return null; // Could not parse version from tag.
            }

            var currentVersion = Assembly.GetExecutingAssembly().GetName().Version;
            if (currentVersion == null || latestVersion <= currentVersion)
            {
                return null; // The current version is up-to-date.
            }

            // Find the correct asset for the current architecture.
            var archSuffix = GetArchitectureSuffix();
            var asset = latestRelease.Assets?.FirstOrDefault(a =>
                a.Name != null && a.Name.EndsWith(archSuffix, StringComparison.OrdinalIgnoreCase));

            // If no specific asset is found, don't offer an update.
            if (asset?.BrowserDownloadUrl == null) return null;

            return new UpdateCheckResult(
                true,
                latestVersion.ToString(),
                latestRelease.HtmlUrl,
                asset.BrowserDownloadUrl
            );
        }
        catch (HttpRequestException)
        {
            // Network error, e.g., no internet connection.
            return null;
        }
        catch (Exception)
        {
            // Other errors (e.g., JSON parsing).
            return null;
        }
    }

    /// <summary>
    /// Gets the appropriate asset suffix based on the current process architecture.
    /// </summary>
    private static string GetArchitectureSuffix()
    {
        return RuntimeInformation.ProcessArchitecture switch
        {
            Architecture.Arm64 => "_win-arm64.zip",
            Architecture.X64 => "_win-x64.zip",
            Architecture.X86 => "_win-x86.zip",
            _ => "_win-x64.zip" // Default fallback for unknown architectures.
        };
    }

    /// <summary>
    /// Gets the current application version as a string.
    /// </summary>
    private static string GetApplicationVersion()
    {
        return Assembly.GetExecutingAssembly().GetName().Version?.ToString() ?? "1.0.0.0";
    }

    public void Dispose()
    {
        _httpClient.Dispose();
        GC.SuppressFinalize(this);
    }
}