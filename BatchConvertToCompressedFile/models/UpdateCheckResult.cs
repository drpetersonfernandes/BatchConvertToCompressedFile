namespace BatchConvertToCompressedFile.models;

/// <summary>
/// Represents the result of an update check.
/// </summary>
/// <param name="IsNewVersionAvailable">Indicates if a new version is available.</param>
/// <param name="LatestVersion">The version number of the latest release.</param>
/// <param name="ReleaseUrl">The URL to the GitHub release page.</param>
/// <param name="DownloadUrl">The direct download URL for the appropriate release asset.</param>
public record UpdateCheckResult(
    bool IsNewVersionAvailable,
    string LatestVersion,
    string? ReleaseUrl,
    string? DownloadUrl
);