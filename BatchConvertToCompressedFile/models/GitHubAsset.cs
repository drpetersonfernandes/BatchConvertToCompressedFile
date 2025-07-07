using System.Text.Json.Serialization;

namespace BatchConvertToCompressedFile.models;

/// <summary>
/// Represents a single asset within a GitHub release.
/// </summary>
internal sealed record GitHubAsset(
    [property: JsonPropertyName("name")] string? Name,
    [property: JsonPropertyName("browser_download_url")]
    string? BrowserDownloadUrl
);