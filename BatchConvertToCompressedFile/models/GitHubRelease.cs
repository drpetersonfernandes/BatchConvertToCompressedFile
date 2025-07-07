using System.Text.Json.Serialization;

namespace BatchConvertToCompressedFile.models;

/// <summary>
/// Represents the structure of a GitHub release API response.
/// </summary>
internal sealed record GitHubRelease(
    [property: JsonPropertyName("tag_name")]
    string? TagName,
    [property: JsonPropertyName("html_url")]
    string? HtmlUrl,
    [property: JsonPropertyName("assets")] List<GitHubAsset>? Assets
);