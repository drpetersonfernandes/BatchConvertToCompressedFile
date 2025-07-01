using System.Net.Http;
using System.Net.Http.Json;

namespace BatchConvertToCompressedFile;

/// <inheritdoc />
/// <summary>
/// Service responsible for sending bug reports to the BugReport API
/// </summary>
public class BugReportService(string apiUrl, string apiKey, string applicationName) : IDisposable
{
    private readonly HttpClient _httpClient = new();
    private readonly string _apiUrl = apiUrl;
    private readonly string _apiKey = apiKey;
    private readonly string _applicationName = applicationName;

    public async Task<bool> SendBugReportAsync(string message)
    {
        try
        {
            _httpClient.DefaultRequestHeaders.Clear();
            _httpClient.DefaultRequestHeaders.Add("X-API-KEY", _apiKey);

            var content = JsonContent.Create(new
            {
                message,
                applicationName = _applicationName
            });

            using var response = await _httpClient.PostAsync(_apiUrl, content);
            return response.IsSuccessStatusCode;
        }
        catch
        {
            return false;
        }
    }

    public void Dispose()
    {
        _httpClient.Dispose();
        GC.SuppressFinalize(this);
    }
}