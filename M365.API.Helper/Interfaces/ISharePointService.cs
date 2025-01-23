namespace M365.API.Helper.Interfaces
{
    public interface ISharePointService
    {
        Task<T?> Get<T>(string url, string? contentType = null, string? acceptLanguage = null);
        Task<IEnumerable<T>> GetCollection<T>(string url, string? contentType = null, string? acceptLanguage = null, IEnumerable<T>? items = null);
        Task<T?> Post<T>(string url, string? body = null, string? contentType = null, string? acceptLanguage = null);
        Task<T?> Patch<T>(string url, string body, string? contentType = null, string? acceptLanguage = null);
        Task<T?> Put<T>(string url, string body, string? contentType = null, string? acceptLanguage = null);
        Task Delete(string url);
    }
}
