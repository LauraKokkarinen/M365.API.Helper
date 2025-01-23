namespace M365.API.Helper.Interfaces
{
    public interface IGraphService
    {
        Task<T?> Get<T>(string url);
        Task<IEnumerable<T>> GetCollection<T>(string url, IEnumerable<T>? items = null);
        Task<T?> Post<T>(string url, string? body = null);
        Task<T?> Patch<T>(string url, string body);
        Task<T?> Put<T>(string url, string body);
        Task Delete(string url);
    }
}
