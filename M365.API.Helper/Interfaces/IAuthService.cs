namespace M365.API.Helper.Interfaces
{
    public interface IAuthService
    {
        Task<string> GetAccessTokenAsync(string resourceUrl);
    }
}
