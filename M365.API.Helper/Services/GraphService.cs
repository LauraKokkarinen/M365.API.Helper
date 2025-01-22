using M365.API.Helper.Enums;
using M365.API.Helper.Interfaces;
using M365.API.Helper.Types;
using System.Net.Http.Headers;
using System.Text.Json;

namespace M365.API.Helper.Services
{
    public class GraphService(IAuthService authService, IHttpService httpService) : IGraphService
    {
        private readonly IAuthService _authService = authService;
        private readonly IHttpService _httpService = httpService;

        /// <summary>
        /// Get the headers for a request.
        /// </summary>
        /// <returns>Authorization header with a Bearer token.</returns>
        private async Task<HttpRequestHeaders> GetHeaders()
        {
            var headers = new HttpRequestMessage().Headers;

            headers.TryAddWithoutValidation("Authorization", $"Bearer {await _authService.GetAccessTokenAsync("https://graph.microsoft.com/")}");

            return headers;
        }

        /// <summary>
        /// Get a single item.
        /// </summary>
        /// <typeparam name="T">The type of object you are expecting to receive in the response</typeparam>
        /// <param name="url">URL of the request</param>
        /// <returns>A single item deserialized into the provided object type.</returns>
        private async Task<T?> Get<T>(string url)
        {
            var response = await _httpService.GetResponseAsync(url, Method.Get, await GetHeaders());

            return response.Value.Deserialize<T>();
        }

        /// <summary>
        /// Get a collection of items (with paging).
        /// </summary>
        /// <typeparam name="T">The type of object you are expecting to receive in the response collection</typeparam>
        /// <param name="url">URL of the request</param>
        /// <param name="items">When recursively paging through the retrieved items, this parameter contains the items collected so far.</param>
        /// <returns>A collection of items deserialized into the provided object type.</returns>
        private async Task<IEnumerable<T>> GetCollection<T>(string url, IEnumerable<T>? items = null)
        {
            items ??= [];

            var response = await Get<CollectionResponse<T>>(url);

            if (response != null)
            {
                if (response?.Value != null)
                    items = items.Concat(response.Value);

                if (response?.ODataNextLink != null)
                    return await GetCollection(response.ODataNextLink, items);
            }

            return items;
        }

        /// <summary>
        /// Make a create request.
        /// </summary>
        /// <typeparam name="T">The type of object you are expecting to receive in the response</typeparam>
        /// <param name="url">URL of the request</param>
        /// <param name="body">Optional body of the request</param>
        /// <returns>The created object.</returns>
        private async Task<T?> Post<T>(string url, string? body = null)
        {
            var response = await _httpService.GetResponseAsync(url, Method.Post, await GetHeaders(), body);

            return response.Value.Deserialize<T>();
        }

        /// <summary>
        /// Make an update request.
        /// </summary>
        /// <typeparam name="T">The type of object you are expecting to receive in the response</typeparam>
        /// <param name="url">URL of the request</param>
        /// <param name="body">Optional body of the request</param>
        /// <returns>The updated object.</returns>
        private async Task<T?> Patch<T>(string url, string body)
        {
            var response = await _httpService.GetResponseAsync(url, Method.Patch, await GetHeaders(), body);

            return response.Value.Deserialize<T>();
        }

        /// <summary>
        /// Make a put request.
        /// </summary>
        /// <typeparam name="T">The type of object you are expecting to receive in the response</typeparam>
        /// <param name="url">URL of the request</param>
        /// <param name="body">Optional body of the request</param>
        /// <returns>The added object.</returns>
        private async Task<T?> Put<T>(string url, string body)
        {
            var response = await _httpService.GetResponseAsync(url, Method.Put, await GetHeaders(), body);

            return response.Value.Deserialize<T>();
        }

        /// <summary>
        /// Make a delete request.
        /// </summary>
        /// <param name="url">URL of the request</param>
        /// <returns></returns>
        private async Task Delete(string url)
        {
            await _httpService.GetResponseAsync(url, Method.Delete, await GetHeaders());
        }
    }
}