using API.Helper.Enums;
using API.Helper.Interfaces;
using API.Helper.Types;
using System.Net.Http.Headers;
using System.Text.Json;

namespace API.Helper.Services
{
    public class SharePointService(IAuthService authService, IHttpService httpService) : ISharePointService
    {
        private readonly IAuthService _authService = authService;
        private readonly IHttpService _httpService = httpService;

        /// <summary>
        /// Get the headers for a request
        /// </summary>
        /// <param name="url">URL of the request</param>
        /// <param name="contentType">Optional content type, defaults to application/json;odata=nometadata</param>
        /// <param name="acceptLanguage">Optional localization you wish to use (e.g., en-US, fi-FI), defaults to the SharePoint site default language.</param>
        /// <returns>Authorization header with a Bearer token, Accept header with the provided content type, and an Accept-Language header with the provided localization.</returns>
        private async Task<HttpRequestHeaders> GetHeaders(string url, string? contentType = null, string? acceptLanguage = null)
        {
            var headers = new HttpRequestMessage().Headers;

            headers.Add("Authorization", $"Bearer {await _authService.GetAccessTokenAsync($"https://{new Uri(url).Host}")}");
            headers.Add("Accept", contentType ?? "application/json;odata=nometadata");
            if (acceptLanguage != null)
                headers.Add("Accept-Language", acceptLanguage);

            return headers;
        }

        /// <summary>
        /// Get a single item
        /// </summary>
        /// <typeparam name="T">The type of object you are expecting to receive in the response</typeparam>
        /// <param name="url">URL of the request</param>
        /// <param name="contentType">Optional content type, defaults to application/json;odata=nometadata</param>
        /// <param name="acceptLanguage">Optional localization you wish to use (e.g., en-US, fi-FI), defaults to the SharePoint site default language.</param>
        /// <returns>A single item deserialized into the provided object type.</returns>

        private async Task<T?> Get<T>(string url, string? contentType = null, string? acceptLanguage = null)
        {
            var response = await _httpService.GetResponseAsync(url, Method.Get, await GetHeaders(url, contentType, acceptLanguage));

            return response.Value.Deserialize<T>();
        }

        /// <summary>
        /// Get a collection of items (with paging)
        /// </summary>
        /// <typeparam name="T">The type of object you are expecting to receive in the response collection</typeparam>
        /// <param name="url">URL of the request</param>
        /// <param name="contentType">Optional content type, defaults to application/json;odata=nometadata</param>
        /// <param name="acceptLanguage">Optional localization you wish to use (e.g., en-US, fi-FI), defaults to the SharePoint site default language.</param>
        /// <param name="items">When recursively paging through the retrieved items, this parameter contains the items collected so far.</param>
        /// <returns>A collection of items deserialized into the provided object type.</returns>
        private async Task<IEnumerable<T>> GetCollection<T>(string url, string? contentType = null, string? acceptLanguage = null, IEnumerable<T>? items = null)
        {
            items ??= [];

            var response = await Get<CollectionResponse<T>>(url, contentType, acceptLanguage);

            if (response != null)
            {
                if (response?.Value != null)
                    items = items.Concat(response.Value);

                if (response?.ODataNextLink != null)
                    return await GetCollection(response.ODataNextLink, contentType, acceptLanguage, items);
            }

            return items;
        }

        /// <summary>
        /// Make a create request
        /// </summary>
        /// <typeparam name="T">The type of object you are expecting to receive in the response</typeparam>
        /// <param name="url">URL of the request</param>
        /// <param name="body">Optional body of the request</param>
        /// <param name="contentType">Optional content type, defaults to application/json;odata=nometadata</param>
        /// <param name="acceptLanguage">Optional localization you wish to use (e.g., en-US, fi-FI), defaults to the SharePoint site default language.</param>
        /// <returns>The created object</returns>
        private async Task<T?> Post<T>(string url, string? body = null, string? contentType = null, string? acceptLanguage = null)
        {
            var response = await _httpService.GetResponseAsync(url, Method.Post, await GetHeaders(url, contentType), body, contentType);

            return response.Value.Deserialize<T>();
        }

        /// <summary>
        /// Make an update request
        /// </summary>
        /// <typeparam name="T">The type of object you are expecting to receive in the response</typeparam>
        /// <param name="url">URL of the request</param>
        /// <param name="body">Optional body of the request</param>
        /// <param name="contentType">Optional content type, defaults to application/json;odata=nometadata</param>
        /// <param name="acceptLanguage">Optional localization you wish to use (e.g., en-US, fi-FI), defaults to the SharePoint site default language.</param>
        /// <returns>The updated object</returns>
        private async Task<T?> Patch<T>(string url, string body, string? contentType = null, string? acceptLanguage = null)
        {
            var headers = await GetHeaders(url, contentType, acceptLanguage);
            headers.Add("X-HTTP-Method", "MERGE");
            headers.Add("If-Match", "*");

            var response = await _httpService.GetResponseAsync(url, Method.Patch, headers, body, contentType);

            return response.Value.Deserialize<T>();
        }

        /// <summary>
        /// Make a put request
        /// </summary>
        /// <typeparam name="T">The type of object you are expecting to receive in the response</typeparam>
        /// <param name="url">URL of the request</param>
        /// <param name="body">Optional body of the request</param>
        /// <param name="contentType">Optional content type, defaults to application/json;odata=nometadata</param>
        /// <param name="acceptLanguage">Optional localization you wish to use (e.g., en-US, fi-FI), defaults to the SharePoint site default language.</param>
        /// <returns>The added object</returns>
        private async Task<T?> Put<T>(string url, string body, string? contentType = null, string? acceptLanguage = null)
        {
            var headers = await GetHeaders(url, contentType, acceptLanguage);
            headers.Add("X-HTTP-Method", "PUT");
            headers.Add("If-Match", "*");

            var response = await _httpService.GetResponseAsync(url, Method.Put, headers, body, contentType);

            return response.Value.Deserialize<T>();
        }

        /// <summary>
        /// Make a delete request
        /// </summary>
        /// <param name="url">URL of the request</param>
        /// <returns></returns>
        private async Task Delete(string url)
        {
            var headers = await GetHeaders(url);
            headers.Add("X-HTTP-Method", "DELETE");
            headers.Add("If-Match", "*");

            await _httpService.GetResponseAsync(url, Method.Delete, headers);
        }
    }
}
