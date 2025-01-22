using System.Text.Json.Serialization;

namespace M365.API.Helper.Types
{
    public class CollectionResponse<T>
    {
        [JsonPropertyName("@odata.nextLink")]
        public string? ODataNextLink { get; set; }
        [JsonPropertyName("value")]
        public IEnumerable<T>? Value { get; set; }
    }
}
