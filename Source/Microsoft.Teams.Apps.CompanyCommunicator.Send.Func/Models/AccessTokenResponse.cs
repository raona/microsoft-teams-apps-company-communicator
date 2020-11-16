

namespace Microsoft.Teams.Apps.CompanyCommunicator.Send.Func.Models
{
    using Newtonsoft.Json;
    /// <summary>
    /// Response from bot token request
    /// </summary>
    public class AccessTokenResponse
    {
        [JsonProperty("token_type")]
        public string TokenType { get; set; }

        [JsonProperty("expires_in")]
        public string ExpiresIn { get; set; }

        [JsonProperty("ext_expires_in")]
        public string ExtExpiresIn { get; set; }

        [JsonProperty("access_token")]
        public string AccessToken { get; set; }
    }
}
