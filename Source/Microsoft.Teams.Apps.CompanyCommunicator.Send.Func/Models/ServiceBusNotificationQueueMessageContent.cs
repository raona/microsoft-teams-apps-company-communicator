

namespace Microsoft.Teams.Apps.CompanyCommunicator.Send.Func.Models
{
    using Newtonsoft.Json.Linq;

    public class ServiceBusNotificationQueueMessageContent
    {
        public string Email { get; set; }

        public JObject ActivityToSend { get; set; }

        public bool? SendAllUsers { get; set; }
    }
}
