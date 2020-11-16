
namespace Microsoft.Teams.Apps.CompanyCommunicator.Send.Func.Services
{
    using System.Collections.Generic;
    using System.Linq;
    using System.Text;
    using System.Threading.Tasks;
    using Microsoft.Azure.ServiceBus;
    using Microsoft.Azure.ServiceBus.Core;
    using Microsoft.Extensions.Configuration;
    using Microsoft.Extensions.Logging;
    using Microsoft.Teams.Apps.CompanyCommunicator.Common.Repositories.TeamData;
    using Microsoft.Teams.Apps.CompanyCommunicator.Common.Repositories.UserData;
    using Microsoft.Teams.Apps.CompanyCommunicator.Send.Func.Models;
    using Newtonsoft.Json;

    public class UserService
    {
        private IConfiguration configuration;
        private UserDataRepository userDataRepository;

        public UserService(IConfiguration configuration, UserDataRepository userDataRepository)
        {
            this.configuration = configuration;
            this.userDataRepository = userDataRepository;
        }

        public async Task notifyAllUsers(Newtonsoft.Json.Linq.JObject activityToSend)
        {
            List<UserDataEntity> deDuplicatedReceiverEntities = new List<UserDataEntity>();

            var usersUserDataEntityDictionary = await this.GetUserDataDictionaryAsync();
            deDuplicatedReceiverEntities.AddRange(usersUserDataEntityDictionary.Select(kvp => kvp.Value));

            var totalMessageCount = deDuplicatedReceiverEntities.Count;

            var allServiceBusMessages = deDuplicatedReceiverEntities
                .Select(userDataEntity =>
                {
                    if (userDataEntity.Upn == null || userDataEntity.Upn == string.Empty)
                    {
                        return null;
                    }

                    var queueMessageContent = new ServiceBusNotificationQueueMessageContent
                    {
                        Email = userDataEntity.Upn,
                        ActivityToSend = activityToSend,
                    };
                    var messageBody = JsonConvert.SerializeObject(queueMessageContent);
                    return new Message(Encoding.UTF8.GetBytes(messageBody));
                })
                .Where(m => m != null)
                .ToList();

            // Create batches to send to the service bus
            var serviceBusBatches = new List<List<Message>>();

            var totalNumberMessages = allServiceBusMessages.Count;
            var batchSize = 100;
            var numberOfCompleteBatches = totalNumberMessages / batchSize;
            var numberMessagesInIncompleteBatch = totalNumberMessages % batchSize;

            for (var i = 0; i < numberOfCompleteBatches; i++)
            {
                var startingIndex = i * batchSize;
                var batch = allServiceBusMessages.GetRange(startingIndex, batchSize);
                serviceBusBatches.Add(batch);
            }

            if (numberMessagesInIncompleteBatch != 0)
            {
                var incompleteBatchStartingIndex = numberOfCompleteBatches * batchSize;
                var incompleteBatch = allServiceBusMessages.GetRange(
                    incompleteBatchStartingIndex,
                    numberMessagesInIncompleteBatch);
                serviceBusBatches.Add(incompleteBatch);
            }

            string serviceBusConnectionString = this.configuration["ServiceBusConnection"];
            string queueName = "company-communicator-notifier";
            var messageSender = new MessageSender(serviceBusConnectionString, queueName);

            // Send batches of messages to the service bus
            foreach (var batch in serviceBusBatches)
            {
                await messageSender.SendAsync(batch);
            }
        }

        private async Task<Dictionary<string, UserDataEntity>> GetUserDataDictionaryAsync()
        {
            var userDataEntities = await this.userDataRepository.GetAllAsync();
            var result = new Dictionary<string, UserDataEntity>();
            foreach (var userDataEntity in userDataEntities)
            {
                result.Add(userDataEntity.AadId, userDataEntity);
            }

            return result;
        }
    }
}
