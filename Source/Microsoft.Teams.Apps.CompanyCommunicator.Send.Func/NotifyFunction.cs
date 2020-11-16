

namespace Microsoft.Teams.Apps.CompanyCommunicator.Send.Func
{
    using System;
    using System.Collections.Generic;
    using System.Net;
    using System.Net.Http;
    using System.Net.Http.Headers;
    using System.Text;
    using System.Threading.Tasks;
    using Microsoft.Azure.Cosmos.Table;
    using Microsoft.Azure.WebJobs;
    using Microsoft.Extensions.Configuration;
    using Microsoft.Extensions.Logging;
    using Microsoft.Teams.Apps.CompanyCommunicator.Common.Repositories.TeamData;
    using Microsoft.Teams.Apps.CompanyCommunicator.Common.Repositories.UserData;
    using Microsoft.Teams.Apps.CompanyCommunicator.Send.Func.Extensions;
    using Microsoft.Teams.Apps.CompanyCommunicator.Send.Func.Models;
    using Microsoft.Teams.Apps.CompanyCommunicator.Send.Func.Services;
    using Newtonsoft.Json;

    public class NotifierFunction
    {
        private static string botAccessToken;
        private static DateTime botAccessTokenExpiration;
        private static IConfiguration configuration;
        private static HttpClient httpClient;
        private static UserService userService;
        private UserDataRepository userDataRepository;
        private TeamDataRepository teamDataRepository;

        public NotifierFunction(
            UserDataRepository userDataRepository,
            TeamDataRepository teamDataRepository)
        {
            this.userDataRepository = userDataRepository ?? throw new ArgumentNullException(nameof(userDataRepository));
            this.teamDataRepository = teamDataRepository ?? throw new ArgumentNullException(nameof(teamDataRepository));
        }

        [FunctionName("NotifyFunction")]
        public async Task Run(
            [ServiceBusTrigger("company-communicator-notifier", Connection = "ServiceBusConnection")]string myQueueItem, ILogger log
            )
        {
            log.LogInformation($"C# ServiceBus queue trigger function processed message: {myQueueItem}");

            NotifierFunction.configuration = NotifierFunction.configuration ??
                new ConfigurationBuilder()
                    .AddEnvironmentVariables()
                    .Build();

            var messageContent = JsonConvert.DeserializeObject<ServiceBusNotificationQueueMessageContent>(myQueueItem);
            log.LogInformation($"Message Content: {messageContent}");

            var totalNumberOfThrottles = 0;

            try
            {
                NotifierFunction.userService = NotifierFunction.userService
                    ?? new UserService(NotifierFunction.configuration, this.userDataRepository);

                if (messageContent.SendAllUsers.HasValue && messageContent.SendAllUsers.Value)
                {
                    // this will retrieve all users and resend the message to the service bus queue,
                    // to be triggered again with the proper email and avoid running lengthy operations.
                    try
                    {
                        await userService.notifyAllUsers(messageContent.ActivityToSend);
                    }
                    catch (Exception ex)
                    {
                        log.LogError(ex.Message);
                    }

                    return;
                }

                var user = await this.userDataRepository.GetUserDataEntityByMailOrUPN(messageContent.Email, log);

                log.LogInformation($"User Entity: {user}");

                // Simply initialize the variable for certain build environments and versions
                var maxNumberOfAttempts = 0;

                // If parsing fails, out variable is set to 0, so need to set the default
                if (!int.TryParse(NotifierFunction.configuration["MaxNumberOfAttempts"], out maxNumberOfAttempts))
                {
                    maxNumberOfAttempts = 1;
                }

                NotifierFunction.httpClient = NotifierFunction.httpClient
                    ?? new HttpClient();

                if (NotifierFunction.botAccessToken == null
                        || NotifierFunction.botAccessTokenExpiration == null
                        || DateTime.UtcNow > NotifierFunction.botAccessTokenExpiration)
                {
                    await this.FetchTokenAsync(NotifierFunction.configuration, NotifierFunction.httpClient);
                }


                var incomingUserDataEntity = user;
                var incomingConversationId = incomingUserDataEntity.ConversationId;

                var conversationId = string.IsNullOrWhiteSpace(incomingConversationId)
                    ? incomingUserDataEntity?.ConversationId
                    : incomingConversationId;

                Task saveUserDataEntityTask = Task.CompletedTask;
                Task saveSentNotificationDataTask = Task.CompletedTask;
                Task setDelayTimeAndSendDelayedRetryTask = Task.CompletedTask;

                if (!string.IsNullOrWhiteSpace(conversationId))
                {
                    incomingUserDataEntity.ConversationId = conversationId;

                    // Check if message is intended for a team
                    if (!conversationId.StartsWith("19:"))
                    {
                        incomingUserDataEntity.PartitionKey = UserDataTableNames.UserDataPartition;
                        incomingUserDataEntity.RowKey = incomingUserDataEntity.AadId;

                        var operation = TableOperation.InsertOrMerge(incomingUserDataEntity);

                        saveUserDataEntityTask = this.userDataRepository.Table.ExecuteAsync(operation);
                    }
                }
                else
                {
                    var isCreateConversationThrottled = false;

                    for (int i = 0; i < maxNumberOfAttempts; i++)
                    {
                        var createConversationUrl = $"{incomingUserDataEntity.ServiceUrl}v3/conversations";
                        using (var requestMessage = new HttpRequestMessage(HttpMethod.Post, createConversationUrl))
                        {
                            requestMessage.Headers.Authorization = new AuthenticationHeaderValue(
                                "Bearer",
                                NotifierFunction.botAccessToken);

                            var payloadString = "{\"bot\": { \"id\": \"28:" + NotifierFunction.configuration["MicrosoftAppId"] + "\"},\"isGroup\": false, \"tenantId\": \"" + incomingUserDataEntity.TenantId + "\", \"members\": [{\"id\": \"" + incomingUserDataEntity.UserId + "\"}]}";
                            requestMessage.Content = new StringContent(payloadString, Encoding.UTF8, "application/json");

                            using (var sendResponse = await NotifierFunction.httpClient.SendAsync(requestMessage))
                            {
                                if (sendResponse.StatusCode == HttpStatusCode.Created)
                                {
                                    var jsonResponseString = await sendResponse.Content.ReadAsStringAsync();
                                    dynamic resp = JsonConvert.DeserializeObject(jsonResponseString);

                                    incomingUserDataEntity.PartitionKey = UserDataTableNames.UserDataPartition;
                                    incomingUserDataEntity.RowKey = incomingUserDataEntity.AadId;
                                    incomingUserDataEntity.ConversationId = resp.id;

                                    var operation = TableOperation.InsertOrMerge(incomingUserDataEntity);

                                    saveUserDataEntityTask = this.userDataRepository.Table.ExecuteAsync(operation);

                                    isCreateConversationThrottled = false;

                                    break;
                                }
                                else if (sendResponse.StatusCode == HttpStatusCode.TooManyRequests)
                                {
                                    isCreateConversationThrottled = true;

                                    totalNumberOfThrottles++;

                                    // Do not delay if already attempted the maximum number of attempts.
                                    if (i != maxNumberOfAttempts - 1)
                                    {
                                        var random = new Random();
                                        await Task.Delay(random.Next(500, 1500));
                                    }
                                }
                            }
                        }
                    }

                    if (isCreateConversationThrottled)
                    {
                        //TODO: We don't have an equivalent for this without notifications
                        //await this.SetDelayTimeAndSendDelayedRetry(NotificationSenderQueueFunction.configuration, messageContent);

                        return;
                    }
                }

                var isSendMessageThrottled = false;

                for (int i = 0; i < maxNumberOfAttempts; i++)
                {
                    var conversationUrl = $"{incomingUserDataEntity.ServiceUrl}v3/conversations/{incomingUserDataEntity.ConversationId}/activities";
                    using (var requestMessage = new HttpRequestMessage(HttpMethod.Post, conversationUrl))
                    {
                        requestMessage.Headers.Authorization = new AuthenticationHeaderValue(
                            "Bearer",
                            NotifierFunction.botAccessToken);

                        var attachmentJsonString = messageContent.ActivityToSend.ToString();
                        var messageString = "{\"type\":\"message\",\"attachments\":[{\"contentType\":\"application/vnd.microsoft.card.adaptive\",\"content\": " + attachmentJsonString + " }]}";
                        requestMessage.Content = new StringContent(messageString, Encoding.UTF8, "application/json");

                        using (var sendResponse = await NotifierFunction.httpClient.SendAsync(requestMessage))
                        {
                            if (sendResponse.StatusCode == HttpStatusCode.Created)
                            {
                                log.LogInformation("MESSAGE SENT SUCCESSFULLY");

                                isSendMessageThrottled = false;

                                break;
                            }
                            else if (sendResponse.StatusCode == HttpStatusCode.TooManyRequests)
                            {
                                log.LogError("MESSAGE THROTTLED");

                                isSendMessageThrottled = true;

                                totalNumberOfThrottles++;

                                // Do not delay if already attempted the maximum number of attempts.
                                if (i != maxNumberOfAttempts - 1)
                                {
                                    var random = new Random();
                                    await Task.Delay(random.Next(500, 1500));
                                }
                            }
                            else
                            {
                                log.LogError($"MESSAGE FAILED: {sendResponse.StatusCode}");

                                return;
                            }
                        }
                    }
                }

                if (isSendMessageThrottled)
                {
                    //TODO: We don't have an equivalent for this without notifications
                    //setDelayTimeAndSendDelayedRetryTask =
                    //    this.SetDelayTimeAndSendDelayedRetry(NotificationSenderQueueFunction.configuration, messageContent);
                }

                await Task.WhenAll(
                    saveUserDataEntityTask,
                    saveSentNotificationDataTask,
                    setDelayTimeAndSendDelayedRetryTask);

            }
            catch (Exception e)
            {
                log.LogError(e, $"ERROR: {e.Message}, {e.GetType()}");
                //var statusCodeToStore = HttpStatusCode.Continue;
                //if (deliveryCount >= NotificationSenderQueueFunction.MaxDeliveryCountForDeadLetter)
                //{
                //    statusCodeToStore = HttpStatusCode.InternalServerError;
                //}

                //await this.SaveSentNotificationData(
                //    messageContent.NotificationId,
                //    messageContent.UserDataEntity.AadId,
                //    totalNumberOfThrottles,
                //    isStatusCodeFromCreateConversation: false,
                //    statusCode: statusCodeToStore);

                //throw e;
            }

        }


        private async Task FetchTokenAsync(
            IConfiguration configuration,
            HttpClient httpClient)
        {
            var values = new Dictionary<string, string>
                {
                    { "grant_type", "client_credentials" },
                    { "client_id", configuration["MicrosoftAppId"] },
                    { "client_secret", configuration["MicrosoftAppPassword"] },
                    { "scope", "https://api.botframework.com/.default" },
                };
            var content = new FormUrlEncodedContent(values);

            using (var tokenResponse = await httpClient.PostAsync("https://login.microsoftonline.com/botframework.com/oauth2/v2.0/token", content))
            {
                if (tokenResponse.StatusCode == HttpStatusCode.OK)
                {
                    var accessTokenContent = await tokenResponse.Content.ReadAsAsync<AccessTokenResponse>();

                    NotifierFunction.botAccessToken = accessTokenContent.AccessToken;

                    var expiresInSeconds = 121;

                    // If parsing fails, out variable is set to 0, so need to set the default
                    if (!int.TryParse(accessTokenContent.ExpiresIn, out expiresInSeconds))
                    {
                        expiresInSeconds = 121;
                    }

                    // Remove two minutes in order to have a buffer amount of time.
                    NotifierFunction.botAccessTokenExpiration = DateTime.UtcNow + TimeSpan.FromSeconds(expiresInSeconds - 120);
                }
                else
                {
                    throw new Exception("Error fetching bot access token.");
                }
            }
        }
    }
}
