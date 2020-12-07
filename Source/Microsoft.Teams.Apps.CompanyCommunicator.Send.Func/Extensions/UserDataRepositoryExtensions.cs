

namespace Microsoft.Teams.Apps.CompanyCommunicator.Send.Func.Extensions
{
    using System;
    using System.Linq;
    using System.Threading.Tasks;
    using Microsoft.Azure.Cosmos.Table;
    using Microsoft.Extensions.Logging;
    using Microsoft.Teams.Apps.CompanyCommunicator.Common.Repositories.UserData;

    public static class UserDataRepositoryExtensions
    {
        /// <summary>
        /// Get User by Mail (only first occurrence)
        /// </summary>
        /// <param name="userDataRepository">The user data repository</param>
        /// <param name="email">The user email</param>
        /// <param name="log">Azure function logger</param>
        /// <returns>A task with the user if found, otherway a task yielding null</returns>
        public static async Task<UserDataEntity> GetUserDataEntityByMailOrUPN(
            this UserDataRepository userDataRepository,
            string email,
            ILogger log)
        {
            var lowerCasedEmail = email.ToLower();
            try
            {
                // TODO: avoid hardcoding this column name
                var users = await userDataRepository.GetWithFilterAsync(
                    TableQuery.CombineFilters(
                        TableQuery.GenerateFilterCondition("Email", QueryComparisons.Equal, lowerCasedEmail),
                        TableOperators.Or,
                        TableQuery.GenerateFilterCondition("Upn", QueryComparisons.Equal, lowerCasedEmail)
                    ),
                    UserDataTableNames.UserDataPartition
                );

                return users.ElementAt(0);
            }
            catch (Exception e)
            {
                log.LogError(e, $"ERROR RETRIEVING USER BY EMAIL: {e.Message}, {e.GetType()}");
                return null;
            }
        }
    }
}
