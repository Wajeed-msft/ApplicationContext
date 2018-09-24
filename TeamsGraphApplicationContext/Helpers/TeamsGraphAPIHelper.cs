using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Threading.Tasks;
using TeamsAdmin.Models;

namespace TeamsAdmin.Helper
{
    /// <summary>
    /// Provides all the functionality for Microsft Teams Graph APIs
    /// </summary>
    public class TeamsGraphApiHelper
    {
        private string _graphApiEndpoint;
        public TeamsGraphApiHelper(string graphEndpoint)
        {
            _graphApiEndpoint = graphEndpoint;
        }

        public async Task CreateNewTeam(NewTeamDetails teamDetails, string token)
        {
            var groupId = await CreateGroupAsyn(token, teamDetails.TeamName, teamDetails.OwnerEmailId);
            if (IsValidGuid(groupId))
            {
                Console.WriteLine($"O365 Group is created for {teamDetails.TeamName}.");
                // Sometimes Team creation fails due to internal error. Added rety mechanism.
                var retryCount = 4;
                string teamId = null;
                do
                {
                    teamId = await CreateTeamAsyn(token, groupId);
                    if (IsValidGuid(teamId))
                    {
                        Console.WriteLine($"Microsoft Teams team is created for {teamDetails.TeamName}.");
                        break;
                    }
                    else
                    {
                        teamId = null;
                    }
                    retryCount--;
                    await Task.Delay(5000);
                } while (retryCount > 0);

                await CreateTeamAndChannels(token, teamDetails, teamId);
            }
        }

        private async Task CreateTeamAndChannels(string token, NewTeamDetails teamDetails, string teamId)
        {
            if (teamId != null)
            {
                // Note: Channle Info is not supported in Application context yet.
                //foreach (var channelName in teamDetails.ChannelNames)
                //{
                //    var channelId = await CreateChannel(token, teamId, channelName, channelName);
                //    if (String.IsNullOrEmpty(channelId))
                //        Console.WriteLine($"Failed to create '{channelName}' channel in '{teamDetails.TeamName}' team.");
                //}

                // Add users:
                foreach (var memberEmailId in teamDetails.MemberEmails)
                {
                    var result = await AddUserToTeam(token, teamId, memberEmailId);

                    if (!result)
                        Console.WriteLine($"Failed to add {memberEmailId} to {teamDetails.TeamName}. Check if user is already part of this team.");
                }

                Console.WriteLine($"Channels, Members Added successfully for '{teamDetails.TeamName}' team.");
            }
            else
            {
                Console.WriteLine($"Failed to create team due to internal error. Please try again later.");
            }
        }

        private async Task<bool> AddUserToTeam(string token, string teamId, string userEmailId)
        {
            var userId = await GetUserId(token, userEmailId);
            return await AddTeamMemberAsync(token, teamId, userId);
        }

        bool IsValidGuid(string guid)
        {
            Guid teamGUID;
            return Guid.TryParse(guid, out teamGUID);
        }

        public async Task<string> CreateChannel(
            string accessToken, string teamId, string channelName, string channelDescription)
        {
            string endpoint = _graphApiEndpoint + $"groups/{teamId}/team/channels";

            ChannelInfoBody channelInfo = new ChannelInfoBody()
            {
                description = channelDescription,
                displayName = channelName
            };

            return await PostRequest(accessToken, endpoint, JsonConvert.SerializeObject(channelInfo));
        }

        public async Task<string> CreateGroupAsyn(
            string accessToken, string groupName, string ownerEmailId)
        {
            string endpoint = _graphApiEndpoint + "groups/";

            GroupInfo groupInfo = new GroupInfo()
            {
                description = "Team for " + groupName,
                displayName = groupName,
                groupTypes = new string[] { "Unified" },
                mailEnabled = true,
                mailNickname = groupName.Replace(" ", "").Replace("-", "") + DateTime.Now.Second,
                securityEnabled = true,
                ownersodatabind = new[] { "https://graph.microsoft.com/v1.0/users/" + GetUserId(accessToken, ownerEmailId).Result }
            };

            return await PostRequest(accessToken, endpoint, JsonConvert.SerializeObject(groupInfo));
        }


        public async Task<bool> AddTeamMemberAsync(
            string accessToken, string teamId, string userId)
        {
            string endpoint = _graphApiEndpoint + $"groups/{teamId}/members/$ref";

            var userData = $"{{ \"@odata.id\": \"https://graph.microsoft.com/beta/directoryObjects/{userId}\" }}";

            using (var client = new HttpClient())
            {
                using (var request = new HttpRequestMessage(HttpMethod.Post, endpoint))
                {
                    request.Headers.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
                    request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", accessToken);
                    request.Content = new StringContent(userData, Encoding.UTF8, "application/json");

                    using (HttpResponseMessage response = await client.SendAsync(request))
                    {
                        if (response.IsSuccessStatusCode)
                        {

                            return true;
                        }
                        return false;
                    }
                }
            }
        }

        public async Task<string> CreateTeamAsyn(
           string accessToken, string groupId)
        {
            // This might need Retries.
            string endpoint = _graphApiEndpoint + $"groups/{groupId}/team";

            TeamSettings teamInfo = new TeamSettings()
            {
                funSettings = new Funsettings() { allowGiphy = true, giphyContentRating = "strict" },
                messagingSettings = new Messagingsettings() { allowUserEditMessages = true, allowUserDeleteMessages = true },
                memberSettings = new Membersettings() { allowCreateUpdateChannels = true }
            };
            return await PutRequest(accessToken, endpoint, JsonConvert.SerializeObject(teamInfo));
        }

        private static async Task<string> PostRequest(string accessToken, string endpoint, string groupInfo)
        {
            using (var client = new HttpClient())
            {
                using (var request = new HttpRequestMessage(HttpMethod.Post, endpoint))
                {
                    request.Headers.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
                    request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", accessToken);
                    request.Content = new StringContent(groupInfo, Encoding.UTF8, "application/json");

                    using (HttpResponseMessage response = await client.SendAsync(request))
                    {
                        if (response.IsSuccessStatusCode)
                        {

                            var createdGroupInfo = JsonConvert.DeserializeObject<ResponseData>(response.Content.ReadAsStringAsync().Result);
                            return createdGroupInfo.id;
                        }
                        return null;
                    }
                }
            }
        }

        private static async Task<string> PutRequest(string accessToken, string endpoint, string groupInfo)
        {
            using (var client = new HttpClient())
            {
                using (var request = new HttpRequestMessage(HttpMethod.Put, endpoint))
                {
                    request.Headers.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
                    request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", accessToken);
                    request.Content = new StringContent(groupInfo, Encoding.UTF8, "application/json");

                    using (HttpResponseMessage response = await client.SendAsync(request))
                    {
                        if (response.IsSuccessStatusCode)
                        {

                            var createdGroupInfo = JsonConvert.DeserializeObject<ResponseData>(response.Content.ReadAsStringAsync().Result);
                            return createdGroupInfo.id;
                        }
                        return null;
                    }
                }
            }
        }

        /// <summary>
        /// Get the current user's id from their profile.
        /// </summary>
        /// <param name="accessToken">Access token to validate user</param>
        /// <returns></returns>
        public async Task<string> GetGroupId(string accessToken, string teamName)
        {
            string endpoint = _graphApiEndpoint + $"/groups?$filter=displayName eq '{teamName}'&$select=id";

            using (var client = new HttpClient())
            {
                using (var request = new HttpRequestMessage(HttpMethod.Get, endpoint))
                {
                    request.Headers.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
                    request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", accessToken);

                    string groupId = "";
                    using (HttpResponseMessage response = await client.SendAsync(request))
                    {
                        if (response.IsSuccessStatusCode)
                        {
                            var json = JObject.Parse(await response.Content.ReadAsStringAsync());
                            try
                            {
                                groupId = json["value"].First["id"].ToString();
                            }
                            catch (Exception)
                            {
                                // Handle edge case.
                            }

                        }
                        return groupId?.Trim();
                    }
                }
            }
        }

        /// <summary>
        /// Get the current user's id from their profile.
        /// </summary>
        /// <param name="accessToken">Access token to validate user</param>
        /// <returns></returns>
        public async Task<string> GetUserId(string accessToken, string userEmailId)
        {
            string endpoint = _graphApiEndpoint + $"users/{userEmailId}";
            string queryParameter = "?$select=id";

            using (var client = new HttpClient())
            {
                using (var request = new HttpRequestMessage(HttpMethod.Get, endpoint + queryParameter))
                {
                    request.Headers.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
                    request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", accessToken);

                    string userId = "";
                    using (HttpResponseMessage response = await client.SendAsync(request))
                    {
                        if (response.IsSuccessStatusCode)
                        {
                            var json = JObject.Parse(await response.Content.ReadAsStringAsync());
                            userId = json.GetValue("id").ToString();
                        }
                        return userId?.Trim();
                    }
                }
            }
        }

        public static async Task<string> POST(string url, string body)
        {
            HttpClient httpClient = new HttpClient();

            var request = new HttpRequestMessage(HttpMethod.Post, url);
            request.Headers.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
            request.Content = new StringContent(body, Encoding.UTF8, "application/x-www-form-urlencoded");
            HttpResponseMessage response = await httpClient.SendAsync(request);
            string responseBody = await response.Content.ReadAsStringAsync();

            if (!response.IsSuccessStatusCode)
                throw new Exception(responseBody);
            return responseBody;
        }

        public class TokenResponse
        {
            public string access_token { get; set; }
        }
    }
}