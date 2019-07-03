using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Threading.Tasks;
using TeamsAdmin.Helper;
using TeamsAdmin.Models;

namespace TeamsGraphApplicationContext
{
    class Program
    {
        static void Main(string[] args)
        {
            // Setup your application here: https://developer.microsoft.com/en-us/graph/docs/concepts/auth_v2_service
            string tenant = ""; // ex: blrdev.onmicrosoft.com
            string appId = "";
            string appSecret = "";
            string graphEndPoint = "https://graph.microsoft.com/beta/";

            // One time process for Admin consent. 
            GetOneTimeAdminConsent(tenant, appId);

            string accessToken = GetAccessToken(tenant, appId, appSecret);

            TeamsGraphApiHelper helper = new TeamsGraphApiHelper(graphEndPoint);

            var channelId = "YourChannelId";
            var teamId = "YourTeamId";
            helper.AddWebsiteTabToChannel(accessToken, teamId, channelId).Wait();

            Console.WriteLine("Added Tab Created successfully");
            Console.ReadLine();
            
            //helper.CreateNewTeam(new NewTeamDetails()
            //{
            //    TeamName = "Application Context Test 2",
            //    OwnerEmailId = "pippen@blrdev.onmicrosoft.com",
            //    // ChannelNames = new List<string>() { "Announcements", "Dev Discussion" }, // Currently channel can't be created in App Context.
            //    MemberEmails = new List<string>() { "pippen@blrdev.onmicrosoft.com", "olo@blrdev.onmicrosoft.com", "poppy@blrdev.onmicrosoft.com" },
            //}, accessToken).Wait();

        }

        private static string GetAccessToken(string tenant, string appId, string appSecret)
        {
            string response = TeamsGraphApiHelper.POST($"https://login.microsoftonline.com/{tenant}/oauth2/v2.0/token",
                              $"grant_type=client_credentials&client_id={appId}&client_secret={appSecret}"
                              + "&scope=https%3A%2F%2Fgraph.microsoft.com%2F.default").Result;

            string accessToken = JsonConvert.DeserializeObject<TeamsGraphApiHelper.TokenResponse>(response).access_token;
            return accessToken;
        }

        private static void GetOneTimeAdminConsent(string tenant, string appId)
        {
            var adminLoginUrl = $"https://login.microsoftonline.com/{tenant}/adminconsent?client_id={appId}&state=12345&redirect_uri=http%3A%2F%2Flocalhost%2Fmyapp%2Fpermissions";
            Process.Start(adminLoginUrl); // THis is needed first time only 
            Console.WriteLine("Press enter once the admin consent is completed");
            Console.ReadLine();// Wait for user to finish login
        }

    }
}
