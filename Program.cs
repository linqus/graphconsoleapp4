using System;
using System.Collections.Generic;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Security;
using System.Threading.Tasks;
using Microsoft.Identity.Client;
using Microsoft.Graph;
using Microsoft.Extensions.Configuration;
using Helpers;

namespace graphconsoleapp
{
    public class Program
    {

        private static IConfigurationRoot? LoadAppSettings()
        {
            try
            {
                var config = new ConfigurationBuilder()
                                .SetBasePath(System.IO.Directory.GetCurrentDirectory())
                                .AddJsonFile("appsettings.json", false, true)
                                .Build();

                if (string.IsNullOrEmpty(config["applicationId"]) || string.IsNullOrEmpty(config["tenantId"]))
                {
                    return null;
                }

                return config;
            }
            catch (System.IO.FileNotFoundException)
            {
                return null;
            }
        }

        private static IAuthenticationProvider CreateAuthorizationProvider(IConfigurationRoot config)
        {
            var clientId = config["applicationId"];
            var authority = $"https://login.microsoftonline.com/{config["tenantId"]}/v2.0";

            List<string> scopes = new List<string>();
            scopes.Add("https://graph.microsoft.com/.default");

            var cca = PublicClientApplicationBuilder.Create(clientId)
                                                    .WithAuthority(authority)
                                                    .WithDefaultRedirectUri()
                                                    .Build();
            return MsalAuthenticationProvider.GetInstance(cca, scopes.ToArray());
        }


        private static GraphServiceClient GetAuthenticatedGraphClient(IConfigurationRoot config)
        {
            var authenticationProvider = CreateAuthorizationProvider(config);
            var graphClient = new GraphServiceClient(authenticationProvider);
            return graphClient;
        }


        public static void Main(string[] args)
        {

            var config = LoadAppSettings();
            if (config == null)
            {
                Console.WriteLine("Invalid appsettings.json file.");
                return;
            }
            var client = GetAuthenticatedGraphClient(config);
            var profileResponse = client.Me.Request().GetAsync().Result;
            Console.WriteLine("Hello " + profileResponse.DisplayName);

            // request 1 - all groups
            /*             Console.WriteLine("\n\nREQUEST 1 - ALL GROUPS:");
                        var requestAllGroups = client.Groups.Request();
                        var resultsAllGroups = requestAllGroups.GetAsync().Result;
                        foreach (var group in resultsAllGroups)
                        {
                            Console.WriteLine(group.Id + ": " + group.DisplayName + " <" + group.Mail + ">");
                        }

                        Console.WriteLine("\nGraph Request:");
                        Console.WriteLine(requestAllGroups.GetHttpRequestMessage().RequestUri); */



            var groupId = "5fd37035-bb3f-4264-b0b5-c9cb3707f7c5";
            // request 2 - one group
/*             Console.WriteLine("\n\nREQUEST 2 - ONE GROUP:");
            var requestGroup = client.Groups[groupId].Request();
            var resultsGroup = requestGroup.GetAsync().Result;
            Console.WriteLine(resultsGroup.Id + ": " + resultsGroup.DisplayName + " <" + resultsGroup.Mail + ">");

            Console.WriteLine("\nGraph Request:");
            Console.WriteLine(requestGroup.GetHttpRequestMessage().RequestUri); */

            // request 4 - group members
            Console.WriteLine("\n\nREQUEST 4 - GROUP MEMBERS:");
            var requestGroupMembers = client.Groups[groupId].Members.Request();
            var resultsGroupMembers = requestGroupMembers.GetAsync().Result;
            foreach (var member in resultsGroupMembers)
            {
                var memberUser = member as Microsoft.Graph.User;
                if (memberUser != null)
                {
                    Console.WriteLine(memberUser.Id + ": " + memberUser.DisplayName + " <" + memberUser.Mail + ">");
                }
            }

            Console.WriteLine("\nGraph Request:");
            Console.WriteLine(requestGroupMembers.GetHttpRequestMessage().RequestUri);



        }

    }
}