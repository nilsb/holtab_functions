﻿using Azure.Identity;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;
using Microsoft.Graph;
using Microsoft.Identity.Client;
using StackExchange.Redis;
using static System.Formats.Asn1.AsnWriter;
using ExecutionContext = Microsoft.Azure.WebJobs.ExecutionContext;

namespace Shared.Models
{
    public class Settings
    {
        public Settings(IConfiguration config, ExecutionContext? context, ILogger? log)
        {
            this.config = config;

            if(context != null)
            {
                this.context = context;

                if(config != null)
                {
                    this.ClientID = config["ClientID"];
                    this.TenantID = config["TenantID"];
                    this.CDNTeamID = config["CDNTeamID"];
                    this.cdnSiteId = config["cdnSiteId"];
                    this.ClientSecret = config["ClientSecret"];
                    this.OrderListID = config["OrderListID"];
                    this.CustomerListID = config["CustomerListID"];
                    this.ProductionChoicesListID = config["ProductionChoicesID"];
                    this.Admins = config["Admins"];
                    this.SqlConnectionString = config["SqlConnectionString"];
                    this.redisConnectionString = config["redisConnectionString"];

                    var scopes = new[] { "https://graph.microsoft.com/.default" };

                    if (!string.IsNullOrEmpty(config["debugFlags"]))
                        this.debugFlags = Newtonsoft.Json.JsonConvert.DeserializeObject<DebugFlags>(config["debugFlags"]);

                    if (!string.IsNullOrEmpty(this.TenantID) && !string.IsNullOrEmpty(this.ClientID) && !string.IsNullOrEmpty(this.ClientSecret))
                    {
                        var options = new TokenCredentialOptions
                        {
                            AuthorityHost = AzureAuthorityHosts.AzurePublicCloud
                        };

                        var clientSecretCredential = new ClientSecretCredential(
                            this.TenantID,
                            this.ClientID,
                            this.ClientSecret,
                            options);
                        this.GraphClient = new GraphServiceClient(clientSecretCredential, scopes);
                    }

                    if (!string.IsNullOrEmpty(this.redisConnectionString))
                    {
                        this.redis = ConnectionMultiplexer.Connect(this.redisConnectionString);
                    }

                    IConfidentialClientApplication app = ConfidentialClientApplicationBuilder.Create(ClientID)
                        .WithClientSecret(ClientSecret)
                        .WithAuthority($"https://login.microsoftonline.com/{TenantID}")
                        .Build();

                    AuthenticationResult result;

                    //try
                    //{
                    //    var securePassword = new System.Security.SecureString();
                    //    foreach (char c in serviceAccountPassword)
                    //        securePassword.AppendChar(c);

                    //    result = app.

                    //    // Use result.AccessToken to call Microsoft Graph API or any other subsequent action

                    //    return new OkObjectResult($"Access Token: {result.AccessToken}");
                    //}
                    //catch (MsalException ex)
                    //{
                    //    return new BadRequestObjectResult($"Error acquiring token: {ex.Message}");
                    //}
                }
            }

            if(log != null)
            {
                this.log = log;
            }
        }

        public DebugFlags? debugFlags { get; set; }
        public GraphServiceClient? GraphClient { get; set; }
        public GraphServiceClient? DelegatedGraphClient { get; set; }
        public ConnectionMultiplexer? redis { get; set; }
        public string TenantID { get; set; } = "";
        public string ClientID { get; set; } = "";
        public string ClientSecret { get; set; } = "";
        public string CDNTeamID { get; set; } = "";
        public string cdnSiteId { get; set; } = "";
        public string OrderListID { get; set; } = "";
        public string CustomerListID { get; set; } = "";
        public string SqlConnectionString { get; set; } = "";
        public string Admins { get; set; } = "";
        public string ProductionChoicesListID { get; set; } = "";
        public string redisConnectionString { get; set; } = "";
        public IConfiguration config { get; set; }
        public ExecutionContext? context { get; set; }
        public ILogger? log { get; set; }
    }
}
