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
        public string CDN2SiteId { get; set; } = "d5f6e456-8705-47ca-ab47-668ff4de20ff";
        public string CDN2LibraryId { get; set; } = "ba1fe0ef-2003-479f-88cc-d3dfaf95463e";
        public string cdnSiteId { get; set; } = "";
        public string OrderListID { get; set; } = "";
        public string CustomerListID { get; set; } = "";
        public string SqlConnectionString { get; set; } = "";
        public string Admins { get; set; } = "";
        public string ProductionChoicesListID { get; set; } = "";
        public string InkopGroupId { get; set; } = "23ee624d-53ed-425e-bd75-a7c6d340ef3f";
        public string InkopSiteId { get; set; } = "holtab.sharepoint.com,d5f6e456-8705-47ca-ab47-668ff4de20ff,2e52127d-db02-45ad-97f4-c052092104e8";
        public string InkopLibraryId { get; set; } = "14747efa-f13a-4cf7-afed-b8f4cdc95e04";
        public string InkopDriveId { get; set; } = "b!VuT21QWHykerR2aP9N4g_30SUi4C261Fl_TAUgkhBOjBEO22eSTxQI9HeXbOaIqG";
        public string InkopFolderId { get; set; } = "01RLQNMH5H3WODIGSDMFC3J5ZPUDHVV2XI";
        public string InkopParentId { get; set; } = "01RLQNMH56Y2GOVW7725BZO354PWSELRRZ";
        public string redisConnectionString { get; set; } = "";
        public IConfiguration config { get; set; }
        public ExecutionContext? context { get; set; }
        public ILogger? log { get; set; }
    }
}
