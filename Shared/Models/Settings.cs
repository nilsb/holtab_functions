﻿using Azure.Identity;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;
using Microsoft.Graph;
using ExecutionContext = Microsoft.Azure.WebJobs.ExecutionContext;

namespace Shared.Models
{
    public class Settings
    {
        public Settings(ExecutionContext? context, ILogger? log)
        {
            if(context != null)
            {
                var config = new ConfigurationBuilder()
                  .SetBasePath(context.FunctionAppDirectory)
                  .AddJsonFile("local.settings.json", optional: true, reloadOnChange: true)
                  .AddEnvironmentVariables()
                  .Build();

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
                        var scopes = new[] { "https://graph.microsoft.com/.default" };
                        this.GraphClient = new GraphServiceClient(clientSecretCredential, scopes);
                    }
                }
            }

            if(log != null)
            {
                this.log = log;
            }
        }

        public GraphServiceClient? GraphClient { get; set; }
        public string? TenantID { get; set; }
        public string? ClientID { get; set; }
        public string? ClientSecret { get; set; }
        public string? CDNTeamID { get; set; }
        public string? cdnSiteId { get; set; }
        public string? OrderListID { get; set; }
        public string? CustomerListID { get; set; }
        public string? SqlConnectionString { get; set; }
        public string? Admins { get; set; }
        public string? ProductionChoicesListID { get; set; }
        public ExecutionContext? context { get; set; }
        public ILogger? log { get; set; }
    }
}