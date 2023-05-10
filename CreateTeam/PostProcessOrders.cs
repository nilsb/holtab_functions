using System;
using System.Threading.Tasks;
using Azure.Identity;
using CreateTeam.Shared;
using Microsoft.ApplicationInsights;
using Microsoft.ApplicationInsights.DataContracts;
using Microsoft.ApplicationInsights.Extensibility;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Host;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;
using Microsoft.Graph;
using Newtonsoft.Json;

namespace CreateTeam
{
    public class PostProcessOrders
    {
        private readonly TelemetryClient telemetryClient;

        public PostProcessOrders(TelemetryConfiguration telemetryConfiguration)
        {
            this.telemetryClient = new TelemetryClient(telemetryConfiguration);
        }

        [FunctionName("PostProcessOrders")]
        public async Task Run([TimerTrigger("0 */5 * * * *")]TimerInfo myTimer, [Queue("createorder"), StorageAccount("AzureWebJobsStorage")] ICollector<string> outputQueueItem, Microsoft.Azure.WebJobs.ExecutionContext context, ILogger log)
        {
            log.LogInformation($"C# Timer trigger function executed at: {DateTime.Now}");

            var config = new ConfigurationBuilder()
              .SetBasePath(context.FunctionAppDirectory)
              .AddJsonFile("local.settings.json", optional: true, reloadOnChange: true)
              .AddEnvironmentVariables()
              .Build();

            string ClientID = config["ClientID"];
            string ClientSecret = config["ClientSecret"];
            string TenantID = config["TenantID"];

            var options = new TokenCredentialOptions
            {
                AuthorityHost = AzureAuthorityHosts.AzurePublicCloud
            };

            var clientSecretCredential = new ClientSecretCredential(
                TenantID,
                ClientID,
                ClientSecret,
                options);
            var scopes = new[] { "https://graph.microsoft.com/.default" };
            var graphClient = new GraphServiceClient(clientSecretCredential, scopes);
            Graph msgraph = new Graph(graphClient, log);
            Common common = new Common(graphClient, config, log, telemetryClient, msgraph);

            var orderItems = common.GetUnhandledOrderItems();

            foreach(var order in orderItems)
            {
                telemetryClient.TrackEvent(new EventTelemetry("Putting message on order queue: " + JsonConvert.SerializeObject(order)));
                log.LogInformation("Putting message on order queue: " + JsonConvert.SerializeObject(order));
                outputQueueItem.Add(JsonConvert.SerializeObject(order));
            }
        }
    }
}
