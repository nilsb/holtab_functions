using System;
using System.IO;
using System.Security.Cryptography.X509Certificates;
using System.Security;
using System.Threading.Tasks;
using AutoMapper;
using Azure.Identity;
using Azure.Messaging.ServiceBus;
using CreateTeam.Models;
using CreateTeam.Shared;
using Microsoft.ApplicationInsights;
using Microsoft.ApplicationInsights.DataContracts;
using Microsoft.ApplicationInsights.Extensibility;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Azure.Amqp.Encoding;
using Microsoft.Azure.Cosmos.Serialization.HybridRow;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Extensions.Http;
using Microsoft.Azure.WebJobs.Host;
using Microsoft.Azure.WebJobs.ServiceBus;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;
using Microsoft.Graph;
using Newtonsoft.Json;
using PnP.Core.Services;

namespace CreateTeam
{
    public class BGCopyRootStructure
    {
        private readonly TelemetryClient telemetryClient;

        public BGCopyRootStructure(TelemetryConfiguration telemetryConfiguration)
        {
            this.telemetryClient = new TelemetryClient(telemetryConfiguration);
        }

        [FunctionName("BGCopyRootStructure")]
        public async Task<IActionResult> Run([HttpTrigger(AuthorizationLevel.Function, new string[] { "post" }, Route = null)] HttpRequest req,
            Microsoft.Azure.WebJobs.ExecutionContext context,
            ILogger log)
        {
            string Message = await new StreamReader(req.Body).ReadToEndAsync();

            log.LogInformation($"Copy root structure queue trigger function processed message: {Message}");
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
            Graph msGraph = new Graph(graphClient, log);
            Common common = new Common(graphClient, config, log, telemetryClient, msGraph);
            telemetryClient.TrackEvent(new EventTelemetry($"Got copy root structure request with message: {Message}"));

            //Parse the incoming message into JSON
            CustomerQueueMessage customerQueueMessage = JsonConvert.DeserializeObject<CustomerQueueMessage>(Message);

            //Get customer object from database
            FindCustomerResult findCustomer = common.GetCustomer(customerQueueMessage.ExternalId, customerQueueMessage.Type, customerQueueMessage.Name);

            if (findCustomer.Success && findCustomer.customer != null && findCustomer.customer != default(Customer))
            {
                Customer customer = findCustomer.customer;

                //try to copy the root structure based on type of customer
                if (await common.CopyRootStructure(customer))
                {
                    telemetryClient.TrackEvent(new EventTelemetry($"Created template folders"));
                    customer.GeneralFolderCreated = true;
                    customer.CopiedRootStructure = true;
                    common.UpdateCustomer(customer, "root structure");

                    return new OkObjectResult(JsonConvert.SerializeObject(Message));
                }
                else
                {
                    //copying the structure didn't work so try again
                    return new UnprocessableEntityObjectResult(JsonConvert.SerializeObject(Message));
                }
            }
            else
            {
                return new BadRequestObjectResult(JsonConvert.SerializeObject(Message));
            }
        }
    }
}
