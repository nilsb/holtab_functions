using System;
using System.IO;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Extensions.Http;
using Microsoft.AspNetCore.Http;
using Microsoft.Extensions.Logging;
using Newtonsoft.Json;
using Microsoft.Extensions.Configuration;
using Azure.Identity;
using Microsoft.Graph;
using CreateTeam.Shared;
using CreateTeam.Models;
using System.Web.Http;
using Microsoft.ApplicationInsights;
using Microsoft.ApplicationInsights.Extensibility;
using AutoMapper;
using Azure.Messaging.ServiceBus;
using Microsoft.ApplicationInsights.DataContracts;
using Microsoft.Azure.WebJobs.ServiceBus;
using Microsoft.Azure.Amqp.Encoding;
using Microsoft.Graph.Models;

namespace CreateTeam
{
    public class BGCreateTeam
    {
        private readonly TelemetryClient telemetryClient;

        public BGCreateTeam(TelemetryConfiguration telemetryConfiguration)
        {
            this.telemetryClient = new TelemetryClient(telemetryConfiguration);
        }

        [FunctionName("BGCreateTeam")]
        public async Task<IActionResult> Run([HttpTrigger(AuthorizationLevel.Function, new string[] { "post" }, Route = null)] HttpRequest req,
            Microsoft.Azure.WebJobs.ExecutionContext context,
            ILogger log)
        {
            string Message = await new StreamReader(req.Body).ReadToEndAsync();
            log.LogInformation($"Create team queue trigger function processed message: {Message}");

            var config = new ConfigurationBuilder()
              .SetBasePath(context.FunctionAppDirectory)
              .AddJsonFile("local.settings.json", optional: true, reloadOnChange: true)
              .AddEnvironmentVariables()
              .Build();

            string ClientID = config["ClientID"];
            string ClientSecret = config["ClientSecret"];
            string TenantID = config["TenantID"];
            string response = string.Empty;

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
            FindGroupResult result = new FindGroupResult() { Success = false };

            //Parse the incoming message into JSON
            CustomerQueueMessage customerQueueMessage = JsonConvert.DeserializeObject<CustomerQueueMessage>(Message);

            //Get customer object from database
            FindCustomerResult findCustomer = common.GetCustomer(customerQueueMessage.ExternalId, customerQueueMessage.Type, customerQueueMessage.Name);

            if (findCustomer.Success && findCustomer.customer != null && findCustomer.customer != default(Customer))
            {
                Customer customer = findCustomer.customer;

                result = await msGraph.GetGroupById(customer.GroupID);

                //if the group was found
                if (result.Success && result.group != null && result.group != default(Group))
                {
                    try
                    {
                        //try to find the team if it already exists or create it if it's missing
                        _ = await common.CreateCustomerOrSupplier(findCustomer.customer);
                    }
                    catch (Exception ex)
                    {
                        return new UnprocessableEntityObjectResult(ex.ToString());
                    }

                    return new UnprocessableEntityObjectResult(JsonConvert.SerializeObject(Message));
                }
                else
                {
                    return new NotFoundObjectResult(JsonConvert.SerializeObject(Message));
                }
            }
            else
            {
                return new BadRequestObjectResult(JsonConvert.SerializeObject(Message));
            }
        }


    }
}
