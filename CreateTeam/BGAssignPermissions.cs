using System;
using Azure.Identity;
using Azure.Messaging.ServiceBus;
using CreateTeam.Shared;
using Microsoft.ApplicationInsights.DataContracts;
using Microsoft.ApplicationInsights;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Host;
using Microsoft.Azure.WebJobs.ServiceBus;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;
using Microsoft.Graph;
using System.Threading.Tasks;
using Microsoft.ApplicationInsights.Extensibility;
using CreateTeam.Models;
using Newtonsoft.Json;
using Microsoft.AspNetCore.Mvc.Formatters;
using AutoMapper;
using Microsoft.Azure.Amqp.Encoding;
using Microsoft.AspNetCore.Http;
using Microsoft.Azure.WebJobs.Extensions.Http;
using System.IO;
using Microsoft.AspNetCore.Mvc;

namespace CreateTeam
{
    public class BGAssignPermissions
    {
        private readonly TelemetryClient telemetryClient;
        private readonly IConfiguration _configuration;

        public BGAssignPermissions(TelemetryConfiguration telemetryConfiguration, IConfiguration configuration)
        {
            this.telemetryClient = new TelemetryClient(telemetryConfiguration);
            _configuration = configuration;
        }

        [FunctionName("BGAssignPermissions")]
        public async Task<IActionResult> Run([HttpTrigger(AuthorizationLevel.Function, new string[] { "post" }, Route = null)] HttpRequest req,
            Microsoft.Azure.WebJobs.ExecutionContext context,
            ILogger log)
        {
            string Message = await new StreamReader(req.Body).ReadToEndAsync();

            log.LogInformation($"Assign permissions queue trigger function processed message: {Message}");
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
            telemetryClient.TrackEvent(new EventTelemetry($"Got assign permissions request with message: {Message}"));

            //Parse the incoming message into JSON
            CustomerQueueMessage customerQueueMessage = JsonConvert.DeserializeObject<CustomerQueueMessage>(Message);

            //Get customer object from database
            FindCustomerResult findCustomer = common.GetCustomer(customerQueueMessage.ExternalId, customerQueueMessage.Type, customerQueueMessage.Name);

            if(findCustomer.Success)
            {
                Customer customer = findCustomer.customer;

                //Try to find the group but assumes it was already created
                FindCustomerGroupResult findCustomerGroup = await common.FindCustomerGroupAndDrive(customer);

                if (findCustomerGroup.Success)
                {
                    //Group was found so try to add the owner
                    try
                    {
                        if (!string.IsNullOrEmpty(customer.Seller))
                        {
                            await msGraph.AddGroupMember(customer.Seller, customer.GroupID);
                        }

                        return new OkObjectResult(JsonConvert.SerializeObject(Message));
                    }
                    catch (Exception ex)
                    {
                        //Failed to add owner, dead-letter
                        return new UnprocessableEntityObjectResult(JsonConvert.SerializeObject(Message));
                    }
                }
                else
                {
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
