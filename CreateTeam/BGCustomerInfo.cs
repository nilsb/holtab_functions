using System;
using System.IO;
using System.Net.Http;
using System.Net.Http.Json;
using System.Text;
using System.Threading.Tasks;
using AutoMapper;
using Azure.Identity;
using Azure.Messaging.ServiceBus;
using Azure.Storage.Queues.Models;
using CreateTeam.Models;
using CreateTeam.Shared;
using Microsoft.ApplicationInsights;
using Microsoft.ApplicationInsights.DataContracts;
using Microsoft.ApplicationInsights.Extensibility;
using Microsoft.AspNet.SignalR.Infrastructure;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Azure.Functions.Worker;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Extensions.Http;
using Microsoft.Azure.WebJobs.Host;
using Microsoft.Azure.WebJobs.ServiceBus;
using Microsoft.Extensions.Azure;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;
using Microsoft.Graph;
using Newtonsoft.Json;
using RtfPipe.Tokens;

namespace CreateTeam
{
    public class BGCustomerInfo
    {
        private readonly TelemetryClient telemetryClient;

        public BGCustomerInfo(TelemetryConfiguration telemetryConfiguration)
        {
            this.telemetryClient = new TelemetryClient(telemetryConfiguration);
        }

        public readonly HttpClient client;

        [ServiceBusOutput("creategroup", Connection = "sbholtabnavConnection")]
        [FunctionName("BGCustomerInfo")]
        public async Task<IActionResult> Run([HttpTrigger(AuthorizationLevel.Function, new string[] { "post" }, Route = null)] HttpRequest req,
            Microsoft.Azure.WebJobs.ExecutionContext context,
            ILogger log)
        {
            string Message = await new StreamReader(req.Body).ReadToEndAsync();

            log.LogInformation($"Customer Information trigger function processed message: {Message}");
            var config = new ConfigurationBuilder()
              .SetBasePath(context.FunctionAppDirectory)
              .AddJsonFile("local.settings.json", optional: true, reloadOnChange: true)
              .AddEnvironmentVariables()
              .Build();

            Common common = new Common(null, config, log, telemetryClient, null);
            telemetryClient.TrackEvent(new EventTelemetry($"Got create customer request with message: {Message}"));

            //Parse the incoming message into JSON
            CustomerMessage customerMessage = JsonConvert.DeserializeObject<CustomerMessage>(Message);

            //Find the customer in the database and update the information or create it if it doesn't exist
            Customer createdCustomer = common.UpdateOrCreateDbCustomer(customerMessage);

            //Make sure the customer record exists in database by trying to find it again
            FindCustomerResult customerResult = common.GetCustomer(customerMessage.CustomerNo, customerMessage.Type, customerMessage.CustomerName);

            //If the customer was found
            if (customerResult.Success && customerResult.customer != null && customerResult.customer != default(Customer))
            {
                //put a message on the create group queue with the updated customer information
                CustomerQueueMessage customerQueueMessage = new CustomerQueueMessage();
                customerQueueMessage.ID = customerResult.customer.ID.ToString();
                customerQueueMessage.ExternalId = customerResult.customer.ExternalId;
                customerQueueMessage.Type = customerResult.customer.Type;
                customerQueueMessage.Name = customerResult.customer.Name;

                //if everything was successful we complete the message
                return new OkObjectResult(JsonConvert.SerializeObject(customerQueueMessage));
            }
            else
            {
                //if the customer was not found or couldn't be created it will be logged by previous functions
                //we abandon the message to release it from peeklock which will make it availabe to read again.
                return new UnprocessableEntityObjectResult(Message);
            }
        }
    }
}
