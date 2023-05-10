using System;
using System.IO;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Extensions.Http;
using Microsoft.AspNetCore.Http;
using Microsoft.Extensions.Logging;
using Newtonsoft.Json;
using CreateTeam.Models;
using CreateTeam.Shared;
using Microsoft.ApplicationInsights;
using Microsoft.ApplicationInsights.Extensibility;
using Microsoft.Extensions.Configuration;
using System.Collections.Generic;
using System.Linq;

namespace CreateTeam
{
    public class GetCustomer
    {
        private readonly TelemetryClient telemetryClient;

        public GetCustomer(TelemetryConfiguration telemetryConfiguration)
        {
            this.telemetryClient = new TelemetryClient(telemetryConfiguration);
        }

        [FunctionName("GetCustomer")]
        public async Task<IActionResult> Run(
            [HttpTrigger(AuthorizationLevel.Anonymous, "get", "post", Route = null)] HttpRequest req,
            Microsoft.Azure.WebJobs.ExecutionContext context,
            ILogger log)
        {
            var config = new ConfigurationBuilder()
              .SetBasePath(context.FunctionAppDirectory)
              .AddJsonFile("local.settings.json", optional: true, reloadOnChange: true)
              .AddEnvironmentVariables()
              .Build();

            string SqlConnectionString = config["SqlConnectionString"];
            Customer foundCustomer = null;

            //tanslate the message sent into an object we can use
            string requestBody = await new StreamReader(req.Body).ReadToEndAsync();
            CustomerMessage msg = JsonConvert.DeserializeObject<CustomerMessage>(requestBody);

            //Try to fetch the customer from the database
            List<Customer> foundCustomers = Services.GetCustomerFromDB(msg.CustomerNo, msg.Type, SqlConnectionString);

            //If the customer was not found we return a 404
            if (foundCustomers == null || foundCustomers.Count <= 0)
            {
                return new NotFoundResult(); //404
            }

            if(foundCustomers.Count > 0)
            {
                if(foundCustomers.Any(c => c.Name == msg.CustomerName))
                {
                    foundCustomer = foundCustomers[0];
                }
                else
                {
                    //If the customer was found but the name is a mismatch
                    return new UnprocessableEntityResult(); //422
                }
            }

            //If the customer was found serialize it to a JSON string so we can use it in the logic app
            string responseMessage = JsonConvert.SerializeObject(foundCustomer);

            return new OkObjectResult(responseMessage); //200
        }
    }
}
