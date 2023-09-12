using System.IO;
using System.Net.Http;
using System.Threading.Tasks;
using Shared.Models;
using Shared;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Extensions.Http;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace CreateTeam
{
    public class BGCustomerInfo
    {
        public readonly HttpClient client;
        private readonly IConfiguration config;

        public BGCustomerInfo(IConfiguration config)
        {
            this.config = config;
        }

        [FunctionName("BGCustomerInfo")]
        public async Task<IActionResult> Run([HttpTrigger(AuthorizationLevel.Function, new string[] { "post" }, Route = null)] HttpRequest req,
            Microsoft.Azure.WebJobs.ExecutionContext context,
            ILogger log)
        {
            string Message = await new StreamReader(req.Body).ReadToEndAsync();
            log.LogInformation($"Customer Information trigger function processed message: {Message}");
            Settings settings = new Settings(config, context, log);
            Graph msGraph = new Graph(settings);
            Common common = new Common(settings, msGraph);

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
