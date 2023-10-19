using System.IO;
using System.Threading.Tasks;
using Azure.Identity;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Extensions.Http;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;
using Microsoft.Graph;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using Shared;
using Shared.Models;

namespace CreateTeam
{
    public class BGCopyRootStructure
    {
        private readonly IConfiguration config;

        public BGCopyRootStructure(IConfiguration config)
        {
            this.config = config;
        }

        [FunctionName("BGCopyRootStructure")]
        public async Task<IActionResult> Run([HttpTrigger(AuthorizationLevel.Function, new string[] { "post" }, Route = null)] HttpRequest req,
            Microsoft.Azure.WebJobs.ExecutionContext context,
            ILogger log)
        {
            Settings settings = new Settings(config, context, log);
            string Message = await new StreamReader(req.Body).ReadToEndAsync();
            bool debug = (settings?.debugFlags?.Customer?.BGCopyRootStructure).HasValue && (settings?.debugFlags?.Customer?.BGCopyRootStructure).Value;
            Graph msGraph = new Graph(settings);
            Common common = new Common(settings, msGraph, debug);

            if(debug)
                log.LogInformation($"Customer BGCopyRootStructure: Copy root structure queue trigger function processed message: {Message}");

            //Parse the incoming message into JSON
            CustomerQueueMessage customerQueueMessage = JsonConvert.DeserializeObject<CustomerQueueMessage>(Message);

            //Get customer object from database
            FindCustomerResult findCustomer = common.GetCustomer(customerQueueMessage.ExternalId, customerQueueMessage.Type, customerQueueMessage.Name, debug);

            if (findCustomer.Success && findCustomer.customer != null && findCustomer.customer != default(Customer))
            {
                Customer customer = findCustomer.customer;

                //try to copy the root structure based on type of customer
                if (await common.CopyRootStructure(customer, debug))
                {
                    if(debug)
                        log.LogInformation($"Customer BGCopyRootStructure: Created template folders");

                    customer.GeneralFolderCreated = true;
                    customer.CopiedRootStructure = true;
                    common.UpdateCustomer(customer, "root structure", debug);

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
