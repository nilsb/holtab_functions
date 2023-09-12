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
using Shared;
using Shared.Models;
using Microsoft.Graph.Models;
using Newtonsoft.Json.Linq;

namespace CreateTeam
{
    public class BGCreateTeam
    {
        private readonly IConfiguration config;

        public BGCreateTeam(IConfiguration config)
        {
            this.config = config;
        }

        [FunctionName("BGCreateTeam")]
        public async Task<IActionResult> Run([HttpTrigger(AuthorizationLevel.Function, new string[] { "post" }, Route = null)] HttpRequest req,
            Microsoft.Azure.WebJobs.ExecutionContext context,
            ILogger log)
        {
            string Message = await new StreamReader(req.Body).ReadToEndAsync();
            dynamic MessageObject = JObject.Parse(Message);
            log.LogInformation($"Create team queue trigger function processed message: {Message}");
            Settings settings = new Settings(config, context, log);
            string response = string.Empty;
            Graph msGraph = new Graph(settings);
            Common common = new Common(settings, msGraph);
            FindGroupResult result = new FindGroupResult() { Success = false };

            //Parse the incoming message into JSON
            dynamic customerQueueMessage = MessageObject.MessageText != null ? MessageObject.MessageText : MessageObject;

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
                        bool createTeamResult = await common.CreateCustomerTeam(findCustomer.customer, result.group);
                    }
                    catch (Exception ex)
                    {
                        return new UnprocessableEntityObjectResult(ex.ToString());
                    }
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

            return new OkObjectResult(JsonConvert.SerializeObject(Message));
        }


    }
}
