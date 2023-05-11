using System;
using Azure.Identity;
using Microsoft.Azure.WebJobs;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;
using Microsoft.Graph;
using System.Threading.Tasks;
using Newtonsoft.Json;
using Microsoft.AspNetCore.Http;
using Microsoft.Azure.WebJobs.Extensions.Http;
using System.IO;
using Microsoft.AspNetCore.Mvc;
using Shared;
using Shared.Models;

namespace Customers
{
    public class BGAssignPermissions
    {
        private readonly IConfiguration config;

        public BGAssignPermissions(IConfiguration config)
        {
            this.config = config;
        }

        [FunctionName("BGAssignPermissions")]
        public async Task<IActionResult> Run([HttpTrigger(AuthorizationLevel.Function, new string[] { "post" }, Route = null)] HttpRequest req,
            Microsoft.Azure.WebJobs.ExecutionContext context,
            ILogger log)
        {
            string Message = await new StreamReader(req.Body).ReadToEndAsync();

            log.LogInformation($"Assign permissions queue trigger function processed message: {Message}");
            Settings settings = new Settings(config, context, log);
            Graph msGraph = new Graph(ref settings);
            Common common = new Common(ref settings, ref msGraph);
            log.LogTrace($"Got assign permissions request with message: {Message}");

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
