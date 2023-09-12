using System.IO;
using System.Net;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Extensions.Http;
using Microsoft.Azure.WebJobs.Extensions.OpenApi.Core.Attributes;
using Microsoft.Azure.WebJobs.Extensions.OpenApi.Core.Enums;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;
using Microsoft.OpenApi.Models;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using Shared;
using Shared.Models;

namespace Orders
{
    public class BGAssignPermissions
    {
        private readonly IConfiguration config;

        public BGAssignPermissions(IConfiguration _config)
        {
            config = _config;
        }

        [FunctionName("BGAssignPermissions")]
        public async Task<IActionResult> Run(
            [HttpTrigger(AuthorizationLevel.Function, new string[] { "post" }, Route = null)] HttpRequest req,
            Microsoft.Azure.WebJobs.ExecutionContext context,
            ILogger log)
        {
            log.LogInformation("Order assign permissions message received.");

            string Message = await new StreamReader(req.Body).ReadToEndAsync();
            Settings settings = new Settings(config, context, log);
            Graph msGraph = new Graph(settings);
            Common common = new Common(settings, msGraph);
            dynamic orderMessage = JObject.Parse(Message);
            Order order = common.GetOrderFromCDN(orderMessage.No);

            if (order?.Customer != null)
            {
                var groupDrive = await common.FindCustomerGroupAndDrive(order.Customer);

                if (groupDrive?.Success == true && groupDrive?.customer != null)
                {
                    if (!string.IsNullOrEmpty(groupDrive.customer.DriveID))
                    {
                        if (!string.IsNullOrEmpty(orderMessage.Seller))
                        {
                            await msGraph.AddGroupOwner(orderMessage.Seller, orderMessage.CustomerGroupID);
                        }

                        if (!string.IsNullOrEmpty(orderMessage.ProjectManager))
                        {
                            await msGraph.AddGroupOwner(orderMessage.ProjectManager, orderMessage.CustomerGroupID);
                        }
                    }
                }
            }

            return new OkObjectResult(JsonConvert.SerializeObject(orderMessage));
        }
    }
}

