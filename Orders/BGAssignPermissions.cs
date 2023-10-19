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
            string Message = await new StreamReader(req.Body).ReadToEndAsync();
            Settings settings = new Settings(config, context, log);
            bool debug = (settings?.debugFlags?.Order?.BGAssignPermission).HasValue && (settings?.debugFlags?.Order?.BGAssignPermission).Value;
            Graph msGraph = new Graph(settings);
            Common common = new Common(settings, msGraph, debug);
            OrderMessage orderMessage = JsonConvert.DeserializeObject<OrderMessage>(Message);
            Order order = common.GetOrderFromCDN(orderMessage.ExternalId, debug);

            if(debug)
                log.LogInformation("Order BGAssignPermissions: Order assign permissions message received.");

            if (order?.Customer != null)
            {
                var groupDrive = await common.FindCustomerGroupAndDrive(order.Customer, debug);

                if (groupDrive?.Success == true && groupDrive?.customer != null)
                {
                    if (!string.IsNullOrEmpty(groupDrive.customer.DriveID))
                    {
                        if (!string.IsNullOrEmpty(orderMessage.Seller))
                        {
                            await msGraph.AddGroupOwner(orderMessage.Seller, orderMessage.CustomerGroupID, debug);
                            await msGraph.AddGroupMember(orderMessage.Seller, orderMessage.CustomerGroupID, debug);
                        }

                        if (!string.IsNullOrEmpty(orderMessage.ProjectManager))
                        {
                            await msGraph.AddGroupOwner(orderMessage.ProjectManager, orderMessage.CustomerGroupID, debug);
                            await msGraph.AddGroupMember(orderMessage.ProjectManager, orderMessage.CustomerGroupID, debug);
                        }

                        order.Handled = true;
                        common.UpdateOrCreateDbOrder(order, debug);
                    }
                }
            }

            return new OkObjectResult(JsonConvert.SerializeObject(orderMessage));
        }
    }
}

