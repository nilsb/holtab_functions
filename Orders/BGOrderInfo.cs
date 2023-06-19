using System.IO;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Extensions.Http;
using Microsoft.AspNetCore.Http;
using Microsoft.Extensions.Logging;
using Newtonsoft.Json;
using Microsoft.Extensions.Configuration;
using Shared.Models;
using Shared;
using Microsoft.Azure.WebJobs.Extensions.OpenApi.Core.Attributes;
using Microsoft.Azure.WebJobs.Extensions.OpenApi.Core.Enums;
using Microsoft.OpenApi.Models;
using System.Net;

namespace Orders
{
    public class BGOrderInfo
    {
        private readonly ILogger<BGOrderInfo> log;
        private readonly IConfiguration config;

        public BGOrderInfo(ILogger<BGOrderInfo> _log, IConfiguration _config)
        {
            log = _log;
            config = _config;
        }

        [FunctionName("BGOrderInfo")]
        public async Task<IActionResult> Run(
            [HttpTrigger(AuthorizationLevel.Function, new string[] { "post" }, Route = null)] HttpRequest req,
            Microsoft.Azure.WebJobs.ExecutionContext context)
        {
            string Message = await new StreamReader(req.Body).ReadToEndAsync();
            log.LogInformation($"Order Information trigger function processed message: {Message}");
            Settings settings = new Settings(config, context, log);
            Graph msGraph = new Graph(settings);
            Common common = new Common(settings, msGraph);
            OrderMessage orderMessage = JsonConvert.DeserializeObject<OrderMessage>(Message);
            var newOrder = default(Order);

            //Find and update or create the order database post
            Order cdnItem = common.GetOrderFromCDN(orderMessage.No);

            if(cdnItem != null)
            {
                log.LogInformation("Found order in Database");
                if (string.IsNullOrEmpty(cdnItem.ExternalId))
                    cdnItem.ExternalId = orderMessage.No; //backwards compatibility

                cdnItem.AdditionalInfo = orderMessage.AdditionalInfo;
                cdnItem.Seller = orderMessage.Seller;
                cdnItem.ProjectManager = orderMessage.ProjectManager;
                cdnItem.Customer = null;
                cdnItem.CustomerNo = orderMessage.CustomerNo;
                cdnItem.CustomerType = orderMessage.CustomerType;
                cdnItem.Type = orderMessage.Type;
                newOrder = common.UpdateOrCreateDbOrder(cdnItem);
                log.LogInformation("Updated order in Database");
            }
            else
            {
                newOrder = new Order() { 
                    No = orderMessage.No,
                    ExternalId = orderMessage.No,
                    AdditionalInfo = orderMessage.AdditionalInfo,
                    CustomerNo = orderMessage.CustomerNo,
                    CustomerType = orderMessage.CustomerType,
                    Seller = orderMessage.Seller,
                    Type = orderMessage.Type,
                    ProjectManager = orderMessage.ProjectManager
                };

                newOrder = common.UpdateOrCreateDbOrder(newOrder);
                log.LogInformation("Created order in Database");
            }

            if (newOrder != null)
            {
                log.LogInformation("Order was created or updated in Database");
                orderMessage.ExternalId = newOrder.ExternalId;

                if(newOrder.Customer != null)
                {
                    if (!string.IsNullOrEmpty(newOrder.Customer.GroupID))
                    {
                        orderMessage.CustomerGroupID = newOrder.Customer.GroupID;
                    }

                    if (!string.IsNullOrEmpty(newOrder.Customer.ExternalId))
                    {
                        orderMessage.CustomerExternalId = newOrder.Customer.ExternalId;
                    }
                }
            }
            else
            {
                log.LogInformation("Could not create or update order in Database");
                return new UnprocessableEntityObjectResult(JsonConvert.SerializeObject(orderMessage));
            }

            return new OkObjectResult(JsonConvert.SerializeObject(orderMessage));
        }
    }
}
