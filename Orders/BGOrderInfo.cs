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
using Newtonsoft.Json.Linq;

namespace Orders
{
    public class BGOrderInfo
    {
        private readonly IConfiguration config;

        public BGOrderInfo(IConfiguration _config)
        {
            config = _config;
        }

        [FunctionName("BGOrderInfo")]
        public async Task<IActionResult> Run(
            [HttpTrigger(AuthorizationLevel.Function, new string[] { "post" }, Route = null)] HttpRequest req,
            Microsoft.Azure.WebJobs.ExecutionContext context,
            ILogger log)
        {
            string Message = await new StreamReader(req.Body).ReadToEndAsync();

            log.LogInformation($"Order Information trigger function processed message: {Message}");
            Settings settings = new Settings(config, context, log);
            bool debug = (settings?.debugFlags?.Order?.BGOrderInfo).HasValue && (settings?.debugFlags?.Order?.BGOrderInfo).Value;
            Graph msGraph = new Graph(settings);
            Common common = new Common(settings, msGraph, debug);

            OrderMessage orderMessage = JsonConvert.DeserializeObject<OrderMessage>(Message);

            if(string.IsNullOrEmpty(orderMessage.No) && !string.IsNullOrEmpty(orderMessage.ExternalId))
            {
                if(debug)
                    log.LogInformation("Order BGOrderInfo: Message did not contain ExternalId so assigning it from No");

                orderMessage.No = orderMessage.ExternalId;
            }

            if (string.IsNullOrEmpty(orderMessage.ExternalId) && !string.IsNullOrEmpty(orderMessage.No))
            {
                if(debug)
                    log.LogInformation("Order BGOrderInfo: Message did not contain No so assigning it from ExternalId");

                orderMessage.ExternalId = orderMessage.No;
            }

            var newOrder = default(Order);

            //Find and update or create the order database post
            Order cdnItem = common.GetOrderFromCDN(orderMessage.ExternalId, debug);

            if(cdnItem != null)
            {
                if(debug)
                    log.LogInformation("Order BGOrderInfo: Found order in Database");

                if (string.IsNullOrEmpty(cdnItem.ExternalId))
                    cdnItem.ExternalId = orderMessage.ExternalId; //backwards compatibility

                cdnItem.AdditionalInfo = orderMessage.AdditionalInfo;
                cdnItem.Seller = orderMessage.Seller;
                cdnItem.ProjectManager = orderMessage.ProjectManager;
                cdnItem.Customer = null;
                cdnItem.CustomerNo = orderMessage.CustomerNo;
                cdnItem.CustomerType = orderMessage.CustomerType;
                cdnItem.Type = orderMessage.Type;
                newOrder = common.UpdateOrCreateDbOrder(cdnItem, debug);
                
                if(debug)
                    log.LogInformation("Order BGOrderInfo: Updated order in Database");
            }
            else
            {
                newOrder = new Order() { 
                    No = orderMessage.ExternalId,
                    ExternalId = orderMessage.ExternalId,
                    AdditionalInfo = orderMessage.AdditionalInfo,
                    CustomerNo = orderMessage.CustomerNo,
                    CustomerType = orderMessage.CustomerType,
                    Seller = orderMessage.Seller,
                    Type = orderMessage.Type,
                    ProjectManager = orderMessage.ProjectManager
                };

                newOrder = common.UpdateOrCreateDbOrder(newOrder, debug);
                
                if(debug)
                    log.LogInformation("Order BGOrderInfo: Created order in Database");
            }

            if (newOrder != null)
            {
                if(debug)
                    log.LogInformation("Order BGOrderInfo: Order was created or updated in Database");

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
                if(debug)
                    log.LogInformation("Order BGOrderInfo: Could not create or update order in Database");

                return new UnprocessableEntityObjectResult(JsonConvert.SerializeObject(orderMessage));
            }

            return new OkObjectResult(JsonConvert.SerializeObject(orderMessage));
        }
    }
}
