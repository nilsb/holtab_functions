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
using Shared.Models;
using Shared;

namespace Orders
{
    public class BGOrderInfo
    {
        private readonly IConfiguration config;

        public BGOrderInfo(IConfiguration config)
        {
            this.config = config;
        }

        [FunctionName("BGOrderInfo")]
        public async Task<IActionResult> Run(
            [HttpTrigger(AuthorizationLevel.Function, "get", "post", Route = null)] HttpRequest req,
            Microsoft.Azure.WebJobs.ExecutionContext context,
            ILogger log)
        {
            string Message = await new StreamReader(req.Body).ReadToEndAsync();
            log.LogInformation($"Order Information trigger function processed message: {Message}");
            Settings settings = new Settings(config, context, log);
            Graph msGraph = new Graph(settings);
            Common common = new Common(settings, msGraph);
            OrderMessage order = JsonConvert.DeserializeObject<OrderMessage>(Message);


            //Find and update or create the order database post
            Order cdnItem = common.GetOrderFromCDN(order.No);

            if(cdnItem != null)
            {
                if (string.IsNullOrEmpty(cdnItem.ExternalId))
                    cdnItem.ExternalId = order.No; //backwards compatibility

                cdnItem.AdditionalInfo = order.AdditionalInfo;
                cdnItem.Seller = order.Seller;
                cdnItem.ProjectManager = order.ProjectManager;
                cdnItem.Customer = null;
                cdnItem.CustomerNo = order.CustomerNo;
                cdnItem.CustomerType = order.CustomerType;
                Order dbOrder = common.UpdateOrCreateDbOrder(cdnItem);
            }
            else
            {
                Order newOrder = new Order() { 
                    No = order.No,
                    ExternalId = order.No,
                    AdditionalInfo = order.AdditionalInfo,
                    CustomerNo = order.CustomerNo,
                    CustomerType = order.CustomerType,
                    Seller = order.Seller,
                    ProjectManager = order.ProjectManager
                };

                Order dbOrder = common.UpdateOrCreateDbOrder(newOrder);
            }

            return new OkObjectResult("");
        }
    }
}
