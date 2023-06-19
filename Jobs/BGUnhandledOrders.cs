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
using Shared;
using Shared.Models;

namespace Jobs
{
    public class BGUnhandledOrders
    {
        private readonly IConfiguration config;

        public BGUnhandledOrders(IConfiguration _config)
        {
            config = _config;
        }

        [FunctionName("BGUnhandledOrders")]
        public IActionResult Run(
            [HttpTrigger(AuthorizationLevel.Function, new string[] { "post" }, Route = null)] HttpRequest req,
            Microsoft.Azure.WebJobs.ExecutionContext context,
            ILogger log)
        {
            log.LogInformation("Triggered get unhandled orders request.");

            Settings settings = new Settings(config, context, log);
            Graph msGraph = new Graph(settings);
            Common common = new Common(settings, msGraph);

            var orderItems = common.GetUnhandledOrderItems();
            
            return new OkObjectResult(JsonConvert.SerializeObject(orderItems));
        }
    }
}

