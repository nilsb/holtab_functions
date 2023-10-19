using System;
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
            Settings settings = new Settings(config, context, log);
            bool debug = (settings?.debugFlags?.Job?.BGUnhandledOrders).HasValue && (settings?.debugFlags?.Job?.BGUnhandledOrders).Value;
            Graph msGraph = new Graph(settings);
            Common common = new Common(settings, msGraph, debug);

            if(debug)
                log.LogInformation($"Job BGUnhandledOrders: Triggered get unhandled orders request. {DateTime.Now}");

            var orderItems = common.GetUnhandledOrderItems(debug);
            
            return new OkObjectResult(JsonConvert.SerializeObject(orderItems));
        }
    }
}

