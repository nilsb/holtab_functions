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

namespace Jobs
{
    public class GetOrderNo
    {
        private readonly IConfiguration config;

        public GetOrderNo(IConfiguration _config)
        {
            config = _config;
        }

        [FunctionName("GetOrderNo")]
        public IActionResult Run(
            [HttpTrigger(AuthorizationLevel.Function, "get", "post", Route = null)] HttpRequest req,
            Microsoft.Azure.WebJobs.ExecutionContext context,
            ILogger log)
        {
            log.LogInformation("Trigger function processed a request to find order number.");

            Settings settings = new Settings(config, context, log);
            bool debug = (settings?.debugFlags?.Job?.PostProcessEmails).HasValue && (settings?.debugFlags?.Job?.PostProcessEmails).Value;
            Graph msGraph = new Graph(settings);
            Common common = new Common(settings, msGraph, debug);

            string input = req.Query["input"];
            string orderno = common.FindOrderNoInString(input);

            if(!string.IsNullOrEmpty(orderno)) {
                return new OkObjectResult(orderno);
            }

            string customerno = common.FindCustomerNoInString(input);

            if (!string.IsNullOrEmpty(customerno))
            {
                return new OkObjectResult(customerno);
            }

            return new UnprocessableEntityObjectResult("");
        }
    }
}
