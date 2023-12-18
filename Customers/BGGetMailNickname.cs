using System;
using System.IO;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Extensions.Http;
using Microsoft.AspNetCore.Http;
using Microsoft.Extensions.Logging;
using Newtonsoft.Json;
using Shared.Models;
using Shared;
using Microsoft.Extensions.Configuration;

namespace Customers
{
    public class BGGetMailNickname
    {
        private readonly IConfiguration config;

        public BGGetMailNickname(IConfiguration _config)
        {
            config = _config;
        }

        [FunctionName("BGGetMailNickname")]
        public async Task<IActionResult> Run(
            [HttpTrigger(AuthorizationLevel.Function, "get", "post", Route = null)] HttpRequest req,
            Microsoft.Azure.WebJobs.ExecutionContext context,
            ILogger log)
        {
            Settings settings = new Settings(config, context, log);
            bool debug = true;
            Graph msGraph = new Graph(settings);
            Common common = new Common(settings, msGraph, debug);

            string customername = req.Query["customername"];
            string customerno = req.Query["customerno"];
            string customertype = req.Query["customertype"];

            string requestBody = await new StreamReader(req.Body).ReadToEndAsync();
            dynamic data = JsonConvert.DeserializeObject(requestBody);
            customername = customername ?? data?.customername;
            customerno = customerno ?? data?.customerno;
            customertype = customertype ?? data?.customertype;

            string responseMessage = common.GetMailNickname(customername, customerno, customertype, true);

            return new OkObjectResult(responseMessage);
        }
    }
}
