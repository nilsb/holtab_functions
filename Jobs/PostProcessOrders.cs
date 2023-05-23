using System;
using Shared;
using Microsoft.Azure.WebJobs;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;
using Newtonsoft.Json;
using Shared.Models;

namespace Jobs
{
    public class PostProcessOrders
    {
        private readonly ILogger<PostProcessOrders> log;
        private readonly IConfiguration config;

        public PostProcessOrders(ILogger<PostProcessOrders> _log, IConfiguration _config)
        {
            log = _log;
            config = _config;
        }

        [FunctionName("PostProcessOrders")]
        public void Run([TimerTrigger("0 */5 * * * *")]TimerInfo myTimer, [Queue("createorder"), StorageAccount("AzureWebJobsStorage")] ICollector<string> outputQueueItem, Microsoft.Azure.WebJobs.ExecutionContext context, ILogger log)
        {
            log.LogInformation($"C# Timer trigger function executed at: {DateTime.Now}");

            Settings settings = new Settings(config, context, log);
            Graph msGraph = new Graph(settings);
            Common common = new Common(settings, msGraph);

            var orderItems = common.GetUnhandledOrderItems();

            foreach(var order in orderItems)
            {
                log.LogTrace("Putting message on order queue: " + JsonConvert.SerializeObject(order));
                outputQueueItem.Add(JsonConvert.SerializeObject(order));
            }
        }
    }
}
