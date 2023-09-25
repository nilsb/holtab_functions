using System;
using Shared;
using Microsoft.Azure.WebJobs;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;
using Newtonsoft.Json;
using Shared.Models;
using System.Net.Http;
using System.Net.Http.Json;

namespace Jobs
{
    public class PostProcessOrders
    {
        private readonly HttpClient _http;
        private readonly IConfiguration config;

        public PostProcessOrders(IConfiguration _config, IHttpClientFactory httpClientFactory)
        {
            config = _config;
            _http = httpClientFactory.CreateClient();
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
                order.QueueCount = order.QueueCount + 1;
                common.UpdateOrder(order, "queue count");

                if (order.QueueCount < 3) {
                    log.LogTrace("Putting message on order queue: " + JsonConvert.SerializeObject(order));
                    var response = _http.PostAsJsonAsync<OrderMessage>("https://prod-43.westeurope.logic.azure.com:443/workflows/f048e29daba148ea989a6aac88aa636b/triggers/manual/paths/invoke?api-version=2016-10-01&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=vDwjEvZsMdJ6QurTJMKVu6NEgK7MorJm3mzSck7NkNA", 
                        new OrderMessage { AdditionalInfo = order.AdditionalInfo, CustomerNo = order.CustomerNo, CustomerType = order.CustomerType, No = order.ExternalId, ProjectManager = order.ProjectManager, Seller = order.Seller, Type = order.Type });
                    //outputQueueItem.Add(JsonConvert.SerializeObject(order));
                }
            }
        }
    }
}
