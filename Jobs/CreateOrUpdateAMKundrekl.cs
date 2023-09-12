using System.IO;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Extensions.Http;
using Microsoft.AspNetCore.Http;
using Microsoft.Extensions.Logging;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using Microsoft.Graph.Models;

namespace Jobs
{
    public static class CreateOrUpdateAMKundrekl
    {
        [FunctionName("CreateOrUpdateAMKundrekl")]
        public static async Task<IActionResult> Run(
            [HttpTrigger(AuthorizationLevel.Function, new string[] { "get", "post" }, Route = null)] HttpRequest req,
            ILogger log)
        {
            log.LogInformation("C# HTTP trigger function processed a request.");

            string name = req.Query["id"];

            string requestBody = await new StreamReader(req.Body).ReadToEndAsync();
            dynamic data = JsonConvert.DeserializeObject<dynamic>(requestBody);
            string responseMessage = "";

            if (data != null )
            {
                data.id = data.regno;
                responseMessage = JsonConvert.SerializeObject(data);
            }

            return new OkObjectResult(responseMessage);
        }
    }
}
