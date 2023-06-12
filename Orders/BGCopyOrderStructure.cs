using System.Collections.Generic;
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
using Microsoft.Graph.Models;
using Microsoft.OpenApi.Models;
using Newtonsoft.Json;
using Shared;
using Shared.Models;

namespace Orders
{
    public class BGCopyOrderStructure
    {
        private readonly ILogger<BGCopyOrderStructure> log;
        private readonly IConfiguration config;

        public BGCopyOrderStructure(ILogger<BGCopyOrderStructure> _log, IConfiguration _config)
        {
            log = _log;
            config = _config;
        }

        [FunctionName("BGCopyOrderStructure")]
        [OpenApiOperation(operationId: "Run", tags: new[] { "name" })]
        [OpenApiSecurity("function_key", SecuritySchemeType.ApiKey, Name = "code", In = OpenApiSecurityLocationType.Query)]
        [OpenApiResponseWithBody(statusCode: HttpStatusCode.OK, contentType: "text/plain", bodyType: typeof(string), Description = "The OK response")]
        [OpenApiResponseWithBody(statusCode: HttpStatusCode.UnprocessableEntity, contentType: "text/plain", bodyType: typeof(string), Description = "Was unable to copy folder structure")]
        public async Task<IActionResult> Run(
            [HttpTrigger(AuthorizationLevel.Function, "get", "post", Route = null)] HttpRequest req,
            Microsoft.Azure.WebJobs.ExecutionContext context)
        {
            log.LogInformation("Copy order structure message recieved.");
            string Message = await new StreamReader(req.Body).ReadToEndAsync();
            Settings settings = new Settings(config, context, log);
            Graph msGraph = new Graph(settings);
            Common common = new Common(settings, msGraph);
            OrderMessage orderMessage = JsonConvert.DeserializeObject<OrderMessage>(Message);
            Order order = common.GetOrderFromCDN(orderMessage.No);

            if (order?.Customer != null && !string.IsNullOrEmpty(orderMessage.OrderParentFolderID) && !string.IsNullOrEmpty(orderMessage.OrderFolderID))
            {
                log.LogInformation("Found customer.");
                var groupDrive = await common.FindCustomerGroupAndDrive(order.Customer);

                if (groupDrive?.Success == true && groupDrive?.customer != null)
                {
                    if (!string.IsNullOrEmpty(groupDrive.customer.DriveID))
                    {
                        log.LogInformation("Group drive found.");
                        orderMessage.DriveID = groupDrive.customer.DriveID;

                        if (!string.IsNullOrEmpty(groupDrive.customer.GeneralFolderID))
                        {
                            var orderFolder = await common.GetOrderFolder(groupDrive.group.Id, groupDrive.groupDrive, order);

                            if(orderFolder != null)
                            {
                                log.LogInformation("Order folder found.");
                                if (orderMessage.NeedStructureCopy == true)
                                {
                                    bool copyStructure = false;
                                    List<DriveItem> templateFolders = await common.GetOrderTemplateFolders(order);

                                    foreach (DriveItem templateFolder in templateFolders)
                                    {
                                        CreateFolderResult result = await msGraph.CopyFolder(groupDrive.group.Id, orderFolder.Id, templateFolder, true, false);

                                        if (result.Success)
                                        {
                                            copyStructure &= true;
                                        }
                                        else
                                        {
                                            copyStructure &= false;
                                        }
                                    }

                                    log.LogInformation("Folder created and structure copied.");
                                    order.StructureCreated = true;
                                    order.Handled = copyStructure;
                                    order.Status = "Folder created and structure copied";
                                    common.UpdateOrder(order, "status");
                                }
                                else
                                {
                                    log.LogInformation("Order folder already existed.");
                                    order.Status = "Order folder already existed";
                                    order.Handled = true;
                                    order.FolderID = orderFolder.Id;
                                    order.DriveFound = true;
                                    order.DriveID = groupDrive.groupDrive.Id;
                                    order.CreatedFolder = true;
                                    order.GroupFound = true;
                                    order.GeneralFolderFound = true;
                                    common.UpdateOrder(order, "status");
                                }
                            }
                            else
                            {
                                log.LogInformation("Order folder not found.");
                                order.Handled = false;
                                order.Status = "Order folder not found";
                                common.UpdateOrder(order, "status");

                                return new UnprocessableEntityObjectResult($"Unable to find order folder for order {order.ExternalId} in customer {order.Customer.Name} ({order.Customer.ExternalId})");
                            }
                        }
                    }
                }
            }

            return new OkObjectResult(JsonConvert.SerializeObject(orderMessage));
        }
    }
}

