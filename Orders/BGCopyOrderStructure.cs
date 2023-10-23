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
using Newtonsoft.Json.Linq;
using Shared;
using Shared.Models;

namespace Orders
{
    public class BGCopyOrderStructure
    {
        private readonly IConfiguration config;

        public BGCopyOrderStructure(IConfiguration _config)
        {
            config = _config;
        }

        [FunctionName("BGCopyOrderStructure")]
        public async Task<IActionResult> Run(
            [HttpTrigger(AuthorizationLevel.Function, new string[] { "post" }, Route = null)] HttpRequest req,
            Microsoft.Azure.WebJobs.ExecutionContext context,
            ILogger log)
        {
            string Message = await new StreamReader(req.Body).ReadToEndAsync();

            Settings settings = new Settings(config, context, log);
            bool debug = (settings?.debugFlags?.Order?.BGCopyOrderStructure).HasValue && (settings?.debugFlags?.Order?.BGCopyOrderStructure).Value;
            Graph msGraph = new Graph(settings);
            Common common = new Common(settings, msGraph, debug);
            OrderMessage orderMessage = JsonConvert.DeserializeObject<OrderMessage>(Message);
            Order order = common.GetOrderFromCDN(orderMessage.No, debug);

            if(debug)
                log.LogInformation("Order BGCopyOrderStructure: Copy order structure message recieved.");

            if (order?.Customer != null && !string.IsNullOrEmpty(orderMessage.OrderParentFolderID) && !string.IsNullOrEmpty(orderMessage.OrderFolderID))
            {
                if(debug)
                    log.LogInformation("Order BGCopyOrderStructure: Found customer.");

                var groupDrive = await common.FindCustomerGroupAndDrive(order.Customer, debug);

                if (groupDrive?.Success == true && groupDrive?.customer != null)
                {
                    if (!string.IsNullOrEmpty(groupDrive.customer.DriveID))
                    {
                        if(debug)
                            log.LogInformation("Order BGCopyOrderStructure: Group drive found.");

                        orderMessage.DriveID = groupDrive.customer.DriveID;

                        if (!string.IsNullOrEmpty(groupDrive.customer.GeneralFolderID))
                        {
                            var orderFolder = await common.GetOrderFolder(groupDrive.groupId, groupDrive.groupDriveId, order, debug);

                            if(orderFolder != null)
                            {
                                if(debug)
                                    log.LogInformation("Order BGCopyOrderStructure: Order folder found.");

                                if (orderMessage.NeedStructureCopy == true)
                                {
                                    bool copyStructure = false;
                                    List<DriveItem> templateFolders = await common.GetOrderTemplateFolders(order, debug);

                                    foreach (DriveItem templateFolder in templateFolders)
                                    {
                                        CreateFolderResult result = await msGraph.CopyFolder(groupDrive.groupId, orderFolder.Id, templateFolder, true, false, debug);

                                        if (result.Success)
                                        {
                                            copyStructure &= true;
                                            if (debug)
                                                log.LogInformation("Order BGCopyOrderStructure: Folder created and structure copied.");

                                            order.StructureCreated = true;
                                            order.Handled = copyStructure;
                                            order.Status = "Folder created and structure copied";
                                            common.UpdateOrder(order, "status", debug);
                                        }
                                        else
                                        {
                                            copyStructure &= false;
                                            if (debug)
                                                log.LogInformation("Order BGCopyOrderStructure: Failed to copy structure.");
                                        }
                                    }
                                }
                                else
                                {
                                    if(debug)
                                        log.LogInformation("Order BGCopyOrderStructure: Order folder already existed.");

                                    order.Status = "Order folder already existed";
                                    order.Handled = true;
                                    order.FolderID = orderFolder.Id;
                                    order.DriveFound = true;
                                    order.DriveID = groupDrive.groupDriveId;
                                    order.CreatedFolder = true;
                                    order.GroupFound = true;
                                    order.GeneralFolderFound = true;
                                    common.UpdateOrder(order, "status", debug);
                                }
                            }
                            else
                            {
                                if(debug)
                                    log.LogInformation("Order BGCopyOrderStructure: Order folder not found.");

                                order.Handled = false;
                                order.Status = "Order folder not found";
                                common.UpdateOrder(order, "status", debug);

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

