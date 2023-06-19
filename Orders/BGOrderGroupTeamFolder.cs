using System;
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
using Microsoft.IdentityModel.Abstractions;
using Microsoft.OpenApi.Models;
using Newtonsoft.Json;
using Shared;
using Shared.Models;

namespace Orders
{
    public class BGOrderGroupTeamFolder
    {
        private readonly IConfiguration config;

        public BGOrderGroupTeamFolder(IConfiguration _config)
        {
            config = _config;
        }

        [FunctionName("BGOrderGroupTeamFolder")]
        public async Task<IActionResult> Run(
            [HttpTrigger(AuthorizationLevel.Function, new string[] { "post" }, Route = null)] HttpRequest req,
            Microsoft.Azure.WebJobs.ExecutionContext context,
            ILogger log)
        {
            log.LogInformation("Find order group, team and general folder message recieved.");
            string Message = await new StreamReader(req.Body).ReadToEndAsync();
            Settings settings = new Settings(config, context, log);
            Graph msGraph = new Graph(settings);
            Common common = new Common(settings, msGraph);
            OrderMessage orderMessage = JsonConvert.DeserializeObject<OrderMessage>(Message);
            Order order = common.GetOrderFromCDN(orderMessage.No);

            if(order?.Customer != null)
            {
                var groupDrive = await common.FindCustomerGroupAndDrive(order.Customer);

                if (groupDrive?.Success == true && groupDrive?.customer != null)
                {
                    if (!string.IsNullOrEmpty(groupDrive.customer.DriveID))
                    {
                        orderMessage.DriveID = groupDrive.customer.DriveID;

                        try
                        {
                            var orderTeam = await settings.GraphClient.Groups[order.Customer.GroupID].Team.GetAsync();

                            if (orderTeam != null)
                            {
                                order.Customer.TeamCreated = true;
                                order.Customer.TeamID = orderTeam.Id;
                            }
                            common.UpdateCustomer(order.Customer, "team info");
                        }
                        catch (Exception ex)
                        {
                            log.LogError(ex.ToString());
                            log.LogTrace($"Failed to find team for {order.Customer.Name} and order {order.ExternalId}.");
                        }

                        if (!string.IsNullOrEmpty(groupDrive.customer.GeneralFolderID))
                        {
                            orderMessage.GeneralFolderID = groupDrive.customer.GeneralFolderID;
                            groupDrive.customer.GeneralFolderCreated = true;
                            common.UpdateCustomer(groupDrive.customer, "drive and folder info.");

                            order.Status = "Incomplete";
                            order.GroupFound = true;
                            common.UpdateOrder(order, "group info");

                            string parentName = common.GetOrderParentFolderName(orderMessage.Type);
                            orderMessage.No = common.GetOrderExternalId(orderMessage.Type, orderMessage.No);
                            var orderParent = await msGraph.CreateFolder(groupDrive.group.Id, groupDrive.customer.GeneralFolderID, parentName);

                            if (orderParent?.Success == true)
                            {
                                order = common.SetFolderStatus(order, true);
                                order.Handled = false;
                                order.Status = "Parent Folder Found/Created";
                                common.UpdateOrder(order, "Parent folder info");

                                //order parent was found so find or create order folder
                                var orderFolder = await msGraph.CreateFolder(groupDrive.group.Id, orderParent.folder.Id, orderMessage.No);

                                if(orderFolder?.Success == true && orderFolder?.Existed == false)
                                {
                                    //order folder was created so update the database
                                    order = common.SetFolderStatus(order, true);

                                    order.CreatedFolder = true;
                                    order.CustomerID = groupDrive.customer.ID;
                                    order.GroupFound = true;
                                    order.GeneralFolderFound = true;
                                    order.FolderID = orderFolder.folder.Id;
                                    order.Handled = false;
                                    order.Status = "Folder Created";
                                    common.UpdateOrder(order, "folder info");

                                    orderMessage.OrderParentFolderID = orderParent.folder.Id;
                                    orderMessage.OrderFolderID = orderFolder.folder.Id;
                                    orderMessage.NeedStructureCopy = true;
                                }
                                else if(orderFolder?.Success == true && orderFolder?.Existed == true)
                                {
                                    //order folder was created so update the database
                                    order = common.SetFolderStatus(order, true);

                                    order.CreatedFolder = true;
                                    order.CustomerID = groupDrive.customer.ID;
                                    order.GroupFound = true;
                                    order.GeneralFolderFound = true;
                                    order.FolderID = orderFolder.folder.Id;
                                    order.Handled = false;
                                    order.Status = "Folder Already Existed";
                                    common.UpdateOrder(order, "folder info");

                                    orderMessage.OrderParentFolderID = orderParent.folder.Id;
                                    orderMessage.OrderFolderID = orderFolder.folder.Id;
                                    orderMessage.NeedStructureCopy = false;
                                }
                                else
                                {
                                    //error creating order folder so update the database
                                    order = common.SetFolderStatus(order, true);

                                    order.CreatedFolder = false;
                                    order.CustomerID = groupDrive.customer.ID;
                                    order.GroupFound = true;
                                    order.GeneralFolderFound = true;
                                    order.FolderID = null;
                                    order.Handled = false;
                                    order.Status = "Error Finding/Creating Folder";
                                    common.UpdateOrder(order, "folder info");

                                    return new UnprocessableEntityObjectResult($"Unable to find or create the order folder in customer {order.Customer.Name} ({order.Customer.ExternalId}) for order {orderMessage.No}");
                                }
                            }
                            else
                            {
                                //error creating order folder so update the database
                                order.CreatedFolder = false;
                                order.CustomerID = groupDrive.customer.ID;
                                order.GroupFound = true;
                                order.GeneralFolderFound = true;
                                order.FolderID = null;
                                order.OrdersFolderFound = false;
                                order.Handled = false;
                                order.Status = "Error Finding/Creating Parent Folder";
                                common.UpdateOrder(order, "folder info");

                                return new UnprocessableEntityObjectResult($"Unable to find or create the order parent folder in customer {order.Customer.Name} ({order.Customer.ExternalId}) for order {orderMessage.No}");
                            }
                        }
                        else
                        {
                            //error finding general folder so update the database
                            order.CreatedFolder = false;
                            order.CustomerID = groupDrive.customer.ID;
                            order.GroupFound = true;
                            order.GeneralFolderFound = false;
                            order.FolderID = null;
                            order.OrdersFolderFound = false;
                            order.Handled = false;
                            order.Status = "Error Finding General Folder";
                            common.UpdateOrder(order, "folder info");

                            return new NotFoundObjectResult($"Unable to find the general folder in customer {order.Customer.Name} ({order.Customer.ExternalId}) for order {orderMessage.No}");
                        }
                    }
                    else
                    {
                        //error finding general folder so update the database
                        order.CreatedFolder = false;
                        order.CustomerID = groupDrive.customer.ID;
                        order.GroupFound = false;
                        order.GeneralFolderFound = false;
                        order.FolderID = null;
                        order.OrdersFolderFound = false;
                        order.Handled = false;
                        order.Status = "Error Finding Group drive for customer";
                        common.UpdateOrder(order, "folder info");

                        return new NotFoundObjectResult($"Unable to find the drive in customer {order.Customer.Name} ({order.Customer.ExternalId}) for order {orderMessage.No}");
                    }
                }
            }

            return new OkObjectResult(JsonConvert.SerializeObject(orderMessage));
        }
    }
}

