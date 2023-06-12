using System;
using System.Collections.Generic;
using System.IO;
using System.Net;
using System.Threading;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Extensions.Http;
using Microsoft.Azure.WebJobs.Extensions.OpenApi.Core.Attributes;
using Microsoft.Azure.WebJobs.Extensions.OpenApi.Core.Enums;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;
using Microsoft.Graph;
using Microsoft.Graph.Models;
using Microsoft.OpenApi.Models;
using Newtonsoft.Json;
using Shared;
using Shared.Models;

namespace Orders
{
    public class BGCreateProject
    {
        private readonly ILogger<BGCreateProject> log;
        private readonly IConfiguration config;

        public BGCreateProject(ILogger<BGCreateProject> _log, IConfiguration _config)
        {
            log = _log;
            config = _config;
        }

        [FunctionName("BGCreateProject")]
        [OpenApiOperation(operationId: "Run", tags: new[] { "name" })]
        [OpenApiSecurity("function_key", SecuritySchemeType.ApiKey, Name = "code", In = OpenApiSecurityLocationType.Query)]
        [OpenApiResponseWithBody(statusCode: HttpStatusCode.OK, contentType: "text/plain", bodyType: typeof(string), Description = "The OK response")]
        public async Task<IActionResult> Run(
            [HttpTrigger(AuthorizationLevel.Function, "get", "post", Route = null)] HttpRequest req,
            Microsoft.Azure.WebJobs.ExecutionContext context)
        {
            log.LogInformation("Create project function processed a request.");

            string Message = await new StreamReader(req.Body).ReadToEndAsync();
            Settings settings = new Settings(config, context, log);
            Graph msGraph = new Graph(settings);
            Common common = new Common(settings, msGraph);
            OrderMessage orderMessage = JsonConvert.DeserializeObject<OrderMessage>(Message);
            Order order = common.GetOrderFromCDN(orderMessage.No);

            if (order?.Customer != null && !string.IsNullOrEmpty(orderMessage.OrderParentFolderID) && !string.IsNullOrEmpty(orderMessage.OrderFolderID))
            {
                var groupDrive = await common.FindCustomerGroupAndDrive(order.Customer);

                if (groupDrive?.Success == true && groupDrive?.customer != null)
                {
                    if (!string.IsNullOrEmpty(groupDrive.customer.DriveID))
                    {
                        orderMessage.DriveID = groupDrive.customer.DriveID;

                        if (!string.IsNullOrEmpty(groupDrive.customer.GeneralFolderID))
                        {
                            var orderFolder = await common.GetOrderFolder(groupDrive.group.Id, groupDrive.groupDrive, order);

                            if (orderFolder != null)
                            {
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

                                    order.StructureCreated = true;
                                    order.Handled = copyStructure;
                                    order.Status = "Folder created and structure copied";
                                    common.UpdateOrder(order, "status");
                                }
                                else
                                {
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

        public async Task<List<DriveItem>> GetProjectTemplates(ILogger log, Graph msgraph, Site cdnSite, string CDNSiteID)
        {
            List<DriveItem> foldersToCreate = new List<DriveItem>();
            var cdnDrive = await msgraph.GetSiteDrive(CDNSiteID);

            if (cdnDrive != null)
            {
                log.LogInformation("Get templates source folder");
                DriveItem folder = await msgraph.FindItem(cdnDrive, "Dokumentstruktur Projektkanal", false);
                List<DriveItem> folderChildren = await msgraph.GetDriveFolderChildren(cdnDrive, folder, true);

                foreach (DriveItem folderChild in folderChildren)
                {
                    foldersToCreate.Add(folderChild);
                }
            }

            return foldersToCreate;
        }

        public async Task<bool> CreateProjectTabs(Settings settings, Site cdnSite, FindOrderGroupAndFolder orderGroup, DriveItem orderFolder, Order order, ILogger log, GraphServiceClient graph, Graph msgraph, string CDNSiteID, string TenantID)
        {
            bool returnValue = false;
            var channel = await msgraph.FindChannel(orderGroup.orderTeam, "Projekt " + order.No);

            if (channel == null)
            {
                try
                {
                    _ = await msgraph.CreateFolder(orderGroup.orderGroup.Id, "Projekt " + order.No);
                    channel = await msgraph.AddChannel(orderGroup.orderTeam, "Projekt " + order.No, "Projekt " + order.No, ChannelMembershipType.Standard);
                }
                catch (Exception ex)
                {
                    log.LogError("Error creating channel: " + ex.Message);
                }
            }

            //wait for channel to become available
            Thread.Sleep(60000);

            if (channel != null)
            {
                try
                {
                    var orderFolderTab = await msgraph.TabExists(orderGroup.orderTeam, channel, "Order");

                    if (!orderFolderTab)
                    {
                        log.LogInformation("Add tab with url " + orderFolder.WebUrl + " to channel " + channel.DisplayName + " in team " + orderGroup.orderTeam.DisplayName);
                        TeamsApp app = await msgraph.GetTeamApp("", "com.microsoft.teamspace.tab.web");

                        if (app != null)
                        {
                            await msgraph.AddChannelApp(orderGroup.orderTeam, app, channel, "Order", null, orderFolder.WebUrl, orderFolder.WebUrl, null);
                        }
                    }
                }
                catch (Exception ex)
                {
                    log.LogError("Error adding order channel tab: " + ex.Message);
                }

                try
                {
                    var offerFolderTab = await msgraph.TabExists(orderGroup.orderTeam, channel, "Offert");

                    if (!offerFolderTab)
                    {
                        DriveItem offerParent = await msgraph.FindItem(orderGroup.orderDrive, orderGroup.generalFolder.Id, "Offert", true);

                        log.LogInformation("Add tab with url " + offerParent.WebUrl + " to channel " + channel.DisplayName + " in team " + orderGroup.orderTeam.DisplayName);
                        TeamsApp app = await msgraph.GetTeamApp("", "com.microsoft.teamspace.tab.web");

                        if (offerParent != null && app != null)
                        {
                            await msgraph.AddChannelApp(orderGroup.orderTeam, app, channel, "Offert", null, offerParent.WebUrl, offerParent.WebUrl, null);
                        }
                    }
                }
                catch (Exception ex)
                {
                    log.LogError("Error adding offer channel tab: " + ex.Message);
                }

                try
                {
                    log?.LogInformation("Looking for planner tab");
                    string tabName = "Checklista - Projekt " + order.No;

                    //try adding checklist
                    var checklistFolderTab = await msgraph.TabExists(orderGroup.orderTeam, channel, tabName);

                    if (!checklistFolderTab)
                    {
                        log?.LogInformation("Looking for template planner");
                        //Try to find the template checklist
                        var planTemplate = await msgraph.PlanExists(settings?.CDNTeamID, "Checklista - Projektledning Template");

                        if (planTemplate != null)
                        {
                            log?.LogInformation("Looking for existing planner");
                            //found template so create the plan if it doesn't exist
                            var existingPlan = await msgraph.PlanExists(orderGroup.orderGroup.Id, tabName);

                            if (existingPlan == null)
                            {
                                log?.LogInformation("Creating new plan");
                                existingPlan = await msgraph.CreatePlanAsync(orderGroup.orderGroup.Id, tabName);

                                //copy buckets and tasks
                                var buckets = await msgraph.GetBucketsAsync(planTemplate.Id);

                                foreach (var bucket in buckets)
                                {
                                    await msgraph.CopyBucketAsync(bucket, existingPlan.Id);
                                }

                                log?.LogInformation("Copied template");

                                //create the planner tab
                                await msgraph.CreatePlannerTabInChannelAsync(graph, TenantID, orderGroup.orderTeam.Id, tabName, channel.Id, existingPlan.Id);
                                log?.LogInformation("Creating planner tab");
                            }
                            else
                            {
                                //create the planner tab
                                await msgraph.CreatePlannerTabInChannelAsync(graph, TenantID, orderGroup.orderTeam.Id, tabName, channel.Id, existingPlan.Id);
                                log?.LogInformation("Creating planner tab");
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    log?.LogError("Error adding checklist channel tab: " + ex.Message);
                }

                try
                {
                    log?.LogInformation("Copy project template files");
                    DriveItem channelFolder = await msgraph.FindItem(orderGroup.orderDrive, "Projekt " + order.No, true);

                    if (channelFolder != null)
                    {
                        List<DriveItem> projectTemplates = await GetProjectTemplates(log, msgraph, cdnSite, CDNSiteID);

                        foreach (DriveItem templateItem in projectTemplates)
                        {
                            var requestBody = new Microsoft.Graph.Drives.Item.Items.Item.Copy.CopyPostRequestBody
                            {
                                ParentReference = new ItemReference
                                {
                                    DriveId = orderGroup.orderDrive.Id,
                                    Id = channelFolder.Id,
                                },
                                Name = templateItem.Name,
                            };

                            log?.LogInformation($"Copy template item {templateItem.Name} to project folder for {order.ExternalId}.");
                            Drive siteDrive = await graph.Sites[CDNSiteID].Drive.GetAsync();
                            var result = await graph.Drives[siteDrive.Id].Items[templateItem.Id].Copy.PostAsync(requestBody);
                        }

                        try
                        {
                            var onenoteTab = await msgraph.TabExists(orderGroup.orderTeam, channel, "Mötesanteckningar");

                            if (!onenoteTab)
                            {
                                DriveItem onenotefile = await msgraph.FindItem(orderGroup.orderDrive, channelFolder.Id, "ProjectMeetingNotes", false);
                                TeamsApp app = await msgraph.GetTeamApp("", "com.microsoft.teamspace.tab.web");

                                if (onenotefile != null && app != null)
                                {
                                    log?.LogInformation("Add onenotetab with url " + onenotefile.WebUrl + " to channel " + channel.DisplayName + " in team " + orderGroup.orderTeam.DisplayName);
                                    await msgraph.AddChannelApp(orderGroup.orderTeam, app, channel, "Mötesanteckningar", null, onenotefile.WebUrl, onenotefile.WebUrl, null);
                                }
                            }
                        }
                        catch (Exception ex)
                        {
                            log?.LogError("Error Adding Onenote Tab: " + ex.Message);
                        }
                    }
                }
                catch (Exception ex)
                {
                    log?.LogError("Error copying project templates: " + ex.Message);
                }

                returnValue = true;
            }

            return returnValue;
        }

    }
}

