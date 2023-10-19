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
using Newtonsoft.Json.Linq;
using Shared;
using Shared.Models;

namespace Orders
{
    public class BGCreateProject
    {
        private readonly IConfiguration config;

        public BGCreateProject(IConfiguration _config)
        {
            config = _config;
        }

        [FunctionName("BGCreateProject")]
        public async Task<IActionResult> Run(
            [HttpTrigger(AuthorizationLevel.Function, new string[] { "post" }, Route = null)] HttpRequest req,
            Microsoft.Azure.WebJobs.ExecutionContext context,
            ILogger log)
        {
            log.LogInformation("Create project function processed a request.");

            string Message = await new StreamReader(req.Body).ReadToEndAsync();
            Settings settings = new Settings(config, context, log);
            bool debug = (settings?.debugFlags?.Order?.BGCreateProject).HasValue && (settings?.debugFlags?.Order?.BGCreateProject).Value;
            Graph msGraph = new Graph(settings);
            Common common = new Common(settings, msGraph, debug);
            OrderMessage orderMessage = JsonConvert.DeserializeObject<OrderMessage>(Message);
            Order order = common.GetOrderFromCDN(orderMessage.No, debug);

            if (order?.Customer != null && !string.IsNullOrEmpty(orderMessage.OrderParentFolderID) && !string.IsNullOrEmpty(orderMessage.OrderFolderID))
            {
                if(debug)
                    log.LogInformation("Order BGCreateProject: Trying to find customer group and drive");

                var groupDrive = await common.FindCustomerGroupAndDrive(order.Customer, debug);

                if (groupDrive?.Success == true && groupDrive?.customer != null)
                {
                    if (!string.IsNullOrEmpty(groupDrive.customer.DriveID))
                    {
                        orderMessage.DriveID = groupDrive.customer.DriveID;

                        if (!string.IsNullOrEmpty(groupDrive.customer.GeneralFolderID))
                        {
                            if(debug)
                                log.LogInformation("Order BGCreateProject: Found customer group and drive, getting order folder");

                            var orderFolder = await common.GetOrderFolder(groupDrive.groupId, groupDrive.groupDriveId, order, debug);

                            if (orderFolder != null)
                            {
                                if(debug)
                                    log.LogInformation("Order BGCreateProject: Found order folder, fetching team for customer");

                                string groupTeamId = await msGraph.GetTeamFromGroup(groupDrive.groupId, debug);

                                if(!string.IsNullOrEmpty(groupTeamId))
                                {
                                    if(debug)
                                        log.LogInformation("Order BGCreateProject: Found team for customer, adding tabs");

                                    _ = await CreateProjectTabs(settings, groupDrive, groupTeamId, orderFolder, order, log, msGraph, debug);
                                }
                            }
                            else
                            {
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

        public async Task<List<DriveItem>> GetProjectTemplates(ILogger log, Graph msgraph, string CDNSiteID, bool debug)
        {
            List<DriveItem> foldersToCreate = new List<DriveItem>();
            string cdnDriveId = await msgraph.GetSiteDrive(CDNSiteID, debug);

            if (!string.IsNullOrEmpty(cdnDriveId))
            {
                if(debug)
                    log.LogInformation("Order BGCreateProject: Get templates source folder");

                DriveItem folder = await msgraph.FindItem(cdnDriveId, "Dokumentstruktur Projektkanal", false, debug);
                List<DriveItem> folderChildren = await msgraph.GetDriveFolderChildren(cdnDriveId, folder.Id, true, debug);

                foreach (DriveItem folderChild in folderChildren)
                {
                    foldersToCreate.Add(folderChild);
                }
            }

            return foldersToCreate;
        }

        public async Task<bool> CreateProjectTabs(Settings settings, FindCustomerGroupResult customerGroup, string orderTeamId, DriveItem orderFolder, Order order, ILogger log, Graph msgraph, bool debug)
        {
            bool returnValue = false;
            var channel = await msgraph.FindChannel(orderTeamId, "Projekt " + order.ExternalId, debug);

            if (channel == null)
            {
                try
                {
                    _ = await msgraph.CreateFolder(customerGroup.groupId, "Projekt " + order.ExternalId, debug);
                    channel = await msgraph.AddChannel(orderTeamId, "Projekt " + order.ExternalId, "Projekt " + order.ExternalId, ChannelMembershipType.Standard, debug);
                }
                catch (Exception ex)
                {
                    log.LogError("Order BGCreateProject: Error creating channel with error " + ex.Message);
                }
            }

            //wait for channel to become available
            Thread.Sleep(60000);

            if (channel != null)
            {
                try
                {
                    var orderFolderTab = await msgraph.TabExists(orderTeamId, channel, "Order", debug);

                    if (!orderFolderTab)
                    {
                        if(debug)
                            log.LogInformation("Order BGCreateProject: Add tab with url " + orderFolder.WebUrl + " to channel " + channel + " in team " + orderTeamId);

                        await msgraph.AddChannelWebApp(orderTeamId, channel, "Order", orderFolder.WebUrl, orderFolder.WebUrl, debug);
                    }
                }
                catch (Exception ex)
                {
                    log.LogError("Order BGCreateProject: Error adding order channel tab with error " + ex.Message);
                }

                try
                {
                    var offerFolderTab = await msgraph.TabExists(orderTeamId, channel, "Offert", debug);

                    if (!offerFolderTab)
                    {
                        DriveItem offerParent = await msgraph.FindItem(customerGroup.groupDriveId, customerGroup.customer.GeneralFolderID, "Offert", true, debug);

                        if(debug)
                            log.LogInformation("Order BGCreateProject: Add tab with url " + offerParent.WebUrl + " to channel " + channel + " in team " + orderTeamId);

                        await msgraph.AddChannelWebApp(orderTeamId, channel, "Offert", offerParent.WebUrl, offerParent.WebUrl, debug);
                    }
                }
                catch (Exception ex)
                {
                    log.LogError("Order BGCreateProject: Error adding offer channel tab: " + ex.Message);
                }

                try
                {
                    if(debug)
                        log.LogInformation("Order BGCreateProject: Looking for planner tab");

                    string tabName = "Checklista - Projekt " + order.ExternalId;

                    //try adding checklist
                    var checklistFolderTab = await msgraph.TabExists(orderTeamId, channel, tabName, debug);

                    if (!checklistFolderTab)
                    {
                        if(debug)
                            log.LogInformation("Order BGCreateProject: Looking for template planner");

                        //Try to find the template checklist
                        var planTemplate = await msgraph.PlanExists(settings?.CDNTeamID, "Checklista - Projektledning Template", debug);

                        if (planTemplate != null)
                        {
                            if(debug)
                                log.LogInformation("Order BGCreateProject: Looking for existing planner");

                            //found template so create the plan if it doesn't exist
                            var existingPlan = await msgraph.PlanExists(customerGroup.groupId, tabName, debug);

                            if (existingPlan == null)
                            {
                                if(debug)
                                    log.LogInformation("Order BGCreateProject: Creating new plan");

                                existingPlan = await msgraph.CreatePlanAsync(customerGroup.groupId, tabName, debug);

                                //copy buckets and tasks
                                var buckets = await msgraph.GetBucketsAsync(planTemplate.Id, debug);

                                if(debug)
                                    log.LogInformation("Order BGCreateProject: Creating buckets");

                                foreach (var bucket in buckets)
                                {
                                    await msgraph.CopyBucketAsync(bucket, existingPlan.Id, debug);
                                }

                                if(debug)
                                    log.LogInformation("Order BGCreateProject: Copied planner template");

                                //create the planner tab
                                //log.LogInformation("Creating planner tab");
                                //await msgraph.CreatePlannerTabInChannelAsync(orderTeam.Id, tabName, channel.Id, existingPlan.Id);
                            }
                            else
                            {
                                //create the planner tab
                                //log.LogInformation("Creating planner tab");
                                //await msgraph.CreatePlannerTabInChannelAsync(orderTeam.Id, tabName, channel.Id, existingPlan.Id);
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    log.LogError("Order BGCreateProject: Error adding checklist channel tab with error " + ex.Message);
                }

                try
                {
                    if(debug)
                        log.LogInformation("Order BGCreateProject: Copy project template files");

                    DriveItem channelFolder = await msgraph.FindItem(customerGroup.groupDriveId, "Projekt " + order.ExternalId, true, debug);

                    if (channelFolder != null)
                    {
                        List<DriveItem> projectTemplates = await GetProjectTemplates(log, msgraph, settings.cdnSiteId, debug);

                        foreach (DriveItem templateItem in projectTemplates)
                        {
                            var requestBody = new Microsoft.Graph.Drives.Item.Items.Item.Copy.CopyPostRequestBody
                            {
                                ParentReference = new ItemReference
                                {
                                    DriveId = customerGroup.groupDriveId,
                                    Id = channelFolder.Id,
                                },
                                Name = templateItem.Name,
                            };

                            if(debug)
                                log.LogInformation($"Order BGCreateProject: Copy template item {templateItem.Name} to project folder for {order.ExternalId}.");

                            string siteDriveId = await msgraph.GetSiteDrive(settings.cdnSiteId, debug);
                            var result = await settings.GraphClient.Drives[siteDriveId].Items[templateItem.Id].Copy.PostAsync(requestBody);
                        }

                        try
                        {
                            var notesTab = await msgraph.GetTab(orderTeamId, channel, "Notes", debug);

                            if (!string.IsNullOrEmpty(notesTab))
                            {
                                await msgraph.RemoveTab(orderTeamId, channel, notesTab, debug);
                            }
                        }
                        catch (Exception ex)
                        {
                            log.LogError("Order BGCreateProject: Error removing default notes tab: " + ex.Message);
                        }

                        try
                        {
                            var onenoteTab = await msgraph.TabExists(orderTeamId, channel, "Anteckningar", debug);

                            if (!onenoteTab)
                            {
                                DriveItem onenotefile = await msgraph.FindItem(customerGroup.groupDriveId, channelFolder.Id, "ProjectMeetingNotes", false, debug);

                                if (onenotefile != null)
                                {
                                    if(debug)
                                        log.LogInformation("Order BGCreateProject: Add onenotetab with url " + onenotefile.WebUrl + " to channel " + channel + " in team " + orderTeamId, debug);

                                    await msgraph.AddChannelWebApp(orderTeamId, channel, "Anteckningar", onenotefile.WebUrl, onenotefile.WebUrl, debug);
                                }
                            }
                        }
                        catch (Exception ex)
                        {
                            log.LogError("Order BGCreateProject: Error Adding Onenote Tab with error " + ex.Message);
                        }
                    }
                }
                catch (Exception ex)
                {
                    log.LogError("Order BGCreateProject: Error copying project templates with error " + ex.Message);
                }

                returnValue = true;
            }

            return returnValue;
        }

    }
}