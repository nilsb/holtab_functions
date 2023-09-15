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
            Graph msGraph = new Graph(settings);
            Common common = new Common(settings, msGraph);
            OrderMessage orderMessage = JsonConvert.DeserializeObject<OrderMessage>(Message);
            Order order = common.GetOrderFromCDN(orderMessage.No);

            if (order?.Customer != null && !string.IsNullOrEmpty(orderMessage.OrderParentFolderID) && !string.IsNullOrEmpty(orderMessage.OrderFolderID))
            {
                log.LogInformation("Trying to find customer group and drive");
                var groupDrive = await common.FindCustomerGroupAndDrive(order.Customer);

                if (groupDrive?.Success == true && groupDrive?.customer != null)
                {
                    if (!string.IsNullOrEmpty(groupDrive.customer.DriveID))
                    {
                        orderMessage.DriveID = groupDrive.customer.DriveID;

                        if (!string.IsNullOrEmpty(groupDrive.customer.GeneralFolderID))
                        {
                            log.LogInformation("Found customer group and drive, getting order folder");
                            var orderFolder = await common.GetOrderFolder(groupDrive.group.Id, groupDrive.groupDrive, order);

                            if (orderFolder != null)
                            {
                                log.LogInformation("Found order folder, fetching team for customer");
                                Team groupTeam = await msGraph.GetTeamFromGroup(groupDrive.group.Id);

                                if(groupTeam != null)
                                {
                                    log.LogInformation("Found team for customer, adding tabs");
                                    _ = await CreateProjectTabs(settings, groupDrive, groupTeam, orderFolder, order, log, msGraph);
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

        public async Task<List<DriveItem>> GetProjectTemplates(ILogger log, Graph msgraph, string CDNSiteID)
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

        public async Task<bool> CreateProjectTabs(Settings settings, FindCustomerGroupResult customerGroup, Team orderTeam, DriveItem orderFolder, Order order, ILogger log, Graph msgraph)
        {
            bool returnValue = false;
            var channel = await msgraph.FindChannel(orderTeam, "Projekt " + order.ExternalId);

            if (channel == null)
            {
                try
                {
                    _ = await msgraph.CreateFolder(customerGroup.group.Id, "Projekt " + order.ExternalId);
                    channel = await msgraph.AddChannel(orderTeam, "Projekt " + order.ExternalId, "Projekt " + order.ExternalId, ChannelMembershipType.Standard);
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
                    var orderFolderTab = await msgraph.TabExists(orderTeam, channel, "Order");

                    if (!orderFolderTab)
                    {
                        log.LogInformation("Add tab with url " + orderFolder.WebUrl + " to channel " + channel.DisplayName + " in team " + orderTeam.DisplayName);
                        await msgraph.AddChannelWebApp(orderTeam, channel, "Order", orderFolder.WebUrl, orderFolder.WebUrl);
                    }
                }
                catch (Exception ex)
                {
                    log.LogError("Error adding order channel tab: " + ex.Message);
                }

                try
                {
                    var offerFolderTab = await msgraph.TabExists(orderTeam, channel, "Offert");

                    if (!offerFolderTab)
                    {
                        DriveItem offerParent = await msgraph.FindItem(customerGroup.groupDrive, customerGroup.customer.GeneralFolderID, "Offert", true);

                        log.LogInformation("Add tab with url " + offerParent.WebUrl + " to channel " + channel.DisplayName + " in team " + orderTeam.DisplayName);
                        await msgraph.AddChannelWebApp(orderTeam, channel, "Offert", offerParent.WebUrl, offerParent.WebUrl);
                    }
                }
                catch (Exception ex)
                {
                    log.LogError("Error adding offer channel tab: " + ex.Message);
                }

                try
                {
                    log.LogInformation("Looking for planner tab");
                    string tabName = "Checklista - Projekt " + order.ExternalId;

                    //try adding checklist
                    var checklistFolderTab = await msgraph.TabExists(orderTeam, channel, tabName);

                    if (!checklistFolderTab)
                    {
                        log.LogInformation("Looking for template planner");
                        //Try to find the template checklist
                        var planTemplate = await msgraph.PlanExists(settings?.CDNTeamID, "Checklista - Projektledning Template");

                        if (planTemplate != null)
                        {
                            log.LogInformation("Looking for existing planner");
                            //found template so create the plan if it doesn't exist
                            var existingPlan = await msgraph.PlanExists(customerGroup.group.Id, tabName);

                            if (existingPlan == null)
                            {
                                log.LogInformation("Creating new plan");
                                existingPlan = await msgraph.CreatePlanAsync(customerGroup.group.Id, tabName);

                                //copy buckets and tasks
                                var buckets = await msgraph.GetBucketsAsync(planTemplate.Id);

                                foreach (var bucket in buckets)
                                {
                                    await msgraph.CopyBucketAsync(bucket, existingPlan.Id);
                                }

                                log.LogInformation("Copied template");

                                //create the planner tab
                                await msgraph.CreatePlannerTabInChannelAsync(orderTeam.Id, tabName, channel.Id, existingPlan.Id);
                                log.LogInformation("Creating planner tab");
                            }
                            else
                            {
                                //create the planner tab
                                await msgraph.CreatePlannerTabInChannelAsync(orderTeam.Id, tabName, channel.Id, existingPlan.Id);
                                log.LogInformation("Creating planner tab");
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    log.LogError("Error adding checklist channel tab: " + ex.Message);
                }

                try
                {
                    log.LogInformation("Copy project template files");
                    DriveItem channelFolder = await msgraph.FindItem(customerGroup.groupDrive, "Projekt " + order.ExternalId, true);

                    if (channelFolder != null)
                    {
                        List<DriveItem> projectTemplates = await GetProjectTemplates(log, msgraph, settings.cdnSiteId);

                        foreach (DriveItem templateItem in projectTemplates)
                        {
                            var requestBody = new Microsoft.Graph.Drives.Item.Items.Item.Copy.CopyPostRequestBody
                            {
                                ParentReference = new ItemReference
                                {
                                    DriveId = customerGroup.groupDrive.Id,
                                    Id = channelFolder.Id,
                                },
                                Name = templateItem.Name,
                            };

                            log.LogInformation($"Copy template item {templateItem.Name} to project folder for {order.ExternalId}.");
                            Drive siteDrive = await msgraph.GetSiteDrive(settings.cdnSiteId);
                            var result = await settings.GraphClient.Drives[siteDrive.Id].Items[templateItem.Id].Copy.PostAsync(requestBody);
                        }

                        try
                        {
                            var notesTab = await msgraph.GetTab(orderTeam, channel, "Notes");

                            if (notesTab != default(TeamsTab))
                            {
                                await msgraph.RemoveTab(orderTeam, channel, notesTab.Id);
                            }
                        }
                        catch (Exception ex)
                        {
                            log.LogError("Error removing default notes tab: " + ex.Message);
                        }

                        try
                        {
                            var onenoteTab = await msgraph.TabExists(orderTeam, channel, "M%F6tesanteckningar");

                            if (!onenoteTab)
                            {
                                DriveItem onenotefile = await msgraph.FindItem(customerGroup.groupDrive, channelFolder.Id, "ProjectMeetingNotes", false);

                                if (onenotefile != null)
                                {
                                    log.LogInformation("Add onenotetab with url " + onenotefile.WebUrl + " to channel " + channel.DisplayName + " in team " + orderTeam.DisplayName);
                                    await msgraph.AddChannelWebApp(orderTeam, channel, "M%F6tesanteckningar", onenotefile.WebUrl, onenotefile.WebUrl);
                                }
                            }
                        }
                        catch (Exception ex)
                        {
                            log.LogError("Error Adding Onenote Tab: " + ex.Message);
                        }
                    }
                }
                catch (Exception ex)
                {
                    log.LogError("Error copying project templates: " + ex.Message);
                }

                returnValue = true;
            }

            return returnValue;
        }

    }
}

