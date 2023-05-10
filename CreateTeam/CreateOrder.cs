using System.Collections.Generic;
using RE = System.Text.RegularExpressions;
using Azure.Identity;
using CreateTeam.Models;
using Microsoft.Azure.WebJobs;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;
using Microsoft.Graph;
using Newtonsoft.Json;
using PnP.Core.Services;
using Microsoft.ApplicationInsights;
using Microsoft.ApplicationInsights.Extensibility;
using Microsoft.ApplicationInsights.DataContracts;
using System.Threading.Tasks;
using CreateTeam.Shared;
using System;
using System.Linq;
using System.Threading;
using System.Numerics;
using Microsoft.Graph.Models;

namespace CreateTeam
{
    public class CreateOrder
    {
        private readonly IPnPContextFactory pnpContextFactory;
        private readonly TelemetryClient telemetryClient;
        private string CDNTeamID;
        private string TenantID;
        private string cdnSiteId;
        private string SqlConnectionString;

        public CreateOrder(IPnPContextFactory pnpContextFactory, TelemetryConfiguration telemetryConfiguration)
        {
            this.pnpContextFactory = pnpContextFactory;
            this.telemetryClient = new TelemetryClient(telemetryConfiguration);
        }


        [FunctionName("CreateOrder")]
        public async Task Run([QueueTrigger("createorder", Connection = "AzureWebJobsStorage")]string myQueueItem, Microsoft.Azure.WebJobs.ExecutionContext context, ILogger log)
        {
            log.LogInformation("Got Create Order Request");

            var config = new ConfigurationBuilder()
              .SetBasePath(context.FunctionAppDirectory)
              .AddJsonFile("local.settings.json", optional: true, reloadOnChange: true)
              .AddEnvironmentVariables()
              .Build();

            string ClientID = config["ClientID"];
            string ClientSecret = config["ClientSecret"];
            TenantID = config["TenantID"];
            cdnSiteId = config["cdnSiteId"];
            string debugProjectNo = config["DebugProjectNo"];
            SqlConnectionString = config["SqlConnectionString"];
            CDNTeamID = config["CDNTeamID"];

            bool DebugProject = !string.IsNullOrEmpty(debugProjectNo) ? true : false;

            var options = new TokenCredentialOptions
            {
                AuthorityHost = AzureAuthorityHosts.AzurePublicCloud
            };

            var clientSecretCredential = new ClientSecretCredential(
                TenantID,
                ClientID,
                ClientSecret,
                options);
            var scopes = new[] { "https://graph.microsoft.com/.default" };
            var graphClient = new GraphServiceClient(clientSecretCredential, scopes);
            Graph msgraph = new Graph(graphClient, log);
            Common common = new Common(graphClient, config, log, telemetryClient, msgraph);
            Services.Log($"Got create order request with message: {myQueueItem}");
            telemetryClient.TrackEvent(new EventTelemetry($"Got create order request with message: {myQueueItem}"));
            telemetryClient.TrackEvent(new EventTelemetry($"Converting order to object"));
            Order order = JsonConvert.DeserializeObject<Order>(myQueueItem);

            if(string.IsNullOrEmpty(order.ExternalId))
                order.ExternalId = order.No; //backwards compatibility and problem with json deserialize

            Order listItem = null;
            Order dbOrder = null;

            //Find and update or create the order CDN post
            dbOrder = common.UpdateOrCreateDbOrder(order);
            listItem = common.GetOrderFromCDN(order.ExternalId);
            Site cdnSite = await graphClient.Sites[cdnSiteId].GetAsync();
            telemetryClient.TrackEvent(new EventTelemetry($"Get Customer and order folder if it exists"));
            FindOrderGroupAndFolder orderGroup = common.GetOrderGroupAndFolder(order.ExternalId);
            string addquery = "";

            if (orderGroup.Success && orderGroup.generalFolder != null && listItem != null)
            {
                //Set status fields for the database
                if (orderGroup.customer != null)
                {
                    listItem.Customer = orderGroup.customer;

                    if(orderGroup?.customer?.ID != Guid.Empty)
                    {
                        listItem.CustomerID = orderGroup.customer.ID;
                    }
                }

                if(orderGroup.orderDrive != null)
                {
                    listItem.DriveFound = true;
                    listItem.DriveID = orderGroup.orderDrive.Id;
                }

                if(orderGroup.generalFolder != null)
                {
                    listItem.GeneralFolderFound = true;
                }

                common.UpdateOrder(listItem, "group and drive info");

                //We found the customer/supplier with a general folder and the order item exists in the CDN
                if (orderGroup.orderTeam != null && orderGroup.customer != null)
                {
                    //try to add seller and project manager as owners
                    try
                    {
                        telemetryClient.TrackEvent(new EventTelemetry($"Trying to add {order.Seller} and {order.ProjectManager} to the team."));

                        if (!string.IsNullOrEmpty(order.Seller))
                        {
                            await msgraph.AddTeamMember(order.Seller, orderGroup.orderTeam.Id, "owner");
                        }

                        if (!string.IsNullOrEmpty(order.ProjectManager))
                        {
                            await msgraph.AddTeamMember(order.ProjectManager, orderGroup.orderTeam.Id, "owner");
                        }
                    }
                    catch (Exception ex)
                    {
                        telemetryClient.TrackException(ex);
                    }
                }
                else
                {
                    telemetryClient.TrackEvent(new EventTelemetry($"No team was found for {orderGroup.orderGroup.DisplayName}."));
                }

                string parentName = "";

                switch (order.Type)
                {
                    case "Order":
                        parentName = "Order";
                        break;
                    case "Project":
                        parentName = "Order";
                        break;
                    case "Quote":
                        parentName = "Offert";
                        RE.Match orderMatch = RE.Regex.Match(order.ExternalId, @"^([A-Z]?\d+)");

                        if (orderMatch.Success)
                        {
                            telemetryClient.TrackEvent(new EventTelemetry($"Changed order no for quote: {order.ExternalId} to: {orderMatch.Value}"));
                            order.ExternalId = orderMatch.Value;
                        }

                        break;
                    case "Offer":
                        parentName = "Offert";
                        RE.Match offerMatch = RE.Regex.Match(order.ExternalId, @"^([A-Z]?\d+)");

                        if (offerMatch.Success)
                        {
                            telemetryClient.TrackEvent(new EventTelemetry($"Changed order no for quote: {order.ExternalId} to: {offerMatch.Value}"));
                            order.ExternalId = offerMatch.Value;
                        }

                        break;
                    case "Purchase":
                        parentName = "Beställning";
                        break;
                    default:
                        break;
                }

                telemetryClient.TrackEvent(new EventTelemetry($"Creating order for {order.ExternalId}"));
                DriveItem orderParent = msgraph.FindItem(orderGroup.orderDrive, orderGroup.generalFolder.Id, parentName, false).Result;

                //the order folder doesn't exist so create it
                telemetryClient.TrackEvent(new EventTelemetry($"Get template folders to create"));
                var foldersToCreate = await GetFoldersToCreate(listItem, log, msgraph, cdnSiteId);

                if (orderGroup.orderFolder == null)
                {
                    if (orderParent != null)
                    {
                        listItem.Status = "Incomplete";
                        listItem.GroupFound = true;

                        if (order.Type == "Order" || order.Type == "Project")
                        {
                            listItem.OffersFolderFound = false;
                            listItem.PurchaseFolderFound = false;
                            listItem.OrdersFolderFound = true;
                        }
                        else if (order.Type == "Quote" || order.Type == "Offer")
                        {
                            listItem.OffersFolderFound = true;
                            listItem.PurchaseFolderFound = false;
                            listItem.OrdersFolderFound = false;
                        }
                        else if (order.Type == "Purchase")
                        {
                            listItem.OffersFolderFound = false;
                            listItem.PurchaseFolderFound = true;
                            listItem.OrdersFolderFound = false;
                        }

                        common.UpdateOrder(listItem, "parent folder info");

                        telemetryClient.TrackEvent(new EventTelemetry($"Found orders folder in group drive {orderGroup.orderGroup.DisplayName}"));
                        var existingorderFolder = msgraph.FindItem(orderGroup.orderDrive, orderParent.Id, order.ExternalId, false).Result;
                        CreateFolderResult orderFolder = new CreateFolderResult() { Success = false };

                        if (existingorderFolder == null)
                        {
                            telemetryClient.TrackEvent(new EventTelemetry($"Creating new order folder for {order.ExternalId} because it doesn't exist"));
                            orderFolder = msgraph.CreateFolder(orderGroup.orderGroup.Id, orderParent.Id, order.ExternalId).Result;
                        }
                        else
                        {
                            telemetryClient.TrackEvent(new EventTelemetry($"Order folder for {order.ExternalId} already exists"));
                            orderFolder.folder = existingorderFolder;
                            orderFolder.Success = true;
                        }

                        if (orderFolder.Success)
                        {
                            listItem.CreatedFolder = true;
                            listItem.FolderID = orderFolder.folder.Id;
                            listItem.Status = "Folder Created";
                            common.UpdateOrder(listItem, "folder info");

                            telemetryClient.TrackEvent(new EventTelemetry($"Order folder for {order.ExternalId} was created or found"));
                            
                            if (listItem != null)
                            {
                                telemetryClient.TrackEvent(new EventTelemetry($"Set order item for {order.ExternalId} as handled"));
                                //_ = common.SetOrderItemHandled(true, listItem, "Folder created").Result;
                            }

                            bool copyStructure = true;

                            foreach (var folder in foldersToCreate)
                            {
                                telemetryClient.TrackEvent(new EventTelemetry($"Copy template folder {folder.Name} to order folder for {order.ExternalId}."));
                                CreateFolderResult result = await msgraph.CopyFolder(orderGroup.orderGroup.Id, orderFolder.folder.Id, folder, true, false);

                                if (result.Success)
                                {
                                    copyStructure &= true;
                                }
                                else
                                {
                                    copyStructure &= false;
                                }
                            }

                            if(order.Type == "Project")
                            {
                                if((DebugProject && order.No == debugProjectNo) || !DebugProject)
                                {
                                    _ = await CreateProjectTabs(cdnSite, orderGroup, orderFolder.folder, order, log, graphClient, msgraph, cdnSiteId, TenantID);
                                }
                            }

                            listItem.StructureCreated = copyStructure;
                            listItem.Handled = copyStructure;
                            listItem.Status = "Folder created and structure copied";
                            common.UpdateOrder(listItem, "status");
                        }
                        else
                        {
                            listItem.Status = "Failed to create folder";
                            listItem.Handled = false;
                            common.UpdateOrder(listItem, "status");

                            if (listItem != null)
                            {
                                telemetryClient.TrackEvent(new EventTelemetry($"Set status failed to create folder on CDN order item for {order.ExternalId}."));
                                //_ = common.SetOrderItemHandled(false, listItem, "Failed to create folder").Result;
                            }

                            telemetryClient.TrackTrace($"Unable to create order folder {order.ExternalId} in team {orderGroup.orderGroup.DisplayName}");
                        }
                    }
                    else
                    {
                        listItem.Status = "Unable to find parent folder";
                        listItem.Handled = false;
                        common.UpdateOrder(listItem, "status");
                        telemetryClient.TrackTrace($"Unable to find parent folder for orders in team {orderGroup.orderGroup.DisplayName}");
                    }
                }
                else
                {
                    listItem.Status = "Folder already existed";
                    listItem.Handled = true;
                    listItem.FolderID = orderGroup.orderFolder.Id;
                    listItem.DriveFound = true;
                    listItem.DriveID = orderGroup.orderDrive.Id;
                    listItem.CreatedFolder = true;
                    listItem.GroupFound = true;
                    listItem.GeneralFolderFound = true;
                    common.UpdateOrder(listItem, "status");

                    if (order.Type == "Project")
                    {
                        if ((DebugProject && order.No == debugProjectNo) || !DebugProject)
                        {
                            _ = await CreateProjectTabs(cdnSite, orderGroup, orderGroup.orderFolder, order, log, graphClient, msgraph, cdnSiteId, TenantID);
                        }
                    }

                    if (listItem != null)
                    {
                        telemetryClient.TrackEvent(new EventTelemetry($"Set status folder created on CDN order item for {order.ExternalId}."));
                        //_ = common.SetOrderItemHandled(true, listItem, "Folder created").Result;
                    }

                    telemetryClient.TrackEvent(new EventTelemetry($"Order folder for {order.ExternalId} already exists."));
                }
            }
            else if (!orderGroup.Success || orderGroup.generalFolder == null)
            {
                listItem.ID = Guid.NewGuid();
                listItem.OrdersFolderFound = false;
                listItem.OffersFolderFound = false;
                listItem.PurchaseFolderFound = false;
                listItem.FolderID = "";
                listItem.DriveFound = false;
                listItem.Customer = null;
                listItem.GeneralFolderFound = false;
                listItem.GroupFound = false;
                listItem.StructureCreated = false;
                listItem.CreatedFolder = false;
                listItem.CustomerNo = order.CustomerNo;
                listItem.CustomerType = order.CustomerType;
                listItem.Status = "Error creating order";
                listItem.Handled = false;

                try
                {
                    listItem.Created = DateTime.Now;
                    common.UpdateOrCreateDbOrder(listItem);
                }
                catch (Exception ex)
                {
                    telemetryClient.TrackException(ex);
                    telemetryClient.TrackTrace(new TraceTelemetry($"Failed to add order {order.ExternalId} to database with query: {addquery}"));
                }

                if (listItem != null)
                {
                    telemetryClient.TrackEvent(new EventTelemetry($"Set status customer or supplier missing on CDN order item for {order.ExternalId}."));
                    //await common.SetOrderItemHandled(false, listItem, "Customer or supplier missing");
                }

                telemetryClient.TrackTrace($"Unable to find group and team for {order.ExternalId}");
            }

            log.LogInformation($"C# Queue trigger function processed: {myQueueItem}");
        }

        public async Task<List<DriveItem>> GetFoldersToCreate(Order order, ILogger log, Graph msgraph, string CDNSiteID)
        {
            List<DriveItem> foldersToCreate = new List<DriveItem>();
            var cdnDrive = await msgraph.GetSiteDrive(CDNSiteID);

            if(cdnDrive != null)
            {
                DriveItem folder = await msgraph.FindItem(cdnDrive, "Dokumentstruktur " + order.Type, false);
                List<DriveItem> folderChildren = await msgraph.GetDriveFolderChildren(cdnDrive, folder, true);
                foldersToCreate.AddRange(folderChildren);
            }

            return foldersToCreate;
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
                
                foreach(DriveItem folderChild in folderChildren)
                {
                    foldersToCreate.Add(folderChild);
                }
            }

            return foldersToCreate;
        }

        public async Task<bool> CreateProjectTabs(Site cdnSite, FindOrderGroupAndFolder orderGroup, DriveItem orderFolder, Order order, ILogger log, GraphServiceClient graph, Graph msgraph, string CDNSiteID, string TenantID)
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

                        if(app != null)
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
                    Services.Log("Looking for planner tab");
                    string tabName = "Checklista - Projekt " + order.No;

                    //try adding checklist
                    var checklistFolderTab = await msgraph.TabExists(orderGroup.orderTeam, channel, tabName);

                    if (!checklistFolderTab)
                    {
                        Services.Log("Looking for template planner");
                        //Try to find the template checklist
                        var planTemplate = await msgraph.PlanExists(graph, CDNTeamID, "Checklista - Projektledning Template");

                        if (planTemplate != null)
                        {
                            Services.Log("Looking for existing planner");
                            //found template so create the plan if it doesn't exist
                            var existingPlan = await msgraph.PlanExists(graph, orderGroup.orderGroup.Id, tabName);
                            
                            if (existingPlan == null)
                            {
                                Services.Log("Creating new plan");
                                existingPlan = await msgraph.CreatePlanAsync(graph, orderGroup.orderGroup.Id, tabName);

                                //copy buckets and tasks
                                var buckets = await msgraph.GetBucketsAsync(graph, planTemplate.Id);

                                foreach (var bucket in buckets)
                                {
                                    await msgraph.CopyBucketAsync(graph, bucket, existingPlan.Id);
                                }
                                Services.Log("Copied template");

                                //create the planner tab
                                await msgraph.CreatePlannerTabInChannelAsync(graph, TenantID, orderGroup.orderTeam.Id, tabName, channel.Id, existingPlan.Id);
                                Services.Log("Creating planner tab");
                            }
                            else
                            {
                                //create the planner tab
                                await msgraph.CreatePlannerTabInChannelAsync(graph, TenantID, orderGroup.orderTeam.Id, tabName, channel.Id, existingPlan.Id);
                                Services.Log("Creating planner tab");
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

                            log.LogInformation($"Copy template item {templateItem.Name} to project folder for {order.ExternalId}.");
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
                                    log.LogInformation("Add onenotetab with url " + onenotefile.WebUrl + " to channel " + channel.DisplayName + " in team " + orderGroup.orderTeam.DisplayName);
                                    await msgraph.AddChannelApp(orderGroup.orderTeam, app, channel, "Mötesanteckningar", null, onenotefile.WebUrl, onenotefile.WebUrl, null);
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
