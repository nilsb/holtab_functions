using System;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Host;
using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Configuration;
using Azure.Identity;
using Microsoft.Graph;
using Newtonsoft.Json;
using CreateTeam.Models;
using System.Collections.Generic;
using RE = System.Text.RegularExpressions;
using System.Threading;
using System.Xml;
using PnP.Core.Services;
using System.Threading.Tasks;
using PnP.Core.Model.SharePoint;
using PnP.Core.QueryModel;
using Microsoft.ApplicationInsights;
using Microsoft.ApplicationInsights.Extensibility;
using Microsoft.ApplicationInsights.DataContracts;
using System.Linq;
using CreateTeam.Shared;
using Microsoft.Graph.Models;

namespace CreateTeam
{
    public class CreateTeam
    {

        private readonly IPnPContextFactory pnpContextFactory;
        private readonly TelemetryClient telemetryClient;

        public CreateTeam(IPnPContextFactory pnpContextFactory, TelemetryConfiguration telemetryConfiguration)
        {
            this.pnpContextFactory = pnpContextFactory;
            this.telemetryClient = new TelemetryClient(telemetryConfiguration);
        }

        [FunctionName("CreateTeam")]
        public async Task Run([QueueTrigger("createcustomer", Connection = "AzureWebJobsStorage")] string myQueueItem, Microsoft.Azure.WebJobs.ExecutionContext context, ILogger log)
        {
            var config = new ConfigurationBuilder()
              .SetBasePath(context.FunctionAppDirectory)
              .AddJsonFile("local.settings.json", optional: true, reloadOnChange: true)
              .AddEnvironmentVariables()
              .Build();

            string ClientID = config["ClientID"];
            string ClientSecret = config["ClientSecret"];
            string TenantID = config["TenantID"];
            string SqlConnectionString = config["SqlConnectionString"];

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

            telemetryClient.TrackEvent(new EventTelemetry($"Got create customer request with message: {myQueueItem}"));
            CustomerMessage msg = JsonConvert.DeserializeObject<CustomerMessage>(myQueueItem);
            Graph msgraph = new Graph(graphClient, log);
            Common common = new Common(graphClient, config, log, telemetryClient, msgraph);
            Customer createdCustomer = common.UpdateOrCreateDbCustomer(msg);
            FindCustomerResult customerResult = common.GetCustomer(msg.CustomerNo, msg.Type, msg.CustomerName);

            if (customerResult.Success)
            {
                if (customerResult.customer != null)
                {
                    Customer customer = customerResult.customer;
                    customer.Seller = msg.Responsible;

                    //Try to find a group for the customer or supplier
                    FindCustomerGroupResult findCustomerGroup = common.FindCustomerGroupAndDrive(msg.CustomerName, msg.CustomerNo, msg.Type);

                    if (findCustomerGroup.Success)
                    {
                        customer = findCustomerGroup.customer;
                        customer.GroupCreated = true;
                        customer.GroupID = findCustomerGroup.group.Id;
                        common.UpdateCustomer(customer, "group info");

                        //If the group already exists try to create the General folder.
                        //No action is taken if the folder already exists.
                        try
                        {
                            var generalFolder = await msgraph.CreateFolder(findCustomerGroup.group.Id, "General");
                            customer.GeneralFolderID = generalFolder.folder.Id;
                            customer.GeneralFolderCreated = true;
                            common.UpdateCustomer(customer, "general folder info");
                        }
                        catch (Exception ex)
                        {
                            telemetryClient.TrackException(ex);
                            telemetryClient.TrackTrace(new TraceTelemetry($"Failed to create general folder for {customer.Name} with error: " + ex.ToString()));
                        }

                        //Copy folder structure for customer or supplier depending on the type in the message
                        if (await common.CopyRootStructure(customer))
                        {
                            telemetryClient.TrackEvent(new EventTelemetry($"Created template folders"));
                            customer.CopiedRootStructure = true;
                            common.UpdateCustomer(customer, "root structure");
                        }

                        try
                        {
                            //Create custom document library columns
                            await CreateColumn(graphClient, findCustomerGroup.group, customer, log, config);
                            customer.CreatedColumnAdditionalInfo = true;
                            customer.CreatedColumnKundnummer = true;
                            customer.CreatedColumnNAVid = true;
                            customer.CreatedColumnProduktionsdokument = true;
                            common.UpdateCustomer(customer, "columns");
                        }
                        catch (Exception ex)
                        {
                            telemetryClient.TrackException(ex);
                            telemetryClient.TrackTrace(new TraceTelemetry($"Failed to create columns for {customer.Name} with error: " + ex.ToString()));
                        }
                    }
                    else
                    {
                        //If the group was not found, create group with team
                        CreateCustomerResult result = await common.CreateCustomerOrSupplier(customer);

                        if (result.group != null)
                        {
                            //If the group was created successfully
                            //Then create the general folder
                            telemetryClient.TrackEvent(new EventTelemetry($"Created group for customer or supplier: {customer.Name}"));

                            customer = result.customer;
                            customer.GroupCreated = true;
                            customer.GroupID = result.group.Id;
                            common.UpdateCustomer(customer, "group info");

                            try
                            {
                                var generalFolder = await msgraph.CreateFolder(result.group.Id, "General");
                                customer.GeneralFolderID = generalFolder.folder.Id;
                                customer.GeneralFolderCreated = true;
                                common.UpdateCustomer(customer, "general folder info");
                            }
                            catch (Exception ex)
                            {
                                telemetryClient.TrackException(ex);
                                telemetryClient.TrackTrace(new TraceTelemetry($"Failed to create general folder for {customer.Name} with error: " + ex.ToString()));
                            }

                            //Copy folder structure for customer or supplier depending on the type in the message
                            if (await common.CopyRootStructure(customer))
                            {
                                telemetryClient.TrackEvent(new EventTelemetry($"Created template folders"));
                                customer.CopiedRootStructure = true;
                                common.UpdateCustomer(customer, "root structure");
                            }

                            try
                            {
                                //Create custom document library columns
                                await CreateColumn(graphClient, result.group, customer, log, config);
                                customer.CreatedColumnAdditionalInfo = true;
                                customer.CreatedColumnKundnummer = true;
                                customer.CreatedColumnNAVid = true;
                                customer.CreatedColumnProduktionsdokument = true;
                                common.UpdateCustomer(customer, "columns");
                            }
                            catch (Exception ex)
                            {
                                telemetryClient.TrackException(ex);
                                telemetryClient.TrackTrace(new TraceTelemetry($"Failed to add columns to {customer.Name} with error: " + ex.ToString()));
                            }
                        }
                    }
                }
                else
                {
                    telemetryClient.TrackTrace(new TraceTelemetry($"Failed to find customer {msg.CustomerName} in database when creating group"));
                }
            }
            else
            {
                telemetryClient.TrackTrace(new TraceTelemetry($"Failed to find customer {msg.CustomerName} in database when creating group"));
            }

            log.LogInformation($"Queue trigger function processed: {myQueueItem}");
        }

        public async Task CreateColumn(GraphServiceClient graphClient, Group group, Customer customer, ILogger log, IConfigurationRoot config)
        {
            string ClientID = config["ClientID"];
            string ClientSecret = config["ClientSecret"];
            string TenantID = config["TenantID"];

            var drive = await graphClient.Groups[group.Id].Drive.GetAsync();
            if(drive == null)
            {
                return;
            }

            var root = await graphClient.Drives[drive.Id].Root.GetAsync();
            var list = await graphClient.Drives[drive.Id].List.GetAsync();
            string siteUrl = root.WebUrl.Substring(0, root.WebUrl.LastIndexOf("/"));

            try
            {
                telemetryClient.TrackEvent(new EventTelemetry($"Adding column Kundnummer to {customer.Name} ({customer.ExternalId})"));
                ColumnDefinition customerNoDef = new ColumnDefinition()
                {
                    Description = "Kundnummer",
                    Text = new TextColumn()
                    {
                        AllowMultipleLines = false,
                        TextType = "plain"
                    },
                    DefaultValue = new DefaultColumnValue() { Value = customer.ExternalId },
                    Name = "Kundnummer",
                    Hidden = false,
                    Required = false,
                    EnforceUniqueValues = false,
                    Indexed = true
                };

                var customerNoCol = await graphClient.Drives[drive.Id].List.Columns.PostAsync(customerNoDef);

                if(customerNoCol != null)
                {
                    using (PnPContext context = await pnpContextFactory.CreateAsync(new Uri(siteUrl)))
                    {
                        var targetList = await context.Web.Lists.GetByTitleAsync(list.DisplayName, l => l.Fields);

                        if (targetList != null)
                        {
                            foreach (var field in targetList.Fields.AsRequested())
                            {
                                if (field.Title == "Kundnummer")
                                {
                                    field.DefaultValue = customer.ExternalId;
                                    await field.UpdateAsync();
                                }
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                telemetryClient.TrackException(ex);
                telemetryClient.TrackTrace(new TraceTelemetry($"Failed to add column Kundnummer to {customer.Name} with error: " + ex.ToString()));
            }

            try
            {
                telemetryClient.TrackEvent(new EventTelemetry($"Adding column NAVid to {customer.Name} ({customer.ExternalId})"));
                ColumnDefinition navIdDef = new ColumnDefinition()
                {
                    Description = "NAVid",
                    Text = new TextColumn()
                    {
                        AllowMultipleLines = false,
                        TextType = "plain"
                    },
                    DefaultValue = new DefaultColumnValue() { Value = "-" },
                    Name = "NAVid",
                    Hidden = false,
                    EnforceUniqueValues = false,
                    Indexed = true
                };

                var navIdCol = await graphClient.Drives[drive.Id].List.Columns.PostAsync(navIdDef);
                
                if(navIdCol != null)
                {
                    using (PnPContext context = await pnpContextFactory.CreateAsync(new Uri(siteUrl)))
                    {
                        var targetList = await context.Web.Lists.GetByTitleAsync(list.DisplayName, l => l.Fields);

                        if (targetList != null)
                        {
                            foreach (var field in targetList.Fields.AsRequested())
                            {
                                if (field.Title == "NAVid")
                                {
                                    field.DefaultValue = "-";
                                    await field.UpdateAsync();
                                }
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                telemetryClient.TrackException(ex);
                telemetryClient.TrackTrace(new TraceTelemetry($"Failed to add column NAVid to {customer.Name} with error: " + ex.ToString()));
            }

            try
            {
                telemetryClient.TrackEvent(new EventTelemetry($"Adding column Produktionsdokument to {customer.Name} ({customer.ExternalId})"));
                var GroupSites = await graphClient.Groups[group.Id].Sites.GetAsync();

                ColumnDefinition isProdDef = new ColumnDefinition()
                {
                    Description = "Produktionsdokument",
                    Choice = new ChoiceColumn()
                    {
                        DisplayAs = "checkBoxes",
                        Choices = await GetProductionChoices(graphClient, log, config)
                    },
                    Name = "Produktionsdokument",
                    Hidden = false,
                    EnforceUniqueValues = false,
                    Indexed = false
                };

                var isProdCol = await graphClient.Drives[drive.Id].List.Columns.PostAsync(isProdDef);
            }
            catch (Exception ex)
            {
                telemetryClient.TrackException(ex);
                telemetryClient.TrackTrace(new TraceTelemetry($"Failed to add column Produktionsdokument to {customer.Name} with error: " + ex.ToString()));
            }
        }

        /// <summary>
        /// Get choices fo the column produktionsdokument from a list in the CDN site.
        /// </summary>
        /// <param name="graphClient"></param>
        /// <param name="log"></param>
        /// <param name="config"></param>
        /// <returns>A list of string values</returns>
        public async Task<List<string>> GetProductionChoices(GraphServiceClient graphClient, ILogger log, IConfigurationRoot config)
        {
            List<string> returnValue = new List<string>();
            var listItems = default(ListItemCollectionResponse);

            try
            {
                listItems = await graphClient.Sites[config["CdnSiteID"]]
                    .Lists[config["ProductionChoicesID"]]
                    .Items
                    .GetAsync();
            }
            catch (Exception ex)
            {
                telemetryClient.TrackException(ex);
                telemetryClient.TrackTrace(new TraceTelemetry($"Failed to get choices for column Produktionsdokument with error: " + ex.ToString()));
            }

            if (listItems?.Value?.Count > 0)
            {
                foreach(var item in listItems.Value)
                {
                    returnValue.Add(item.Fields.AdditionalData["Title"].ToString());
                }
            }

            return returnValue;
        }
    }
}
