using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using Azure.Identity;
using Azure.Messaging.ServiceBus;
using CreateTeam.Models;
using CreateTeam.Shared;
using Microsoft.ApplicationInsights;
using Microsoft.ApplicationInsights.DataContracts;
using Microsoft.ApplicationInsights.Extensibility;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Azure.Cosmos.Serialization.HybridRow;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Extensions.Http;
using Microsoft.Azure.WebJobs.Host;
using Microsoft.Azure.WebJobs.ServiceBus;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;
using Microsoft.Graph;
using Microsoft.Graph.Models;
using Newtonsoft.Json;
using PnP.Core.QueryModel;
using PnP.Core.Services;

namespace CreateTeam
{
    public class BGCreateColumns
    {
        private readonly TelemetryClient telemetryClient;
        private readonly IPnPContextFactory pnpContextFactory;

        public BGCreateColumns(IPnPContextFactory pnpContextFactory, TelemetryConfiguration telemetryConfiguration)
        {
            this.pnpContextFactory = pnpContextFactory;
            this.telemetryClient = new TelemetryClient(telemetryConfiguration);
        }

        [FunctionName("BGCreateColumns")]
        public async Task<IActionResult> Run([HttpTrigger(AuthorizationLevel.Function, new string[] { "post" }, Route = null)] HttpRequest req,
            Microsoft.Azure.WebJobs.ExecutionContext context,
            ILogger log)
        {
            string Message = await new StreamReader(req.Body).ReadToEndAsync();

            log.LogInformation($"Create columns queue trigger function processed message: {Message}");
            var config = new ConfigurationBuilder()
              .SetBasePath(context.FunctionAppDirectory)
              .AddJsonFile("local.settings.json", optional: true, reloadOnChange: true)
              .AddEnvironmentVariables()
              .Build();
            string ClientID = config["ClientID"];
            string ClientSecret = config["ClientSecret"];
            string TenantID = config["TenantID"];
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
            Graph msGraph = new Graph(graphClient, log);
            Common common = new Common(graphClient, config, log, telemetryClient, msGraph);
            telemetryClient.TrackEvent(new EventTelemetry($"Got copy root structure request with message: {Message}"));

            //Parse the incoming message into JSON
            CustomerQueueMessage customerQueueMessage = JsonConvert.DeserializeObject<CustomerQueueMessage>(Message);

            //Get customer object from database
            FindCustomerResult findCustomer = common.GetCustomer(customerQueueMessage.ExternalId, customerQueueMessage.Type, customerQueueMessage.Name);

            if (findCustomer.Success && findCustomer.customer != null && findCustomer.customer != default(Customer))
            {
                Customer customer = findCustomer.customer;

                //Try to find the group and drive for the customer
                //This also assigns GroupId, DriveID and GeneralFolderID in the database if it was missing
                //The returned object contains the group object, the drive object, the root folder object and the general folder object
                FindCustomerGroupResult findCustomerGroup = await common.FindCustomerGroupAndDrive(customer);

                //if the group was found
                if (findCustomerGroup.Success && findCustomerGroup.group != null && findCustomerGroup.group != default(Group))
                {
                    try
                    {
                        //Create custom document library columns
                        await CreateColumn(graphClient, msGraph, common, findCustomerGroup.group, customer, log, config);
                        customer.CreatedColumnAdditionalInfo = true;
                        customer.CreatedColumnKundnummer = true;
                        customer.CreatedColumnNAVid = true;
                        customer.CreatedColumnProduktionsdokument = true;
                        common.UpdateCustomer(customer, "columns");

                        return new OkObjectResult(JsonConvert.SerializeObject(Message));
                    }
                    catch (Exception ex)
                    {
                        telemetryClient.TrackException(ex);
                        telemetryClient.TrackTrace(new TraceTelemetry($"Failed to add columns to {customer.Name} with error: " + ex.ToString()));
                    }

                    return new UnprocessableEntityObjectResult(JsonConvert.SerializeObject(Message));
                }
                else
                {
                    return new UnprocessableEntityObjectResult(JsonConvert.SerializeObject(Message));
                }
            }
            else
            {
                return new BadRequestObjectResult(JsonConvert.SerializeObject(Message));
            }
        }

        public async Task CreateColumn(GraphServiceClient graphClient, Graph msGraph, Common common, Group group, Customer customer, ILogger log, IConfigurationRoot config)
        {
            var drive = await msGraph.GetGroupDrive(group);

            if (drive == null)
            {
                return;
            }

            var root = await graphClient.Drives[drive.Id].Root.GetAsync();
            var list = await graphClient.Drives[drive.Id].List.GetAsync();
            string siteUrl = drive.WebUrl.Substring(0, drive.WebUrl.LastIndexOf("/"));
            var groupsite = group.Sites.FirstOrDefault();

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

                var customerNoCol = await graphClient.Sites[groupsite.Id].Lists[list.Id].Columns.PostAsync(customerNoDef);

                //if the column was created set the default value
                if (customerNoCol != null)
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

                var navIdCol = await graphClient.Sites[groupsite.Id].Lists[list.Id].Columns.PostAsync(navIdDef);

                if (navIdCol != null)
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

                var isProdCol = await graphClient.Sites[groupsite.Id].Lists[list.Id].Columns.PostAsync(isProdDef);
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
            ListItemCollectionResponse listItems = default(ListItemCollectionResponse);

            try
            {
                listItems = await graphClient.Sites[config["CdnSiteID"]]
                    .Lists[config["ProductionChoicesID"]]
                    .Items
                    .GetAsync(config => {
                        config.QueryParameters.Expand = new string[] { "Fields" };
                    });
            }
            catch (Exception ex)
            {
                telemetryClient.TrackException(ex);
                telemetryClient.TrackTrace(new TraceTelemetry($"Failed to get choices for column Produktionsdokument with error: " + ex.ToString()));
            }

            if (listItems != null && listItems?.Value?.Count > 0)
            {
                foreach (var item in listItems.Value)
                {
                    returnValue.Add(item.Fields.AdditionalData["Title"].ToString());
                }
            }

            return returnValue;
        }

    }
}
