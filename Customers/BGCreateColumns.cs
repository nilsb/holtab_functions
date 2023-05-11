using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using Azure.Identity;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Extensions.Http;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;
using Microsoft.Graph;
using Microsoft.Graph.Models;
using Newtonsoft.Json;
using Shared;
using Shared.Models;

namespace CreateTeam
{
    public class BGCreateColumns
    {
        [FunctionName("BGCreateColumns")]
        public async Task<IActionResult> Run([HttpTrigger(AuthorizationLevel.Function, new string[] { "post" }, Route = null)] HttpRequest req,
            Microsoft.Azure.WebJobs.ExecutionContext context,
            ILogger log,
            IConfiguration config)
        {
            string Message = await new StreamReader(req.Body).ReadToEndAsync();

            log.LogInformation($"Create columns queue trigger function processed message: {Message}");
            Settings settings = new Settings(context, log);
            Graph msGraph = new Graph(ref settings);
            Common common = new Common(ref settings, ref msGraph);
            log.LogTrace($"Got copy root structure request with message: {Message}");

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
                        await CreateColumn(settings, msGraph, common, findCustomerGroup.group, customer);
                        customer.CreatedColumnAdditionalInfo = true;
                        customer.CreatedColumnKundnummer = true;
                        customer.CreatedColumnNAVid = true;
                        customer.CreatedColumnProduktionsdokument = true;
                        common.UpdateCustomer(customer, "columns");

                        return new OkObjectResult(JsonConvert.SerializeObject(Message));
                    }
                    catch (Exception ex)
                    {
                        log.LogError(ex.ToString());
                        log.LogTrace($"Failed to add columns to {customer.Name} with error: " + ex.ToString());
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

        public async Task CreateColumn(Settings settings, Graph msGraph, Common common, Group? group, Customer? customer)
        {
            var drive = await msGraph.GetGroupDrive(group);

            if (drive == null)
            {
                return;
            }

            var root = await settings.GraphClient.Drives[drive.Id].Root.GetAsync();
            var list = await settings.GraphClient.Drives[drive.Id].List.GetAsync();
            string siteUrl = drive.WebUrl.Substring(0, drive.WebUrl.LastIndexOf("/"));
            var groupsite = group.Sites.FirstOrDefault();

            try
            {
                settings.log.LogTrace($"Adding column Kundnummer to {customer.Name} ({customer.ExternalId})");
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

                var customerNoCol = await settings.GraphClient.Sites[groupsite.Id].Lists[list.Id].Columns.PostAsync(customerNoDef);
            }
            catch (Exception ex)
            {
                settings.log.LogError(ex.ToString());
                settings.log.LogTrace($"Failed to add column Kundnummer to {customer.Name} with error: " + ex.ToString());
            }

            try
            {
                settings.log.LogTrace($"Adding column NAVid to {customer.Name} ({customer.ExternalId})");
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

                var navIdCol = await settings.GraphClient.Sites[groupsite.Id].Lists[list.Id].Columns.PostAsync(navIdDef);
            }
            catch (Exception ex)
            {
                settings.log.LogError(ex.ToString());
                settings.log.LogTrace($"Failed to add column NAVid to {customer.Name} with error: " + ex.ToString());
            }

            try
            {
                settings.log.LogTrace($"Adding column Produktionsdokument to {customer.Name} ({customer.ExternalId})");

                ColumnDefinition isProdDef = new ColumnDefinition()
                {
                    Description = "Produktionsdokument",
                    Choice = new ChoiceColumn()
                    {
                        DisplayAs = "checkBoxes",
                        Choices = await GetProductionChoices(settings)
                    },
                    Name = "Produktionsdokument",
                    Hidden = false,
                    EnforceUniqueValues = false,
                    Indexed = false
                };

                var isProdCol = await settings.GraphClient.Sites[groupsite.Id].Lists[list.Id].Columns.PostAsync(isProdDef);
            }
            catch (Exception ex)
            {
                settings.log.LogError(ex.ToString());
                settings.log.LogTrace($"Failed to add column Produktionsdokument to {customer.Name} with error: " + ex.ToString());
            }
        }

        /// <summary>
        /// Get choices fo the column produktionsdokument from a list in the CDN site.
        /// </summary>
        /// <param name="graphClient"></param>
        /// <param name="log"></param>
        /// <param name="config"></param>
        /// <returns>A list of string values</returns>
        public async Task<List<string>> GetProductionChoices(Settings settings)
        {
            List<string> returnValue = new List<string>();
            ListItemCollectionResponse listItems = default(ListItemCollectionResponse);

            try
            {
                listItems = await settings.GraphClient.Sites[settings.cdnSiteId]
                    .Lists[settings.ProductionChoicesListID]
                    .Items
                    .GetAsync(config => {
                        config.QueryParameters.Expand = new string[] { "Fields" };
                    });
            }
            catch (Exception ex)
            {
                settings.log?.LogError(ex.ToString());
                settings.log?.LogTrace($"Failed to get choices for column Produktionsdokument with error: " + ex.ToString());
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
