using System;
using System.IO;
using System.Text;
using System.Threading.Tasks;
using AutoMapper;
using Azure.Identity;
using Azure.Messaging.ServiceBus;
using CreateTeam.Models;
using CreateTeam.Shared;
using Microsoft.ApplicationInsights;
using Microsoft.ApplicationInsights.DataContracts;
using Microsoft.ApplicationInsights.Extensibility;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Server.IIS.Core;
using Microsoft.Azure.Amqp.Encoding;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Extensions.Http;
using Microsoft.Azure.WebJobs.Host;
using Microsoft.Azure.WebJobs.ServiceBus;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;
using Microsoft.Graph;
using Microsoft.Graph.Models;
using Newtonsoft.Json;

namespace CreateTeam
{
    public class BGCreateGroup
    {
        private readonly TelemetryClient telemetryClient;

        public BGCreateGroup(TelemetryConfiguration telemetryConfiguration)
        {
            this.telemetryClient = new TelemetryClient(telemetryConfiguration);
        }

        [FunctionName("BGCreateGroup")]
        public async Task<IActionResult> Run([HttpTrigger(AuthorizationLevel.Function, new string[] { "post" }, Route = null)] HttpRequest req,
            Microsoft.Azure.WebJobs.ExecutionContext context,
            ILogger log)
        {
            string Message = await new StreamReader(req.Body).ReadToEndAsync();

            log.LogInformation($"Create group queue trigger function processed message: {Message}");
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
            telemetryClient.TrackEvent(new EventTelemetry($"Got create group request with message: {Message}"));

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
                    //Set the status group created on the customer object
                    customer.GroupCreated = true;

                    //if the drive was found set the drive id in the output message
                    if (findCustomerGroup.groupDrive != null && findCustomerGroup.groupDrive != default(Drive))
                    {
                        //set the status and drive id in the customer object
                        customer.DriveID = findCustomerGroup.groupDrive.Id;
                    }
                    else
                    {
                        //something went wrong because the default drive was not available. epic failiure
                        return new UnprocessableEntityObjectResult(Message);
                    }

                    //check if the general folder was found
                    if (findCustomerGroup.generalFolder != null && findCustomerGroup.generalFolder != default(DriveItem))
                    {
                        //set the status and if of general folder in the customer object
                        customer.GeneralFolderCreated = true;
                        customer.GeneralFolderID = findCustomerGroup.generalFolder.Id;

                        //update the database with the new customer information
                        common.UpdateCustomer(customer, "group and drive info");

                        //the group and general folder exists so continue
                        return new OkObjectResult(Message);
                    }
                    else
                    {
                        //if the general folder was not found try to create it
                        try
                        {
                            CreateFolderResult generalFolder = await msGraph.CreateFolder(findCustomerGroup.group.Id, "General");
                            customer.GeneralFolderID = generalFolder.folder.Id;
                            customer.GeneralFolderCreated = true;

                            //update the database with the new customer information
                            common.UpdateCustomer(customer, "group and drive info");

                            //the general folder and group exists so continue
                            return new OkObjectResult(Message);
                        }
                        catch (Exception ex)
                        {
                            telemetryClient.TrackException(ex);
                            telemetryClient.TrackTrace(new TraceTelemetry($"Failed to create general folder for {customer.Name} with error: " + ex.ToString()));
                        }

                        //couldn't create the general folder so we need to try this all over again.
                        if (string.IsNullOrEmpty(customer.GeneralFolderID))
                        {
                            return new UnprocessableEntityObjectResult(Message);
                        }
                    }

                    return new UnprocessableEntityObjectResult(Message);
                }
                else
                {
                    //If the group was not found, create group
                    CreateCustomerResult result = await common.CreateCustomerGroup(customer);

                    if (result.Success && result.customer != null && result.customer != default(Customer))
                    {
                        //group create was sent successfully
                        //if the drive was found then the group ws created super fast
                        if (result.customer.DriveID != null)
                        {
                            customer = result.customer;
                            customer.GroupCreated = true;

                            //try to create the general folder;
                            try
                            {
                                CreateFolderResult generalFolder = await msGraph.CreateFolder(result.group.Id, "General");

                                if (generalFolder.Success)
                                {
                                    customer.GeneralFolderID = generalFolder.folder.Id;
                                    customer.GeneralFolderCreated = true;
                                }
                            }
                            catch (Exception ex)
                            {
                                telemetryClient.TrackException(ex);
                                telemetryClient.TrackTrace(new TraceTelemetry($"Failed to create general folder for {customer.Name} with error: " + ex.ToString()));
                            }

                            //the general folder couldn't be created so we need to try this all over again
                            if (string.IsNullOrEmpty(customer.GeneralFolderID))
                            {
                                return new UnprocessableEntityObjectResult(Message);
                            }

                            //update the database with the new customer information
                            common.UpdateCustomer(customer, "group and drive info");

                            //everything went ok so send message to assign owner and copy root structure
                            return new OkObjectResult(Message);
                        }
                        else
                        {
                            //if the group was not yet available we wait for it
                            //the drive couldn't be found so wait for the group to become available
                            return new UnprocessableEntityObjectResult(Message);
                        }
                    }
                    else
                    {
                        //the group could not be created. epic failure.
                        return new UnprocessableEntityObjectResult(Message);
                    }
                }

            }
            else
            {
                //customer not found in DB
                return new UnprocessableEntityObjectResult(Message);
            }
        }
    }
}
