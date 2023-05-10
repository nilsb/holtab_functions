using System;
using System.Threading.Tasks;
using RE = System.Text.RegularExpressions;
using Microsoft.Azure.WebJobs;
using Microsoft.Extensions.Logging;
using Newtonsoft.Json;
using Microsoft.Extensions.Configuration;
using System.Collections.Generic;
using Azure.Identity;
using Microsoft.Graph;
using CreateTeam.Models;
using Microsoft.ApplicationInsights;
using Microsoft.ApplicationInsights.Extensibility;
using Microsoft.ApplicationInsights.DataContracts;
using CreateTeam.Shared;
using System.Linq;
using Portable.Xaml.Markup;
using Microsoft.Graph.Models;

namespace CreateTeam
{
    public class HandleEmail
    {
        private readonly TelemetryClient telemetryClient;

        public HandleEmail(TelemetryConfiguration telemetryConfiguration)
        {
            this.telemetryClient = new TelemetryClient(telemetryConfiguration);
        }

        [FunctionName("HandleEmail")]
        public async Task Run(
            [QueueTrigger("receive", Connection = "AzureWebJobsStorage")] string myQueueItem,
            Microsoft.Azure.WebJobs.ExecutionContext context, 
            ILogger log)
        {
            log.LogInformation("Got handle email request with message " + myQueueItem);

            HandleEmailMessage data = new HandleEmailMessage();

            try
            {
                telemetryClient.TrackEvent(new EventTelemetry($"Got handle email request with message: {myQueueItem}"));
                data = JsonConvert.DeserializeObject<HandleEmailMessage>(myQueueItem);
            }
            catch (Exception ex)
            {
                telemetryClient.TrackException(ex);
                telemetryClient.TrackTrace("Failed to convert queue message to object with error: " + ex.ToString());
            }

            var config = new ConfigurationBuilder()
              .SetBasePath(context.FunctionAppDirectory)
              .AddJsonFile("local.settings.json", optional: true, reloadOnChange: true)
              .AddEnvironmentVariables()
              .Build();

            string ClientID = config["ClientID"];
            string ClientSecret = config["ClientSecret"];
            string TenantID = config["TenantID"];
            string CDNTeamID = config["CDNTeamID"];
            string HistoryMonths = config["HistoryMonths"];
            int historyMonths = 0;
            Int32.TryParse(HistoryMonths, out historyMonths);

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
            string orderNo;
            string customerNo = "";

            if (string.IsNullOrEmpty(data.Title))
            {
                orderNo = common.FindOrderNoInString(data.Filename);

                if (!string.IsNullOrEmpty(orderNo))
                {
                    telemetryClient.TrackEvent(new EventTelemetry($"Found orderno: {orderNo} in filename: {data.Filename}"));
                }
                else
                {
                    customerNo = common.FindCustomerNoInString(data.Filename);

                    if (!string.IsNullOrEmpty(customerNo))
                    {
                        telemetryClient.TrackEvent(new EventTelemetry($"Found customer no: {customerNo} in filename: {data.Filename}"));
                    }
                }
            }
            else
            {
                orderNo = common.FindOrderNoInString(data.Title);

                if (!string.IsNullOrEmpty(orderNo))
                {
                    telemetryClient.TrackEvent(new EventTelemetry($"Found orderno: {orderNo} in email subject: {data.Title}"));
                }
                else
                {
                    customerNo = common.FindCustomerNoInString(data.Title);

                    if (!string.IsNullOrEmpty(customerNo))
                    {
                        telemetryClient.TrackEvent(new EventTelemetry($"Found customer no: {customerNo} in subject: {data.Title}"));
                    }
                }
            }

            //handle order related emails
            if (!string.IsNullOrEmpty(orderNo))
            {
                FindOrderGroupAndFolder orderFolder = common.GetOrderGroupAndFolder(orderNo);

                if (orderFolder.Success && orderFolder.orderFolder != null)
                {
                    telemetryClient.TrackEvent(new EventTelemetry($"Found order group: {orderFolder.orderTeam.DisplayName} for order no: {orderNo}"));

                    //get cdn group
                    Group cdnGroup = await graphClient.Groups[CDNTeamID].GetAsync();

                    if (cdnGroup != null)
                    {
                        //get drive for cdn group
                        Drive cdnDrive = await msgraph.GetGroupDrive(cdnGroup);

                        if (cdnDrive != null)
                        {
                            telemetryClient.TrackEvent(new EventTelemetry($"Found CDN team and drive"));
                            
                            //Loop through email folders
                            for(int i = 0; i <= historyMonths; i++)
                            {
                                _ = await ProcessCDNFiles("General", msgraph, cdnDrive, data, orderNo, i, CDNTeamID, orderFolder);
                                _ = await ProcessCDNFiles("Salesemails", msgraph, cdnDrive, data, orderNo, i, CDNTeamID, orderFolder);
                            }
                        }
                    }
                }
                else
                {
                    if (!orderFolder.Success)
                    {
                        if (!string.IsNullOrEmpty(data.Sender))
                        {
                            Chat newChat = await msgraph.CreateOneOnOneChat(new List<string>
                            {
                                data.Sender,
                                config["SendFrom"]
                            });

                            if (newChat != null)
                            {
                                ChatMessage msg = await msgraph.SendOneOnOneMessage(newChat, config["MessageNoCustomer"].Replace("<orderno>", orderNo));
                            }
                        }

                        telemetryClient.TrackTrace(new TraceTelemetry($"Unable to find order group for {orderNo}"));
                    }
                    else if (orderFolder.orderFolder == null)
                    {
                        if (!string.IsNullOrEmpty(data.Sender))
                        {
                            Chat newChat = await msgraph.CreateOneOnOneChat(new List<string>
                            {
                                data.Sender,
                                config["SendFrom"]
                            });

                            if (newChat != null)
                            {
                                ChatMessage msg = await msgraph.SendOneOnOneMessage(newChat, config["MessageNoOrderFolder"].Replace("<orderno>", orderNo));
                            }
                        }

                        telemetryClient.TrackTrace(new TraceTelemetry($"Unable to find order folder for {orderNo} in group {orderFolder.orderGroup.DisplayName}"));
                    }
                }
            }

            if (!string.IsNullOrEmpty(customerNo))
            {
                FindCustomerResult customerResult = common.GetCustomer(customerNo, "Supplier");

                if (customerResult.Success && customerResult.customers.Count > 0)
                {
                    Customer dbCustomer = customerResult.customers.OrderByDescending(c => c.Created).Take(1).FirstOrDefault();

                    if(dbCustomer != null)
                    {
                        telemetryClient.TrackEvent(new EventTelemetry($"Found customer {dbCustomer.Name} in CDN"));
                        FindCustomerGroupResult customerGroupResult = common.FindCustomerGroupAndDrive(dbCustomer.Name, dbCustomer.ExternalId, dbCustomer.Type);

                        if (customerGroupResult.Success && customerGroupResult.group != null)
                        {
                            telemetryClient.TrackEvent(new EventTelemetry($"Found customer group and drive for {dbCustomer.Name}"));
                            //find email destination folder
                            DriveItem email_folder = await msgraph.FindItem(customerGroupResult.groupDrive, "General/E-Post", false);

                            //Destination folder for emails missing, create it
                            if (email_folder == null)
                            {
                                await msgraph.CreateFolder(customerGroupResult.group.Id, customerGroupResult.generalFolder.Id, "E-Post");
                                telemetryClient.TrackEvent(new EventTelemetry($"Created email folder in {dbCustomer.Name}"));
                            }
                            else
                            {
                                telemetryClient.TrackEvent(new EventTelemetry($"Found email folder in {dbCustomer.Name}"));
                            }

                            //get cdn group
                            Group cdnGroup = await graphClient.Groups[CDNTeamID].GetAsync();

                            if (cdnGroup != null)
                            {
                                //get drive for cdn group
                                Drive cdnDrive = await msgraph.GetGroupDrive(cdnGroup);

                                if (cdnDrive != null)
                                {
                                    telemetryClient.TrackEvent(new EventTelemetry($"Found CDN team and drive"));
                                    //Loop through email folders 
                                    for(int i = 0; i <= historyMonths; i++)
                                    {
                                        //get current email folder
                                        DriveItem emailFolder = await msgraph.FindItem(cdnDrive, "General/EmailMessages_" + DateTime.Now.AddMonths(-i).Month.ToString() + "_" + DateTime.Now.AddMonths(-i).Year.ToString(), false);

                                        if (emailFolder != default(DriveItem))
                                        {
                                            telemetryClient.TrackEvent(new EventTelemetry($"Found CDN email folder General/EmailMessages_" + DateTime.Now.AddMonths(-i).Month.ToString() + "_" + DateTime.Now.AddMonths(-i).Year.ToString()));
                                            OrderFiles foundOrderFiles = await GetOrderFiles(cdnDrive, emailFolder, data, customerNo, msgraph);

                                            if (foundOrderFiles.file != null)
                                            {
                                                telemetryClient.TrackEvent(new EventTelemetry($"Move order file: {foundOrderFiles.file.Name}"));
                                                //move the order file
                                                if (await msgraph.MoveFile(
                                                    new CopyItem(CDNTeamID, emailFolder.Id, foundOrderFiles.file.Name, foundOrderFiles.file.Id),
                                                    new CopyItem(customerGroupResult.group.Id, email_folder.Id, foundOrderFiles.file.Name, "")
                                                    ))
                                                {
                                                    //move corresponding files
                                                    foreach (var correspondingFile in foundOrderFiles.associated)
                                                    {
                                                        telemetryClient.TrackEvent(new EventTelemetry($"Move corresponding file: {correspondingFile.Name}"));
                                                        await msgraph.MoveFile(
                                                            new CopyItem(CDNTeamID, emailFolder.Id, correspondingFile.Name, correspondingFile.Id),
                                                            new CopyItem(customerGroupResult.group.Id, email_folder.Id, correspondingFile.Name, "")
                                                        );
                                                    }
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                        }
                        else
                        {
                            if (!string.IsNullOrEmpty(data.Sender))
                            {
                                Chat newChat = await msgraph.CreateOneOnOneChat(new List<string>
                            {
                                data.Sender,
                                config["SendFrom"]
                            });

                                if (newChat != null)
                                {
                                    ChatMessage msg = await msgraph.SendOneOnOneMessage(newChat, config["MessageNoCustomer"].Replace("<orderno>", orderNo));
                                }
                            }

                            telemetryClient.TrackTrace(new TraceTelemetry($"Unable to find customer and group for no {customerNo} "));
                        }
                    }
                    else
                    {
                        if (!string.IsNullOrEmpty(data.Sender))
                        {
                            Chat newChat = await msgraph.CreateOneOnOneChat(new List<string>
                            {
                                data.Sender,
                                config["SendFrom"]
                            });

                            if (newChat != null)
                            {
                                ChatMessage msg = await msgraph.SendOneOnOneMessage(newChat, config["MessageNoCustomer"].Replace("<orderno>", orderNo));
                            }
                        }

                        telemetryClient.TrackTrace(new TraceTelemetry($"Unable to find customer and group for no {customerNo} "));
                    }
                }
                else
                {
                    telemetryClient.TrackTrace(new TraceTelemetry($"Unable to find customer {customerNo} in CDN list"));
                }
            }
        }

        public async Task<bool> ProcessCDNFiles(string root, Graph msgraph, Drive cdnDrive, HandleEmailMessage data, string orderNo, int i, string CDNTeamID, FindOrderGroupAndFolder orderFolder)
        {
            //get current email folder
            DriveItem emailFolder = await msgraph.FindItem(cdnDrive, root + "/EmailMessages_" + DateTime.Now.AddMonths(-i).Month.ToString() + "_" + DateTime.Now.AddMonths(-i).Year.ToString(), false);

            if (emailFolder != default(DriveItem))
            {
                telemetryClient.TrackEvent(new EventTelemetry($"Found CDN email folder General/EmailMessages_" + DateTime.Now.AddMonths(-i).Month.ToString() + "_" + DateTime.Now.AddMonths(-i).Year.ToString()));
                OrderFiles foundOrderFiles = await GetOrderFiles(cdnDrive, emailFolder, data, orderNo, msgraph);

                if (foundOrderFiles.file != null)
                {
                    telemetryClient.TrackEvent(new EventTelemetry($"Move order file: {foundOrderFiles.file.Name}"));
                    //move the order file
                    if (await msgraph.MoveFile(
                        new CopyItem(CDNTeamID, emailFolder.Id, foundOrderFiles.file.Name, foundOrderFiles.file.Id),
                        new CopyItem(orderFolder.orderGroup.Id, orderFolder.orderFolder.Id, foundOrderFiles.file.Name, "")
                        ))
                    {
                        //move corresponding files
                        foreach (var correspondingFile in foundOrderFiles.associated)
                        {
                            telemetryClient.TrackEvent(new EventTelemetry($"Move corresponding file: {correspondingFile.Name}"));
                            await msgraph.MoveFile(
                                new CopyItem(CDNTeamID, emailFolder.Id, correspondingFile.Name, correspondingFile.Id),
                                new CopyItem(orderFolder.orderGroup.Id, orderFolder.orderFolder.Id, correspondingFile.Name, "")
                            );
                        }
                    }
                }
            }

            return true;
        }

        public async Task<OrderFiles> GetOrderFiles(Drive cdnDrive, DriveItem emailFolder, HandleEmailMessage data, string orderNo, Graph msgraph)
        {
            OrderFiles returnValue = new OrderFiles();
            returnValue.associated = new List<DriveItem>();
            var emailChildren = await msgraph.GetDriveFolderChildren(cdnDrive, emailFolder, false);

            if (String.IsNullOrEmpty(data.Title))
            {
                //if filename was sent get the order pdf
                foreach (var emailChild in emailChildren)
                {
                    if (emailChild.Name.Contains(orderNo))
                    {
                        returnValue.file = emailChild;
                        string fileid = RE.Regex.Match(returnValue.file.Name, @"(\d+)\.[a-z]*[A-Z]*$").Groups[1].Value;

                        //then get all corresponding files by comparing id
                        foreach (var mailfile in emailChildren)
                        {
                            string mailid = RE.Regex.Match(mailfile.Name, @"(\d+)\.[a-z]*[A-Z]*$").Groups[1].Value;

                            if (mailid == fileid && (mailfile.Id != returnValue.file.Id))
                            {
                                returnValue.associated.Add(mailfile);
                            }
                        }

                        break;
                    }
                }
            }
            else
            {
                //if only title was sent get the order email
                foreach (var emailChild in emailChildren)
                {
                    bool isOrderFile = emailChild.Name.StartsWith(orderNo + "_") && emailChild.Name.ToLower().EndsWith("pdf");

                    if (emailChild.Name.Contains(orderNo) && !isOrderFile)
                    {
                        returnValue.file = emailChild;
                        string fileid = RE.Regex.Match(returnValue.file.Name, @"(\d+)\.[a-z]*[A-Z]*$").Groups[1].Value;

                        //then get all corresponding files by comparing id
                        foreach (var mailfile in emailChildren)
                        {
                            string mailid = RE.Regex.Match(mailfile.Name, @"(\d+)\.[a-z]*[A-Z]*$").Groups[1].Value;

                            if (mailid == fileid && (mailfile.Id != returnValue.file.Id))
                            {
                                returnValue.associated.Add(mailfile);
                            }
                        }

                        break;
                    }
                }
            }

            return returnValue;
        }
    }
}
