using System;
using System.Threading.Tasks;
using RE = System.Text.RegularExpressions;
using Microsoft.Azure.WebJobs;
using Microsoft.Extensions.Logging;
using Newtonsoft.Json;
using Microsoft.Extensions.Configuration;
using System.Collections.Generic;
using Shared.Models;
using Shared;
using System.Linq;
using Microsoft.Graph.Models;
using Newtonsoft.Json.Linq;

namespace Jobs
{
    public class HandleEmail
    {
        private readonly IConfiguration config;

        public HandleEmail(IConfiguration _config)
        {
            config = _config;
        }

        [FunctionName("HandleEmail")]
        public async Task Run(
            [QueueTrigger("receive", Connection = "AzureWebJobsStorage")] string myQueueItem,
            Microsoft.Azure.WebJobs.ExecutionContext context,
            ILogger log)
        {
            log.LogInformation("Got handle email request with message " + myQueueItem);
            HandleEmailMessage data = JsonConvert.DeserializeObject<HandleEmailMessage>(myQueueItem);

            Settings settings = new Settings(config, context, log);
            Graph msGraph = new Graph(settings);
            Common common = new Common(settings, msGraph);
            string HistoryMonths = config["HistoryMonths"];
            int historyMonths = 0;
            Int32.TryParse(HistoryMonths, out historyMonths);
            string orderNo;
            string customerNo = "";

            if (string.IsNullOrEmpty(data.Title))
            {
                orderNo = common.FindOrderNoInString(data.Filename);

                if (!string.IsNullOrEmpty(orderNo))
                {
                    log?.LogInformation($"Found orderno: {orderNo} in filename: {data.Filename}");
                }
                else
                {
                    customerNo = common.FindCustomerNoInString(data.Filename);

                    if (!string.IsNullOrEmpty(customerNo))
                    {
                        log?.LogInformation($"Found customer no: {customerNo} in filename: {data.Filename}");
                    }
                }
            }
            else
            {
                orderNo = common.FindOrderNoInString(data.Title);

                if (!string.IsNullOrEmpty(orderNo))
                {
                    log?.LogInformation($"Found orderno: {orderNo} in email subject: {data.Title}");
                }
                else
                {
                    customerNo = common.FindCustomerNoInString(data.Title);

                    if (!string.IsNullOrEmpty(customerNo))
                    {
                        log?.LogInformation($"Found customer no: {customerNo} in subject: {data.Title}");
                    }
                }
            }

            //handle order related emails
            if (!string.IsNullOrEmpty(orderNo))
            {
                FindOrderGroupAndFolder orderFolder = common.GetOrderGroupAndFolder(orderNo);

                if (orderFolder.Success && orderFolder.orderFolder != null)
                {
                    log?.LogInformation($"Found order group: {orderFolder.orderTeamId} for order no: {orderNo}");

                    //get drive for cdn group
                    string cdnDriveId = await msGraph.GetGroupDrive(settings.CDNTeamID);

                    if (!string.IsNullOrEmpty(cdnDriveId))
                    {
                        log?.LogInformation($"Found CDN team and drive");
                            
                        //Loop through email folders
                        for(int i = 0; i <= historyMonths; i++)
                        {
                            await ProcessCDNFiles("General", msGraph, cdnDriveId, data, orderNo, i, settings.CDNTeamID, orderFolder, log);
                            await ProcessCDNFiles("Salesemails", msGraph, cdnDriveId, data, orderNo, i, settings.CDNTeamID, orderFolder, log);
                        }
                    }
                }
                else
                {
                    if (!orderFolder.Success)
                    {
                        if (!string.IsNullOrEmpty(data.Sender))
                        {
                            Chat newChat = await msGraph.CreateOneOnOneChat(new List<string>
                            {
                                data.Sender,
                                config["SendFrom"]
                            });

                            if (newChat != null)
                            {
                                ChatMessage msg = await msGraph.SendOneOnOneMessage(newChat, config["MessageNoCustomer"].Replace("<orderno>", orderNo));
                            }
                        }

                        log?.LogError($"Unable to find order group for {orderNo}");
                    }
                    else if (orderFolder.orderFolder == null)
                    {
                        if (!string.IsNullOrEmpty(data.Sender))
                        {
                            Chat newChat = await msGraph.CreateOneOnOneChat(new List<string>
                            {
                                data.Sender,
                                config["SendFrom"]
                            });

                            if (newChat != null)
                            {
                                ChatMessage msg = await msGraph.SendOneOnOneMessage(newChat, config["MessageNoOrderFolder"].Replace("<orderno>", orderNo));
                            }
                        }

                        log?.LogError($"Unable to find order folder for {orderNo} in group {orderFolder.orderGroupId}");
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
                        log?.LogInformation($"Found customer {dbCustomer.Name} in CDN");
                        FindCustomerGroupResult customerGroupResult = common.FindCustomerGroupAndDrive(dbCustomer.Name, dbCustomer.ExternalId, dbCustomer.Type);

                        if (customerGroupResult.Success && !string.IsNullOrEmpty(customerGroupResult.groupId))
                        {
                            log?.LogInformation($"Found customer group and drive for {dbCustomer.Name}");
                            //find email destination folder
                            DriveItem email_folder = await msGraph.FindItem(customerGroupResult.groupDriveId, "General/E-Post", false);

                            //Destination folder for emails missing, create it
                            if (email_folder == null)
                            {
                                await msGraph.CreateFolder(customerGroupResult.groupId, customerGroupResult.generalFolder.Id, "E-Post");
                                log?.LogInformation($"Created email folder in {dbCustomer.Name}");
                            }
                            else
                            {
                                log?.LogInformation($"Found email folder in {dbCustomer.Name}");
                            }

                            //get drive for cdn group
                            string cdnDriveId = await msGraph.GetGroupDrive(settings.CDNTeamID);

                            if (!string.IsNullOrEmpty(cdnDriveId))
                            {
                                log?.LogInformation($"Found CDN team and drive");
                                //Loop through email folders 
                                for(int i = 0; i <= historyMonths; i++)
                                {
                                    //get current email folder
                                    DriveItem emailFolder = await msGraph.FindItem(cdnDriveId, "General/EmailMessages_" + DateTime.Now.AddMonths(-i).Month.ToString() + "_" + DateTime.Now.AddMonths(-i).Year.ToString(), false);

                                    if (emailFolder != default(DriveItem))
                                    {
                                        log?.LogInformation($"Found CDN email folder General/EmailMessages_" + DateTime.Now.AddMonths(-i).Month.ToString() + "_" + DateTime.Now.AddMonths(-i).Year.ToString());
                                        OrderFiles foundOrderFiles = await GetOrderFiles(cdnDriveId, emailFolder, data, customerNo, msGraph);

                                        if (foundOrderFiles.file != null)
                                        {
                                            log?.LogInformation($"Move order file: {foundOrderFiles.file.Name}");
                                            //move the order file
                                            if (await msGraph.MoveFile(
                                                new CopyItem(settings.CDNTeamID, emailFolder.Id, foundOrderFiles.file.Name, foundOrderFiles.file.Id),
                                                new CopyItem(customerGroupResult.groupId, email_folder.Id, foundOrderFiles.file.Name, "")
                                                ))
                                            {
                                                //move corresponding files
                                                foreach (var correspondingFile in foundOrderFiles.associated)
                                                {
                                                    log?.LogInformation($"Move corresponding file: {correspondingFile.Name}");
                                                    await msGraph.MoveFile(
                                                        new CopyItem(settings.CDNTeamID, emailFolder.Id, correspondingFile.Name, correspondingFile.Id),
                                                        new CopyItem(customerGroupResult.groupId, email_folder.Id, correspondingFile.Name, "")
                                                    );
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
                                Chat newChat = await msGraph.CreateOneOnOneChat(new List<string>
                            {
                                data.Sender,
                                config["SendFrom"]
                            });

                                if (newChat != null)
                                {
                                    ChatMessage msg = await msGraph.SendOneOnOneMessage(newChat, config["MessageNoCustomer"].Replace("<orderno>", orderNo));
                                }
                            }

                            log?.LogError($"Unable to find customer and group for no {customerNo} ");
                        }
                    }
                    else
                    {
                        if (!string.IsNullOrEmpty(data.Sender))
                        {
                            Chat newChat = await msGraph.CreateOneOnOneChat(new List<string>
                            {
                                data.Sender,
                                config["SendFrom"]
                            });

                            if (newChat != null)
                            {
                                ChatMessage msg = await msGraph.SendOneOnOneMessage(newChat, config["MessageNoCustomer"].Replace("<orderno>", orderNo));
                            }
                        }

                        log?.LogError($"Unable to find customer and group for no {customerNo} ");
                    }
                }
                else
                {
                    log?.LogError($"Unable to find customer {customerNo} in CDN list");
                }
            }
        }

        public async Task<bool> ProcessCDNFiles(string root, Graph msgraph, string cdnDriveId, HandleEmailMessage data, string orderNo, int i, string CDNTeamID, FindOrderGroupAndFolder orderFolder, ILogger log)
        {
            //get current email folder
            DriveItem emailFolder = await msgraph.FindItem(cdnDriveId, root + "/EmailMessages_" + DateTime.Now.AddMonths(-i).Month.ToString() + "_" + DateTime.Now.AddMonths(-i).Year.ToString(), false);

            if (emailFolder != default(DriveItem))
            {
                log?.LogInformation($"Found CDN email folder General/EmailMessages_" + DateTime.Now.AddMonths(-i).Month.ToString() + "_" + DateTime.Now.AddMonths(-i).Year.ToString());
                OrderFiles foundOrderFiles = await GetOrderFiles(cdnDriveId, emailFolder, data, orderNo, msgraph);

                if (foundOrderFiles.file != null)
                {
                    log?.LogInformation($"Move order file: {foundOrderFiles.file.Name}");

                    //move the order file
                    if (await msgraph.MoveFile(
                        new CopyItem(CDNTeamID, emailFolder.Id, foundOrderFiles.file.Name, foundOrderFiles.file.Id),
                        new CopyItem(orderFolder.orderGroupId, orderFolder.orderFolder.Id, foundOrderFiles.file.Name, "")
                        ))
                    {
                        //move corresponding files
                        foreach (var correspondingFile in foundOrderFiles.associated)
                        {
                            log?.LogInformation($"Move corresponding file: {correspondingFile.Name}");
                            await msgraph.MoveFile(
                                new CopyItem(CDNTeamID, emailFolder.Id, correspondingFile.Name, correspondingFile.Id),
                                new CopyItem(orderFolder.orderGroupId, orderFolder.orderFolder.Id, correspondingFile.Name, "")
                            );
                        }
                    }
                }
            }

            return true;
        }

        public async Task<OrderFiles> GetOrderFiles(string cdnDriveId, DriveItem emailFolder, HandleEmailMessage data, string orderNo, Graph msgraph)
        {
            OrderFiles returnValue = new OrderFiles();
            returnValue.associated = new List<DriveItem>();
            var emailChildren = await msgraph.GetDriveFolderChildren(cdnDriveId, emailFolder.Id, false);

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
