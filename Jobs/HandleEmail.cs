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
            bool debug = (settings?.debugFlags?.Job?.PostProcessEmails).HasValue && (settings?.debugFlags?.Job?.PostProcessEmails).Value;
            Graph msGraph = new Graph(settings);
            Common common = new Common(settings, msGraph, debug);

            string HistoryMonths = config["HistoryMonths"];
            int historyMonths = 0;
            Int32.TryParse(HistoryMonths, out historyMonths);
            string orderNo = "";
            string customerNo = "";

            //use filename in message
            if (!string.IsNullOrEmpty(data.Filename))
            {
                orderNo = common.FindOrderNoInString(data.Filename);

                if (!string.IsNullOrEmpty(orderNo) && debug)
                {
                    log?.LogInformation($"Job HandelEmail: Found orderno {orderNo} in filename {data.Filename}");
                }
                else
                {
                    customerNo = common.FindCustomerNoInString(data.Filename);

                    if (!string.IsNullOrEmpty(customerNo) && debug)
                    {
                        log?.LogInformation($"Job HandleEmail: Found customerno {customerNo} in filename {data.Filename}");
                    }
                }
            }
            else if(!string.IsNullOrEmpty(data.Title)) 
            {
                //use title in message
                orderNo = common.FindOrderNoInString(data.Title);

                if (!string.IsNullOrEmpty(orderNo) && debug)
                {
                    log?.LogInformation($"Job HandleEmail: Found orderno {orderNo} in email subject {data.Title}");
                }
                else
                {
                    customerNo = common.FindCustomerNoInString(data.Title);

                    if (!string.IsNullOrEmpty(customerNo) && debug)
                    {
                        log?.LogInformation($"Job HandleEmail: Found customerno {customerNo} in subject {data.Title}");
                    }
                }
            }

            //handle order related emails
            if (!string.IsNullOrEmpty(orderNo))
            {
                FindOrderGroupAndFolder orderFolder = common.GetOrderGroupAndFolder(orderNo, debug);

                if (orderFolder.Success && orderFolder.orderFolderId != null)
                {
                    if(debug)
                        log?.LogInformation($"Job HandleEmail: Found order group: {orderFolder.orderTeamId} for order no: {orderNo}");

                    //get drive for cdn group
                    string cdnDriveId = await msGraph.GetGroupDrive(settings.CDNTeamID, debug);

                    if (!string.IsNullOrEmpty(cdnDriveId))
                    {
                        if(debug)
                            log?.LogInformation($"Job HandleEmail: Found CDN team and drive");
                            
                        //Loop through email folders
                        for(int i = 0; i <= historyMonths; i++)
                        {
                            await ProcessCDNFiles("General", msGraph, cdnDriveId, data, orderNo, i, settings.CDNTeamID, orderFolder, log, settings, debug);
                            await ProcessCDNFiles("Salesemails", msGraph, cdnDriveId, data, orderNo, i, settings.CDNTeamID, orderFolder, log, settings, debug);
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

                        if(debug)
                            log?.LogError($"Job HandleEmail: Unable to find order group for {orderNo}");
                    }
                    else if (orderFolder.orderFolderId == null)
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

                        if(debug)
                            log?.LogError($"Job HandleEmail: Unable to find order folder for {orderNo} in group {orderFolder.orderGroupId}");
                    }
                }
            }

            if (!string.IsNullOrEmpty(customerNo))
            {
                FindCustomerResult customerResult = common.GetCustomer(customerNo, "Supplier", debug);

                if (customerResult.Success && customerResult.customers.Count > 0)
                {
                    Customer dbCustomer = customerResult.customers.OrderByDescending(c => c.Created).Take(1).FirstOrDefault();

                    if(dbCustomer != null)
                    {
                        if(debug)
                            log?.LogInformation($"Job HandleEmail: Found customer {dbCustomer.Name} in CDN");

                        FindCustomerGroupResult customerGroupResult = common.FindCustomerGroupAndDrive(dbCustomer.Name, dbCustomer.ExternalId, dbCustomer.Type, debug);

                        if (customerGroupResult.Success && !string.IsNullOrEmpty(customerGroupResult.groupId))
                        {
                            if(debug)
                                log?.LogInformation($"Job HandleEmail: Found customer group and drive for {dbCustomer.Name}");

                            //find email destination folder
                            DriveItem email_folder = await msGraph.FindItem(customerGroupResult.groupDriveId, "General/E-Post", false, debug);

                            //Destination folder for emails missing, create it
                            if (email_folder == null)
                            {
                                await msGraph.CreateFolder(customerGroupResult.groupId, customerGroupResult.generalFolderId, "E-Post", debug);

                                if(debug)
                                    log?.LogInformation($"Job HandleEmail: Created email folder in {dbCustomer.Name}");
                            }
                            else if(debug)
                            {
                                log?.LogInformation($"Job HandleEmail: Found email folder in {dbCustomer.Name}");
                            }

                            //get drive for cdn group
                            string cdnDriveId = await msGraph.GetGroupDrive(settings.CDNTeamID, debug);

                            if (!string.IsNullOrEmpty(cdnDriveId))
                            {
                                if(debug)
                                    log?.LogInformation($"Job HandleEmail: Found CDN team and drive");

                                //Loop through email folders 
                                for(int i = 0; i <= historyMonths; i++)
                                {
                                    //get current email folder
                                    DriveItem emailFolder = await msGraph.FindItem(cdnDriveId, "General/EmailMessages_" + DateTime.Now.AddMonths(-i).Month.ToString() + "_" + DateTime.Now.AddMonths(-i).Year.ToString(), false, debug);

                                    if (emailFolder != default(DriveItem))
                                    {
                                        if(debug)
                                            log?.LogInformation($"Job HandleEmail: Found CDN email folder General/EmailMessages_" + DateTime.Now.AddMonths(-i).Month.ToString() + "_" + DateTime.Now.AddMonths(-i).Year.ToString());

                                        OrderFiles foundOrderFiles = await GetOrderFiles(cdnDriveId, emailFolder, data, customerNo, msGraph, settings, debug);

                                        if (foundOrderFiles.file != null)
                                        {
                                            if(debug)
                                                log?.LogInformation($"Job HandleEmail: Move order file {foundOrderFiles.file.Name}");

                                            //move the order file
                                            if (await msGraph.MoveFile(
                                                new CopyItem(settings.CDNTeamID, emailFolder.Id, foundOrderFiles.file.Name, foundOrderFiles.file.Id),
                                                new CopyItem(customerGroupResult.groupId, email_folder.Id, foundOrderFiles.file.Name, ""),
                                                debug
                                                ))
                                            {
                                                //move corresponding files
                                                foreach (var correspondingFile in foundOrderFiles.associated)
                                                {
                                                    if(debug)
                                                        log?.LogInformation($"Job HandleEmail: Move corresponding file: {correspondingFile.Name}");

                                                    await msGraph.MoveFile(
                                                        new CopyItem(settings.CDNTeamID, emailFolder.Id, correspondingFile.Name, correspondingFile.Id),
                                                        new CopyItem(customerGroupResult.groupId, email_folder.Id, correspondingFile.Name, ""),
                                                        debug
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

                            if(debug)
                                log?.LogError($"Job HandleEmail: Unable to find customer and group for no {customerNo} ");
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

                        if(debug)
                            log?.LogError($"Job HandleEmail: Unable to find customer and group for no {customerNo} ");
                    }
                }
                else
                {
                    if(debug)
                        log?.LogError($"Job HandleEmail: Unable to find customer {customerNo} in CDN list");
                }
            }
        }

        public async Task<bool> ProcessCDNFiles(string root, Graph msgraph, string cdnDriveId, HandleEmailMessage data, string orderNo, int i, string CDNTeamID, FindOrderGroupAndFolder orderFolder, ILogger log, Settings settings, bool debug)
        {
            //get current email folder
            DriveItem emailFolder = await msgraph.FindItem(cdnDriveId, root + "/EmailMessages_" + DateTime.Now.AddMonths(-i).Month.ToString() + "_" + DateTime.Now.AddMonths(-i).Year.ToString(), false, debug);

            if (emailFolder != default(DriveItem))
            {
                if(debug)
                    log?.LogInformation($"Job HandleEmail: Found CDN email folder General/EmailMessages_" + DateTime.Now.AddMonths(-i).Month.ToString() + "_" + DateTime.Now.AddMonths(-i).Year.ToString());

                OrderFiles foundOrderFiles = await GetOrderFiles(cdnDriveId, emailFolder, data, orderNo, msgraph, settings, debug);

                if (foundOrderFiles.file != null)
                {
                    if(debug)
                        log?.LogInformation($"Job HandleEmail: Move order file {foundOrderFiles.file.Name}");

                    //move the order file
                    if (await msgraph.MoveFile(
                        new CopyItem(CDNTeamID, emailFolder.Id, foundOrderFiles.file.Name, foundOrderFiles.file.Id),
                        new CopyItem(orderFolder.orderGroupId, orderFolder.orderFolderId, foundOrderFiles.file.Name, ""),
                        debug
                        ))
                    {
                        //move corresponding files
                        foreach (var correspondingFile in foundOrderFiles.associated)
                        {
                            if(debug)
                                log?.LogInformation($"Job HandleEmail: Move corresponding file {correspondingFile.Name}");

                            await msgraph.MoveFile(
                                new CopyItem(CDNTeamID, emailFolder.Id, correspondingFile.Name, correspondingFile.Id),
                                new CopyItem(orderFolder.orderGroupId, orderFolder.orderFolderId, correspondingFile.Name, ""),
                                debug
                            );
                        }
                    }
                }
            }

            return true;
        }

        public async Task<OrderFiles> GetOrderFiles(string cdnDriveId, DriveItem emailFolder, HandleEmailMessage data, string orderNo, Graph msgraph, Settings settings, bool debug)
        {
            OrderFiles returnValue = new OrderFiles();
            returnValue.associated = new List<DriveItem>();

            var emailChildren = await msgraph.GetDriveFolderChildren(cdnDriveId, emailFolder.Id, false, debug, true);

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
