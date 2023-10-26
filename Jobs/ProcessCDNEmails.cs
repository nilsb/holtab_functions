using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Net.Mail;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using Azure.Storage.Blobs;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Host;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;
using Microsoft.Graph;
using Microsoft.Graph.Drives.Item.Items.Item.CreateUploadSession;
using Microsoft.Graph.Models;
using Shared;
using Shared.Models;
using YamlDotNet.Serialization.NodeTypeResolvers;

namespace Jobs
{
    public class ProcessCDNEmails
    {
        private readonly IConfiguration config;
        private const int ChunkSize = 320 * 1024; // This is 320 KB. Adjust based on your requirement.
        private const int pagesize = 50;

        public ProcessCDNEmails(IConfiguration config)
        {
            this.config = config;
        }

        [FunctionName("ProcessCDNEmails")]
        public async Task Run([TimerTrigger("0 */10 * * * *")] TimerInfo myTimer,
            Microsoft.Azure.WebJobs.ExecutionContext context,
            ILogger log)
        {
            Settings settings = new Settings(config, context, log);
            bool debug = true;
            Graph msGraph = new Graph(settings);
            Common common = new Common(settings, msGraph, debug);

            if (debug)
            {
                log?.LogInformation($"ProcessCDNEmails: trigger function executed at: {DateTime.Now}");
                log?.LogInformation("ProcessCDNEmails: GetCDNTeam");
            }

            string team = await msGraph.GetTeamFromGroup(settings.CDNTeamID, true);

            if (!string.IsNullOrEmpty(team))
            {
                var teamDrive = await msGraph.GetGroupDrive(team, true);
                var primaryChannel = await settings.GraphClient.Teams[team].PrimaryChannel.GetAsync();

                if (primaryChannel != null)
                {
                    if (debug)
                        log?.LogInformation("ProcessCDNEmails: Get messages in team");

                    int count = pagesize;

                    var messages = await settings.GraphClient.Teams[team].Channels[primaryChannel.Id].Messages.Delta.GetAsDeltaGetResponseAsync((requestConfiguration) =>
                    {
                        requestConfiguration.QueryParameters.Top = pagesize;
                    });

                    await ProcessMessages(messages.Value, primaryChannel, team, teamDrive, msGraph, settings, common, log, debug);

                    while (!string.IsNullOrEmpty(messages.OdataNextLink)) {
                        messages = await settings.GraphClient.Teams[team].Channels[primaryChannel.Id].Messages.Delta.GetAsDeltaGetResponseAsync((requestConfiguration) =>
                        {
                            requestConfiguration.QueryParameters.Top = pagesize;
                            requestConfiguration.QueryParameters.Skip = count;
                            count += pagesize;
                        });

                        await ProcessMessages(messages.Value, primaryChannel, team, teamDrive, msGraph, settings, common, log, debug);
                    }
                }
            }

        }

        private string ExtractSubFolderNameFromContentUrl(string contentUrl)
        {
            var match = Regex.Match(contentUrl, "/General/([^/]+)/");
            return match.Success ? match.Groups[1].Value : null;
        }

        private async Task ProcessMessages(List<ChatMessage> messages, Channel primaryChannel, string team, string teamDrive, Graph msGraph, Settings settings, Common common, ILogger log, bool debug)
        {
            foreach (var message in messages)
            {
                bool moved = false;

                if (debug)
                    log?.LogInformation("ProcessCDNEmails: " + team + ": " + message.Subject);

                var msg = await settings.GraphClient.Teams[team].Channels[primaryChannel.Id].Messages[message.Id].GetAsync();
                string orderno = common.FindOrderNoInString(msg.Subject);
                string customerno = common.FindCustomerNoInString(msg.Subject);

                if (!string.IsNullOrEmpty(orderno))
                {
                    var groupAndFolder = common.GetOrderGroupAndFolder(orderno, true);

                    if (groupAndFolder.Success)
                    {
                        moved = await ProcessAttachments(msg, primaryChannel, team, teamDrive, groupAndFolder.orderGroupId, groupAndFolder.orderFolderId, msGraph, settings, log, debug);
                    }
                }
                else if (!string.IsNullOrEmpty(customerno))
                {
                    FindCustomerResult customerResult = common.GetCustomer(customerno, "Supplier", debug);

                    if (customerResult.Success && customerResult.customers.Count > 0)
                    {
                        Customer dbCustomer = customerResult.customers.OrderByDescending(c => c.Created).Take(1).FirstOrDefault();

                        if (dbCustomer != null)
                        {
                            if (debug)
                                log?.LogInformation($"ProcessCDNEmails: Found customer {dbCustomer.Name} in CDN");

                            FindCustomerGroupResult customerGroupResult = common.FindCustomerGroupAndDrive(dbCustomer.Name, dbCustomer.ExternalId, dbCustomer.Type, debug);

                            if (customerGroupResult.Success && !string.IsNullOrEmpty(customerGroupResult.groupId))
                            {
                                if (debug)
                                    log?.LogInformation($"ProcessCDNEmails: Found customer group and drive for {dbCustomer.Name}");

                                var customerTeam = await msGraph.GetTeamFromGroup(customerGroupResult.groupId, debug);

                                if (customerTeam != null)
                                {
                                    if (debug)
                                        log?.LogInformation($"ProcessCDNEmails: Found customer team for {dbCustomer.Name}");

                                    var customerPrimaryChannel = await settings.GraphClient.Teams[customerTeam].PrimaryChannel.GetAsync();

                                    if (customerPrimaryChannel != null)
                                    {
                                        var customerPrimaryChannelFolder = await settings.GraphClient.Teams[customerTeam].Channels[customerPrimaryChannel.Id].FilesFolder.GetAsync();

                                        if (customerPrimaryChannelFolder != null)
                                        {
                                            if (debug)
                                                log?.LogInformation($"ProcessCDNEmails: Found primary channel in team for {dbCustomer.Name}");

                                            var emailsfolder = await msGraph.FindItem(customerGroupResult.groupDriveId, customerPrimaryChannelFolder.Id, "E-Post", false, debug);

                                            if (emailsfolder != null)
                                            {
                                                if (debug)
                                                    log?.LogInformation($"ProcessCDNEmails: Found primary channel in team for {dbCustomer.Name}");

                                                moved = await ProcessAttachments(msg, primaryChannel, team, teamDrive, customerGroupResult.groupId, emailsfolder.Id, msGraph, settings, log, debug);
                                            }
                                        }
                                    }

                                }
                            }
                        }
                    }
                }

                if (moved)
                {
                    try
                    {
                        await settings.GraphClient.Teams[team].Channels[primaryChannel.Id].Messages[msg.Id].SoftDelete.PostAsync();
                    }
                    catch (Exception)
                    {
                    }
                }
            }

        }

        private async Task<bool> ProcessAttachments(ChatMessage msg, Channel primaryChannel, string team, string teamDrive, string destinationGroup, string destinationFolder, Graph msGraph, Settings settings, ILogger log, bool debug)
        {
            bool returnValue = true;

            var attachments = msg.Attachments;

            if (debug)
                log?.LogInformation($"ProcessCDNEmails: Processing attachments");

            foreach (var attachment in attachments)
            {
                var contentUrl = attachment.ContentUrl;

                if (debug)
                    log?.LogInformation($"ProcessCDNEmails: Attachment content URL {contentUrl}");

                var subfolder = ExtractSubFolderNameFromContentUrl(contentUrl);

                if (debug)
                    log?.LogInformation($"ProcessCDNEmails: Extracted subfolder name {subfolder}");

                if (!string.IsNullOrEmpty(contentUrl) && !string.IsNullOrEmpty(subfolder))
                {
                    var primaryChannelFolder = await settings.GraphClient.Teams[team].Channels[primaryChannel.Id].FilesFolder.GetAsync();

                    if(primaryChannelFolder != null)
                    {
                        if (debug)
                            log?.LogInformation($"ProcessCDNEmails: Trying to find folder {primaryChannelFolder.Name}/{subfolder}");

                        try
                        {
                            var folder = await msGraph.FindItem(teamDrive, primaryChannelFolder.Id, subfolder, false, debug);

                            if (folder != null)
                            {
                                var searchFile = await msGraph.FindItem(teamDrive, folder.Id, attachment.Name, false, debug);

                                if (searchFile != null)
                                {
                                    if (debug)
                                        log?.LogInformation($"ProcessCDNEmails: Trying to download item {primaryChannelFolder.Name}/{subfolder}/{attachment.Name}");

                                    var file = await msGraph.DownloadFile(team, folder.Id, attachment.Name, debug);

                                    if (file != null && file.Contents != Stream.Null && file.Contents.Length > 0)
                                    {
                                        bool uploadResult = await msGraph.UploadFile(destinationGroup, destinationFolder, attachment.Name, file.Contents, debug);

                                        if (uploadResult)
                                        {
                                            if (debug)
                                                log?.LogInformation($"ProcessCDNEmails: Uploaded file {attachment.Name} to destination");

                                            returnValue &= true;
                                        }
                                        else
                                        {
                                            if (debug)
                                                log?.LogError($"ProcessCDNEmails: Failed to upload {attachment.Name}");

                                            returnValue &= false;
                                        }
                                    }
                                    else
                                    {
                                        if (debug)
                                            log?.LogError($"ProcessCDNEmails: Failed to download {primaryChannelFolder.Name}/{subfolder}/{attachment.Name}");

                                        returnValue &= false;
                                    }
                                }
                            }
                            else
                            {
                                returnValue &= false;
                            }
                        }
                        catch (Exception ex)
                        {
                            log?.LogError($"ProcessCDNEmail: Failed to copy file {attachment.Name} to team {team} ");
                        }
                    }
                    else
                    {
                        returnValue &= false;
                    }
                }
                else 
                { 
                    returnValue &= false; 
                }
            }

            return returnValue;
        }
    }
}
