using System;
using System.Collections.Generic;
using System.IO;
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

namespace Jobs
{
    public class ProcessCDNEmails
    {
        private readonly IConfiguration config;
        private const int ChunkSize = 320 * 1024; // This is 320 KB. Adjust based on your requirement.

        public ProcessCDNEmails(IConfiguration config)
        {
            this.config = config;
        }

        [FunctionName("ProcessCDNEmails")]
        public async Task Run([TimerTrigger("0 */30 * * * *")]TimerInfo myTimer,
            Microsoft.Azure.WebJobs.ExecutionContext context,
            ILogger log)
        {
            log.LogInformation($"ProcessCDNEmails trigger function executed at: {DateTime.Now}");
            Settings settings = new Settings(config, context, log);
            bool debug = (settings?.debugFlags?.Customer?.BGCustomerInfo).HasValue && (settings?.debugFlags?.Customer?.BGCustomerInfo).Value;
            Graph msGraph = new Graph(settings);
            Common common = new Common(settings, msGraph, debug);

            log?.LogInformation("GetCDNTeam");
            string team = await msGraph.GetTeamFromGroup(settings.CDNTeamID, true);

            if (!string.IsNullOrEmpty(team))
            {
                var primaryChannel = await settings.GraphClient.Teams[team].PrimaryChannel.GetAsync();

                if(primaryChannel != null)
                {
                    log?.LogInformation("Get messages in team");
                    var messages = await settings.GraphClient.Teams[team].Channels[primaryChannel.Id].Messages.GetAsync((requestConfiguration) =>
                    {
                        requestConfiguration.QueryParameters.Top = 50;
                    });

                    foreach (var message in messages?.Value)
                    {
                        bool moved = false;
                        log?.LogInformation(team + ": " + message);

                        var msg = await settings.GraphClient.Teams[team].Channels[primaryChannel.Id].Messages[message.Id].GetAsync();
                        string orderno = common.FindOrderNoInString(msg.Subject);

                        if (!string.IsNullOrEmpty(orderno))
                        {
                            var groupAndFolder = common.GetOrderGroupAndFolder(orderno, true);

                            if (groupAndFolder.Success)
                            {
                                //var emailStream = new MemoryStream(Encoding.UTF8.GetBytes(msg.Body.Content));
                                //bool uploadMsgResult = await msGraph.UploadFile(groupAndFolder.orderGroupId, groupAndFolder.orderFolder.Id, $"{msg.Id}.txt", emailStream, true);

                                //if (!uploadMsgResult)
                                //{
                                //    log?.LogError($"unable to upload file {msg.Subject}.txt");
                                //    moved = false;
                                //}
                                //else
                                //{
                                //    moved = true;
                                //}

                                var attachments = msg.Attachments;

                                foreach (var attachment in attachments)
                                {
                                    var contentUrl = attachment.ContentUrl;
                                    var subfolder = ExtractSubFolderNameFromContentUrl(contentUrl);

                                    if (!string.IsNullOrEmpty(contentUrl) && !string.IsNullOrEmpty(subfolder))
                                    {
                                        var folder = await msGraph.FindItem(team, subfolder, false, true);

                                        if(folder != null)
                                        {
                                            var file = await msGraph.DownloadFile(team, folder.Id, attachment.Name, true);
                                            bool uploadResult = await msGraph.UploadFile(groupAndFolder.orderGroupId, groupAndFolder.orderFolder.Id, attachment.Name, file.Contents, true);

                                            if (uploadResult)
                                            {
                                                log?.LogInformation($"Uploaded file {attachment.Name} to group for order {orderno}");
                                                moved &= true;
                                            }
                                            else
                                            {
                                                log?.LogError($"Failed to upload {attachment.Name}");
                                                moved &= false;
                                            }
                                        }
                                    }
                                }
                            }
                        }

                        if (moved)
                        {
                            await settings.GraphClient.Teams[team].Channels[primaryChannel.Id].Messages[msg.Id].SoftDelete.PostAsync();
                        }
                    }

                }
            }

        }

        private string ExtractSubFolderNameFromContentUrl(string contentUrl)
        {
            var match = Regex.Match(contentUrl, "/General/([^/]+)/");
            return match.Success ? match.Groups[1].Value : null;
        }

    }
}
