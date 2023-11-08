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
        private readonly string inkopid = "da469009-c460-4369-92fc-3c3da320c7fe";
        private readonly IConfiguration config;
        private const int ChunkSize = 320 * 1024; // This is 320 KB. Adjust based on your requirement.
        private const int pagesize = 50;
        private string inkopDriveId = "b!Mtdmyl658UqrleOgpLyHOOkJJoILMYlAqRvB302xJFf6fnQUOvH3TK_tuPTNyV4E";
        private string inkopFolderId = "01TTN2ZDN3PPBVAPKDIVGLEED4YGHPSOIN";

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
            bool debug = false;
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

                    var messages = await settings.GraphClient.Teams[team].Channels[primaryChannel.Id].Messages.GetAsync((requestConfiguration) =>
                    {
                        requestConfiguration.QueryParameters.Top = pagesize;
                    });

                    await common.ProcessMessages(messages.Value, primaryChannel, team, teamDrive, msGraph, settings, common, log, debug);

                    while (!string.IsNullOrEmpty(messages.OdataNextLink) && count <= 400) {
                        count += pagesize;

                        messages = await settings.GraphClient.Teams[team].Channels[primaryChannel.Id].Messages.GetAsync((requestConfiguration) =>
                        {
                            requestConfiguration.QueryParameters.Top = pagesize;
                            requestConfiguration.QueryParameters.Skip = count;
                        });

                        await common.ProcessMessages(messages.Value, primaryChannel, team, teamDrive, msGraph, settings, common, log, debug);
                    }
                }
            }
        }

    }
}
