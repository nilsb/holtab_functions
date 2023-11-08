using System;
using System.Threading.Tasks;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Host;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;
using Shared;
using Shared.Models;

namespace Jobs
{
    public class ProcessCDNHistory
    {
        private readonly IConfiguration config;
        private const int ChunkSize = 320 * 1024; // This is 320 KB. Adjust based on your requirement.
        private const int pagesize = 50;

        public ProcessCDNHistory(IConfiguration config)
        {
            this.config = config;
        }

        [FunctionName("ProcessCDNHistory")]
        public async Task RunAsync([TimerTrigger("0 */30 * * * *")]TimerInfo myTimer,
            Microsoft.Azure.WebJobs.ExecutionContext context,
            ILogger log)
        {
            Settings settings = new Settings(config, context, log);
            bool debug = true;
            Graph msGraph = new Graph(settings);
            Common common = new Common(settings, msGraph, debug);

            if (debug)
            {
                log?.LogInformation($"ProcessCDNHistory: trigger function executed at: {DateTime.Now}");
                log?.LogInformation("ProcessCDNHistory: GetCDNTeam");
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

                    int count = 0;

                    try
                    {
                        var dbSetting = common.GetSettingFromDB("MessageSkip", debug);
                        
                        if(dbSetting != null && !string.IsNullOrEmpty(dbSetting.Value))
                        {
                            int.TryParse(dbSetting.Value, out count);
                        }
                    }
                    catch (Exception)
                    {
                    }

                    var messages = await settings.GraphClient.Teams[team].Channels[primaryChannel.Id].Messages.GetAsync((requestConfiguration) =>
                    {
                        requestConfiguration.QueryParameters.Top = pagesize;
                        requestConfiguration.QueryParameters.Skip = count;
                    });

                    await common.ProcessMessages(messages.Value, primaryChannel, team, teamDrive, msGraph, settings, common, log, debug);

                    while (!string.IsNullOrEmpty(messages.OdataNextLink))
                    {
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
