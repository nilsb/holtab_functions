using System;
using System.Threading.Tasks;
using Azure.Storage.Blobs;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Host;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;
using Shared;
using Shared.Models;

namespace Jobs
{
    public class ProcessCDNEmails
    {
        private readonly IConfiguration config;

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

            string team = await msGraph.GetTeamFromGroup(settings.CDNTeamID, true);

            if (!string.IsNullOrEmpty(team))
            {
                var messages = await settings.GraphClient.Teams[team].PrimaryChannel.Messages.GetAsync();

                foreach(var message in messages?.Value)
                {
                    var msg = await settings.GraphClient.Teams[team].PrimaryChannel.Messages[message.Id].GetAsync();
                    var attachments = msg.Attachments;

                    foreach(var attachment in attachments)
                    {
                        string orderno = common.FindOrderNoInString(msg.Subject);

                        if(!string.IsNullOrEmpty(orderno))
                        {
                            var groupAndFolder = common.GetOrderGroupAndFolder(orderno, true);

                            if (groupAndFolder.Success)
                            {
                            }
                        }
                    }
                }
            }

        }
    }
}
