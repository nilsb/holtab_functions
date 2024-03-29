using System;
using System.Collections.Generic;
using Shared;
using RE = System.Text.RegularExpressions;
using Microsoft.Azure.WebJobs;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;
using Shared.Models;
using System.Threading.Tasks;
using Microsoft.Graph.Models;

namespace Jobs
{
    public class PostProcessEmails
    {
        private readonly IConfiguration config;

        public PostProcessEmails(IConfiguration _config)
        {
            config = _config;
        }

        [FunctionName("PostProcessEmails")]
        public async Task Run([TimerTrigger("0 */5 * * * *")] TimerInfo myTimer, [Queue("receive"), StorageAccount("AzureWebJobsStorage")] ICollector<string> outputQueueItem, Microsoft.Azure.WebJobs.ExecutionContext context, ILogger log)
        {

            Settings settings = new Settings(config, context, log);
            bool debug = (settings?.debugFlags?.Job?.PostProcessEmails).HasValue && (settings?.debugFlags?.Job?.PostProcessEmails).Value;
            Graph msGraph = new Graph(settings);
            Common common = new Common(settings, msGraph, debug);

            if(debug)
                log.LogInformation($"Job PostProcessEmails: Timer trigger function executed at: {DateTime.Now}");

            string orderNo = "";
            List<OrderFiles> orderFiles = new List<OrderFiles>();
            string groupDriveId = await msGraph.GetGroupDrive(settings.CDNTeamID, debug);
            DriveItem emailMessagesFolder = await common.GetEmailsFolder("General", DateTime.Now.Month.ToString(), DateTime.Now.Year.ToString(), debug);
            var emailMessagesChildren = await msGraph.GetDriveFolderChildren(groupDriveId, emailMessagesFolder.Id, false, debug);
            emailMessagesFolder = await common.GetEmailsFolder("Salesemails", DateTime.Now.Month.ToString(), DateTime.Now.Year.ToString(), debug);
            var salesEmailMessagesChildren = await msGraph.GetDriveFolderChildren(groupDriveId, emailMessagesFolder.Id, false, debug);

            if(emailMessagesChildren != null && salesEmailMessagesChildren != null)
            {
                emailMessagesChildren.AddRange(salesEmailMessagesChildren);
            }
            else if(salesEmailMessagesChildren != null)
            {
                emailMessagesChildren = salesEmailMessagesChildren;
            }

            if(emailMessagesChildren != null)
            {
                foreach (var emailChild in emailMessagesChildren)
                {
                    //if (emailChild.CreatedDateTime.Value >= DateTime.Now.AddHours(-1))
                    //    continue;

                    if (emailChild.Name.ToLowerInvariant().EndsWith("pdf"))
                    {
                        orderNo = common.FindOrderNoInString(emailChild.Name);

                        if (!string.IsNullOrEmpty(orderNo) && emailChild.Name.StartsWith(orderNo))
                        {
                            if(debug)
                                log.LogInformation($"Job PostProcessEmails: Found orderno {orderNo} in filename {emailChild.Name}");

                            var order = new OrderFiles();
                            order.file = emailChild;
                            order.associated = new List<DriveItem>();
                            string fileid = RE.Regex.Match(emailChild.Name, @"(\d+)\.[a-z]*[A-Z]*$").Groups[1].Value;

                            //then get all corresponding files by comparing id
                            foreach (var mailfile in emailMessagesChildren)
                            {
                                string mailid = RE.Regex.Match(mailfile.Name, @"(\d+)\.[a-z]*[A-Z]*$").Groups[1].Value;

                                if (mailid == fileid && (mailfile.Id != emailChild.Id))
                                {
                                    order.associated.Add(mailfile);
                                }
                            }

                            if(debug)
                                log.LogInformation("Job PostProcessEmails: Putting message on handle email queue { \"filename\": \"" + emailChild.Name + "\", \"Source\": \"PostProcess\" }");

                            outputQueueItem.Add("{ \"filename\": \"" + emailChild.Name + "\", \"Source\": \"PostProcess\" }");
                            orderFiles.Add(order);
                        }
                    }
                }

                foreach (var emailChild in emailMessagesChildren)
                {
                    //if (emailChild.CreatedDateTime.Value >= DateTime.Now.AddHours(-1))
                    //    continue;

                    if (debug && !emailChild.Name.ToLowerInvariant().EndsWith("pdf") && !orderFiles.Exists(of => of.associated.Exists(ofa => ofa.Name == emailChild.Name)))
                        log.LogInformation("Putting message on handle email queue: { \"title\": \"" + emailChild.Name + "\", \"Source\": \"PostProcess\" }");

                    outputQueueItem.Add("{ \"title\": \"" + emailChild.Name + "\", \"Source\": \"PostProcess\" }");
                }
            }
        }
    }
}
