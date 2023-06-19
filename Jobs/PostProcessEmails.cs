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
            log.LogInformation($"PostProcessEmail Timer trigger function executed at: {DateTime.Now}");

            Settings settings = new Settings(config, context, log);
            Graph msGraph = new Graph(settings);
            Common common = new Common(settings, msGraph);
            string orderNo = "";
            List<OrderFiles> orderFiles = new List<OrderFiles>();
            var groupDrive = await msGraph.GetGroupDrive(settings.CDNTeamID);
            DriveItem emailMessagesFolder = await common.GetEmailsFolder("General", DateTime.Now.Month.ToString(), DateTime.Now.Year.ToString());
            var emailMessagesChildren = await msGraph.GetDriveFolderChildren(groupDrive, emailMessagesFolder, false);
            emailMessagesFolder = await common.GetEmailsFolder("Salesemails", DateTime.Now.Month.ToString(), DateTime.Now.Year.ToString());
            var salesEmailMessagesChildren = await msGraph.GetDriveFolderChildren(groupDrive, emailMessagesFolder, false);

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
                            log.LogInformation($"Found orderno: {orderNo} in filename: {emailChild.Name}");
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

                            log.LogInformation("Putting message on handle email queue: { \"filename\": \"" + emailChild.Name + "\", \"Source\": \"PostProcess\" }");
                            outputQueueItem.Add("{ \"filename\": \"" + emailChild.Name + "\", \"Source\": \"PostProcess\" }");
                            orderFiles.Add(order);
                        }
                    }
                }

                foreach (var emailChild in emailMessagesChildren)
                {
                    //if (emailChild.CreatedDateTime.Value >= DateTime.Now.AddHours(-1))
                    //    continue;

                    if (!emailChild.Name.ToLowerInvariant().EndsWith("pdf") && !orderFiles.Exists(of => of.associated.Exists(ofa => ofa.Name == emailChild.Name)))
                        log.LogInformation("Putting message on handle email queue: { \"title\": \"" + emailChild.Name + "\", \"Source\": \"PostProcess\" }");
                    outputQueueItem.Add("{ \"title\": \"" + emailChild.Name + "\", \"Source\": \"PostProcess\" }");
                }
            }
        }
    }
}
