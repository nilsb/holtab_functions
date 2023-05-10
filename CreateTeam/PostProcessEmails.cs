using System;
using System.Collections.Generic;
using Azure.Identity;
using CreateTeam.Shared;
using RE = System.Text.RegularExpressions;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Host;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;
using Microsoft.Graph;
using CreateTeam.Models;
using System.Threading;
using System.IO;
using Microsoft.ApplicationInsights;
using Microsoft.ApplicationInsights.Extensibility;
using Microsoft.ApplicationInsights.DataContracts;
using System.Threading.Tasks;
using System.Linq;
using Microsoft.Graph.Models;

namespace CreateTeam
{
    public class PostProcessEmails
    {
        private readonly TelemetryClient telemetryClient;

        public PostProcessEmails(TelemetryConfiguration telemetryConfiguration)
        {
            this.telemetryClient = new TelemetryClient(telemetryConfiguration);
        }

        [FunctionName("PostProcessEmails")]
        public async Task Run([TimerTrigger("0 */5 * * * *")] TimerInfo myTimer, [Queue("receive"), StorageAccount("AzureWebJobsStorage")] ICollector<string> outputQueueItem, Microsoft.Azure.WebJobs.ExecutionContext context, ILogger log)
        {
            log.LogInformation($"PostProcessEmail Timer trigger function executed at: {DateTime.Now}");

            var config = new ConfigurationBuilder()
              .SetBasePath(context.FunctionAppDirectory)
              .AddJsonFile("local.settings.json", optional: true, reloadOnChange: true)
              .AddEnvironmentVariables()
              .Build();

            string ClientID = config["ClientID"];
            string ClientSecret = config["ClientSecret"];
            string TenantID = config["TenantID"];
            string CDNTeamID = config["CDNTeamID"];

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
            string orderNo = "";

            Graph msgraph = new Graph(graphClient, log);
            Common common = new Common(graphClient, config, log, telemetryClient, msgraph);
            List<OrderFiles> orderFiles = new List<OrderFiles>();
            var groupDrive = await graphClient.Groups[CDNTeamID].Drive.GetAsync();
            DriveItem emailMessagesFolder = await common.GetEmailsFolder("General", DateTime.Now.Month.ToString(), DateTime.Now.Year.ToString());
            var emailMessagesChildren = await msgraph.GetDriveFolderChildren(groupDrive, emailMessagesFolder, false);
            emailMessagesFolder = await common.GetEmailsFolder("Salesemails", DateTime.Now.Month.ToString(), DateTime.Now.Year.ToString());
            var salesEmailMessagesChildren = await msgraph.GetDriveFolderChildren(groupDrive, emailMessagesFolder, false);

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
