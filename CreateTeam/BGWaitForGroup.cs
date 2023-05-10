using System;
using System.Threading;
using System.Threading.Tasks;
using AutoMapper;
using Azure.Messaging.ServiceBus;
using CreateTeam.Models;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Host;
using Microsoft.Azure.WebJobs.ServiceBus;
using Microsoft.Extensions.Logging;
using Newtonsoft.Json;

namespace CreateTeam
{
    public class BGWaitForGroup
    {
        [FunctionName("BGWaitForGroup")]
        public async Task Run([ServiceBusTrigger("waitforgroup", Connection = "sbholtabnavConnection")] ServiceBusReceivedMessage Message,
            ServiceBusMessageActions messageActions,
            Microsoft.Azure.WebJobs.ExecutionContext context,
            [ServiceBus("creategroup", Connection = "sbholtabnavConnection")] IAsyncCollector<dynamic> output,
            ILogger log)
        {
            log.LogInformation($"Wait for group queue trigger function processed message: {Message.Body}");

            //Parse the incoming message into JSON
            CustomerQueueMessage customerQueueMessage = Message.Body.ToObjectFromJson<CustomerQueueMessage>();

            Thread.Sleep(60000);

            await output.AddAsync(customerQueueMessage);
            await output.FlushAsync();
            await messageActions.CompleteMessageAsync(Message);
        }
    }
}
