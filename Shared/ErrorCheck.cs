using Microsoft.Extensions.Logging;
using Shared.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Shared
{
    public class ErrorCheck
    {
        private readonly Settings? settings;
        private readonly Services? services;
        private readonly ILogger? log;

        public ErrorCheck(ILogger _log, Settings _settings, Services _services)
        {
            log = _log;
            settings = _settings;
            services = _services;
        }

        public bool CheckOrder(Order? order, string Caller)
        {
            bool returnValue = true;

            if (order == null)
            {
                log?.LogError($"Fatal error! {Caller}: Order object was null.");
                returnValue = false;
            }

            if (order?.ID == Guid.Empty)
            {
                log?.LogTrace($"Fatal error! {Caller}: Order database ID was empty");
                returnValue = false;
            }

            return returnValue;
        }

        public bool CheckInit()
        {
            bool returnValue = true;

            if(services == null)
            {
                log?.LogError("Fatal error! Initialization failed: Services object is null.");
                returnValue = false;
            }

            if (services?.init == false)
            {
                log?.LogError("Fatal error! Initialization failed: SqlConnectionString missing.");
                returnValue = false;
            }

            if(settings == null)
            {
                log?.LogError("Fatal error! Initialization failed: Settings object is null.");
                returnValue = false;
            }

            if(settings?.log == null)
            {
                returnValue = false;
            }

            if(settings?.context == null)
            {
                log?.LogError("Fatal error! Initialization failed: Settings: ExecutionContext object is null.");
                returnValue = false;
            }

            if(string.IsNullOrEmpty(settings?.SqlConnectionString))
            {
                log?.LogError("Fatal error! Initialization failed: SqlConnectionString missing.");
                returnValue = false;
            }

            if(settings?.GraphClient == null)
            {
                log?.LogError("Fatal error! Initialization failed: Settings: GraphClient object is null.");
                returnValue = false;
            }

            if (string.IsNullOrEmpty(settings?.cdnSiteId))
            {
                log?.LogError("Fatal error! Initialization failed: cndSiteId config value missing.");
                returnValue = false;
            }

            if (string.IsNullOrEmpty(settings?.CDNTeamID))
            {
                log?.LogError("Fatal error! Initialization failed: CDNTeamID config value missing.");
                returnValue = false;
            }

            if (string.IsNullOrEmpty(settings?.TenantID))
            {
                log?.LogError("Fatal error! Initialization failed: TenantID config value missing.");
                returnValue = false;
            }

            if (string.IsNullOrEmpty(settings?.ClientID))
            {
                log?.LogError("Fatal error! Initialization failed: ClientID config value missing.");
                returnValue = false;
            }

            if (string.IsNullOrEmpty(settings?.ClientSecret))
            {
                log?.LogError("Fatal error! Initialization failed: ClientSecret config value missing.");
                returnValue = false;
            }

            if (string.IsNullOrEmpty(settings?.CustomerListID))
            {
                log?.LogError("Fatal error! Initialization failed: CustomerListID config value missing.");
                returnValue = false;
            }

            if (string.IsNullOrEmpty(settings?.OrderListID))
            {
                log?.LogError("Fatal error! Initialization failed: OrderListID config value missing.");
                returnValue = false;
            }

            return returnValue;
        }
    }
}
