using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Shared
{
    public class KeyVaultSecrets
    {
        public const string SqlConnectionString = "SqlConnectionString";
        public const string cdnSiteId = "cdnSiteId";
        public const string ClientID = "ClientID";
        public const string ClientSecret = "ClientSecret";
        public const string ProductionChoicesID = "ProductionChoicesID";
        public const string TenantID = "TenantID";
        public const string MailQueueUri = "MailQueueUri";
        public const string CustomerQueueUri = "CustomerQueueUri";
        public const string OrderQueueUri = "OrderQueueUri";
        public const string StorageAccountName = "StorageAccountName";
        public const string StorageAccountKey = "StorageAccountKey";
        public const string CDNTeamID = "CDNTeamID";
        public const string Admins = "Admins";
        public const string CertificateThumbPrint = "CertificateThumbPrint";
        public const string sbholtabnavConnection = "sbholtabnavConnection";
        public const string redisConnectionString = "redisConnectionString";
        public const string ConfigConnectionString = "ConfigConnectionString";
        public const string AzureAppConfigConnection = "AzureAppConfigConnection";
        public const string CustomerCardAppId = "CustomerCardAppId";
        public const string debugFlags = "debugFlags";
    }
}
