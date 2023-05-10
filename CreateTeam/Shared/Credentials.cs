using Microsoft.Extensions.Configuration;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Threading.Tasks;

namespace CreateTeam.Shared
{
    public class Credentials
    {
        public NetworkCredential GetCurrent(string Url)
        {
            var builder = new ConfigurationBuilder().AddJsonFile(@"appsettings.json");

            ICredentials credentials = CredentialCache.DefaultNetworkCredentials;
            var credential = credentials.GetCredential(new Uri(Url), "Basic");

            if (credential.UserName.Length <= 0)
            {
                var sectionUser = builder.Build().GetSection("SPUser");
                credential = new NetworkCredential(sectionUser.GetSection("UserName").Value, sectionUser.GetSection("Password").Value, sectionUser.GetSection("Domain").Value);
            }

            return credential;
        }
    }
}
