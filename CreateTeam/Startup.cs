using Microsoft.Azure.Functions.Extensions.DependencyInjection;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using System;
using System.Collections.Generic;
using System.Security;
using System.Security.Cryptography.X509Certificates;
using System.Text;

[assembly: FunctionsStartup(typeof(CreateTeam.Startup))]
namespace CreateTeam
{
    public class Startup : FunctionsStartup
    {
        public override void ConfigureAppConfiguration(IFunctionsConfigurationBuilder builder)
        {
            string cs = Environment.GetEnvironmentVariable("ConfigConnectionString");
        }

        public override void Configure(IFunctionsHostBuilder builder)
        {
            var config = builder.GetContext().Configuration;
            var azureFunctionSettings = new AzureFunctionSettings();
            config.Bind(azureFunctionSettings);
            
            builder.Services.AddSingleton(config);
            
            builder.Services.AddPnPCore(options =>
            {
                // Disable telemetry because of mixed versions on AppInsights dependencies
                options.DisableTelemetry = true;

                // Configure an authentication provider with certificate (Required for app only)
                var authProvider = new PnP.Core.Auth.X509CertificateAuthenticationProvider(azureFunctionSettings.ClientID,
                    azureFunctionSettings.TenantID,
                    new X509Certificate2(System.IO.Path.Combine(builder.GetContext().ApplicationRootPath,"holtabappkeys-LogicAppGraphQuery-20211203.pfx"), new SecureString())
                );

                // And set it as default
                options.DefaultAuthenticationProvider = authProvider;

                // Add a default configuration with the site configured in app settings
                //options.Sites.Add("Default",
                //       new PnP.Core.Services.Builder.Configuration.PnPCoreSiteOptions
                //       {
                //           SiteUrl = azureFunctionSettings.SiteUrl,
                //           AuthenticationProvider = authProvider
                //       });
            });
        }
    }
}
