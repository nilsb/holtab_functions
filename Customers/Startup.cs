using Azure.Extensions.AspNetCore.Configuration.Secrets;
using Azure.Identity;
using Microsoft.Azure.Functions.Extensions.DependencyInjection;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Reflection;

[assembly: FunctionsStartup(typeof(Customers.Startup))]
namespace Customers
{
    public class Startup : FunctionsStartup
    {
        public override void ConfigureAppConfiguration(IFunctionsConfigurationBuilder builder)
        {
            var config = builder.ConfigurationBuilder
                .SetBasePath(builder.GetContext().ApplicationRootPath)
                .AddJsonFile("local.settings.json", optional: true, reloadOnChange: true)
                .AddEnvironmentVariables()
                .Build();

            var keyVaultName = config["KeyVaultName"];
            var keyVaultUri = $"https://{keyVaultName}.vault.azure.net";
            var manager = new KeyVaultSecretManager();
            var env = config["env"];
            var managedIdentityId = config["managedIdentityId"];

            if (!string.IsNullOrEmpty(env) && env == "Development")
            {
                var clientsecretcredentialoptions = new ClientSecretCredentialOptions();
                
                builder.ConfigurationBuilder.AddAzureKeyVault(
                    new Uri(keyVaultUri),
                    new ClientSecretCredential(
                        "b92af8fa-fba0-4e75-b132-1a903207cd98",
                        "10817a97-cf12-4a01-9105-84a03b9d4e4f",
                        "Hh48Q~QlqHoTEEh~OR6syndl1Iacbpau.VcIqbyk"),
                    manager);
            }

            if (!string.IsNullOrEmpty(env) && env == "Production")
            {
                builder.ConfigurationBuilder.AddAzureKeyVault(
                    new Uri(keyVaultUri),
                    new DefaultAzureCredential(),
                    manager);
            }

            config = builder.ConfigurationBuilder.Build();

            var secrets = typeof(KeyVaultSecrets).GetFields(BindingFlags.Public | BindingFlags.Static | BindingFlags.FlattenHierarchy)
                .Where(f => f.FieldType == typeof(string))
                .Select(f => f.GetValue(null))
                .Cast<string>()
                .ToList();

            foreach (var secret in secrets)
            {
                var secretValue = config[$"{secret}"];
                builder.ConfigurationBuilder.AddInMemoryCollection(new Dictionary<string, string>
                {
                    { secret, secretValue }
                });

                if(secret == "AzureAppConfigConnection")
                {
                    if(secretValue != null)
                    {
                        builder.ConfigurationBuilder.AddAzureAppConfiguration(secretValue, true);
                    }
                }
            }

            config = builder.ConfigurationBuilder.Build();
        }

        public override void Configure(IFunctionsHostBuilder builder)
        {
            var config = builder.GetContext().Configuration;
            builder.Services.AddSingleton(config);
        }
    }
}
