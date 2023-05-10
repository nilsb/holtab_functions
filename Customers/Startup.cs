using Azure.Extensions.AspNetCore.Configuration.Secrets;
using Azure.Identity;
using Microsoft.Azure.Functions.Extensions.DependencyInjection;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;

[assembly: FunctionsStartup(typeof(Customers.Startup))]
namespace Customers
{
    public class Startup : FunctionsStartup
    {
        public override void ConfigureAppConfiguration(IFunctionsConfigurationBuilder builder)
        {
            var keyVaultName = Environment.GetEnvironmentVariable("KeyVaultName");
            var keyVaultUri = $"https://{keyVaultName}.vault.azure.net";
            var manager = new KeyVaultSecretManager();

            builder.ConfigurationBuilder.AddAzureKeyVault(
                new Uri(keyVaultUri), 
                new DefaultAzureCredential(),
                manager);
            
            var config = builder.ConfigurationBuilder.Build();

            var secrets = typeof(KeyVaultSecrets).GetFields(BindingFlags.Public | BindingFlags.Static | BindingFlags.FlattenHierarchy)
                .Where(f => f.FieldType == typeof(string))
                .Select(f => f.GetValue(null))
                .Cast<string>()
                .ToList();

            foreach (var secret in secrets)
            {
                var secretValue = config[$"secrets:{secret}"];
                builder.ConfigurationBuilder.AddInMemoryCollection(new Dictionary<string, string>
                {
                    { secret, secretValue }
                });

                if(secret == "AzureAppConfigConnection")
                {
                    builder.ConfigurationBuilder.AddAzureAppConfiguration(secretValue, true);
                }
            }

            config = builder.ConfigurationBuilder.Build();
        }

        public override void Configure(IFunctionsHostBuilder builder)
        {
        }
    }
}
