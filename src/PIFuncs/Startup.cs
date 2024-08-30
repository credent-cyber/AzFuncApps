using Microsoft.Azure.Functions.Extensions.DependencyInjection;
using System;
using System.Collections.Generic;
using System.Text;
using PnP.Core.Services;
using PnP.Core;
using PnP.Core.Auth;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Configuration;
using System.Security.Cryptography.X509Certificates;
using System.IO;
using Serilog;
using Microsoft.Azure.WebJobs.Hosting;
using Microsoft.AspNetCore.Builder; // Add this namespace


[assembly: WebJobsStartup(typeof(Demo.Startup))]

namespace Demo
{

    public class Startup : FunctionsStartup
    {
        private const string TenantID = "cf92019c-152d-42f6-bbcc-0cf96e6b0108";

        public override void Configure(IFunctionsHostBuilder builder)
        {
            var dir = "C:/home/site/wwwroot";

#if DEBUG
            dir = Directory.GetCurrentDirectory();
#endif

            var config = new ConfigurationBuilder()
               .SetBasePath(dir)
               .AddJsonFile("host.json", optional: true, reloadOnChange: true)
               .AddEnvironmentVariables()
               .Build();

            var certPath = Path.Combine(dir, config.GetValue<string>("CertPath"));
            var certPwd = string.Empty;

#if DEBUG 
            certPwd = "D0n0ts@ythis";
#else
            certPwd = GetEnvironmentVariable("Pwd");
#endif
            builder.Services.AddLogging((builder) =>
            {
                Log.Logger = new LoggerConfiguration()
                    .WriteTo.File("logs/func.logs")
                    .CreateLogger();

                builder.AddSerilog();
            });

            var settings = new AzureFunctionSettings();
            config.Bind(settings);

            builder.Services.AddSingleton(settings);

            var cert = new X509Certificate2(certPath, settings.Pwd);
            var log = Log.Logger;

            log.Information($"Certificate Thumbprint: {settings.CertificateThumbPrint}");

            builder.Services.AddPnPCore(options => {
                options.DisableTelemetry = true;
                var authProvider = new X509CertificateAuthenticationProvider(
                    settings.ClientId,
                    settings.TenantId,
                    cert
                );

                options.DefaultAuthenticationProvider = authProvider;

                options.Sites.Add("Default", new PnP.Core.Services.Builder.Configuration.PnPCoreSiteOptions
                {
                    SiteUrl = settings.SiteUrl,
                    AuthenticationProvider = authProvider
                });

                options.Sites.Add("TestPortal", new PnP.Core.Services.Builder.Configuration.PnPCoreSiteOptions
                {
                    SiteUrl = settings.TestPortal,
                    AuthenticationProvider = authProvider
                });
            });

            // Add CORS policy
            //builder.Services.AddCors(options =>
            //{
            //    options.AddPolicy("AllowSpecificOrigin",
            //        builder => builder.WithOrigins("https://credentinfotec.sharepoint.com")
            //                          .AllowAnyHeader()
            //                          .AllowAnyMethod());
            //});
        }

        private static string GetEnvironmentVariable(string name)
        {
            return System.Environment.GetEnvironmentVariable(name, EnvironmentVariableTarget.Process);
        }
    }
}
