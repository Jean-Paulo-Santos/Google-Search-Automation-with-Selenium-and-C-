using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using Microsoft.Extensions.Logging;
using System;
using WorkerServiceForResearch;

namespace WorkerServiceForResearch
{
    public class Program
    {
        public static void Main(string[] args)
        {
            CreateHostBuilder(args).Build().Run();
        }

        public static IHostBuilder CreateHostBuilder(string[] args) =>
            Host.CreateDefaultBuilder(args)
                .ConfigureServices((hostContext, services) =>
                {
                    // Adiciona o servi�o do Worker
                    services.AddHostedService<Worker>();

                    // Adiciona a classe GoogleSearch como um servi�o injet�vel
                    services.AddSingleton<GoogleSearch>();

                    // Configura o logger
                    services.AddLogging(configure => configure.AddConsole());
                });
    }
}
