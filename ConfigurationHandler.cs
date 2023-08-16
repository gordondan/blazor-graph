using Microsoft.Extensions.Configuration;

namespace BlazorGraph
{
    public class ConfigurationHandler
    {
        private IConfiguration _configuration;

        public ConfigurationHandler(string[] args)
        {
            _configuration = new ConfigurationBuilder()
                .AddCommandLine(args)
                .AddJsonFile("appsettings.json")
                .Build();
        }

        public AppSettings GetAppSettings()
        {
            var settings = new AppSettings();
            _configuration.GetSection("BlazorComponentAnalyzer").Bind(settings);
            _configuration.Bind(settings);
            return settings;
        }

    }

}
