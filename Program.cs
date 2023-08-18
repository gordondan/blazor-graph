using BlazorGraph;
using Microsoft.CodeAnalysis;
using Serilog;

namespace BlazorComponentAnalyzer
{
    internal class Program
    {
        private static void Main(string[] args)
        {
            DeleteLogFiles();

            // Set up Serilog logging
            Log.Logger = new LoggerConfiguration()
                .MinimumLevel.Debug()
                .Enrich.WithProperty("Timestamp", DateTime.UtcNow)
                .WriteTo.Console(new CustomConsoleFormatter())
                .WriteTo.File(new CustomFileFormatter(),
                    "log.txt",
                    rollingInterval: RollingInterval.Day,
                    rollOnFileSizeLimit: true,
                    shared: true,
                    retainedFileCountLimit: 10)
                .CreateLogger();

            try
            {
                // Get the MST/MDT time zone
                TimeZoneInfo mstTimeZone = TimeZoneInfo.FindSystemTimeZoneById("Mountain Standard Time");
                // Convert UTC to MST/MDT
                DateTime localTime = TimeZoneInfo.ConvertTimeFromUtc(DateTime.UtcNow, mstTimeZone);
                Log.Verbose($"Starting Blazor Graph {localTime.ToLongDateString()} {localTime.ToLongTimeString()}", true);

                var configHandler = new ConfigurationHandler(args);
                var settings = configHandler.GetAppSettings();

                var razorFiles = GetRazorFiles(settings.Directory);
                BlazorComponentExtractor.Configure(settings);
                var componentRelations = BlazorComponentExtractor.ExtractComponentRelationsFromRazorFiles(razorFiles);
                BlazorComponentExtractor.PrintComponentRelations(componentRelations);

                var graphGenerator = new MermaidGraphGenerator(settings);
                var graph = graphGenerator.GenerateMermaidGraph(componentRelations);
                graphGenerator.SaveToMermaidFile(graph);

                var visioGenerator = new VisioDiagramGenerator(settings);
                visioGenerator.GenerateVisioDiagram(componentRelations);

                Log.Information("Mermaid dependency graph generated in 'blazorDependencyGraph.mmd'.");
            }
            catch (Exception ex)
            {
                // Log any unhandled exceptions
                Log.Fatal(ex, "An unexpected error occurred.");
            }
            finally
            {
                // Ensure all logs are flushed before the application exits
                Log.CloseAndFlush();
            }
        }

        private static string GetCurrentSolutionDirectory()
        {
            var currentDirectory = Directory.GetCurrentDirectory();
            var solutionFile = Directory.GetFiles(currentDirectory, "*.sln").FirstOrDefault();

            if (solutionFile != null)
            {
                return Path.GetDirectoryName(solutionFile);
            }

            return null;
        }

        private static IEnumerable<string> GetRazorFiles(string inputDirectory)
        {
            string directory = inputDirectory;

            if (string.IsNullOrEmpty(directory))
            {
                directory = GetCurrentSolutionDirectory();

                if (string.IsNullOrEmpty(directory))
                {
                    Log.Verbose("Please provide the directory path where the .razor files are located:");
                    directory = Console.ReadLine();
                }
            }

            if (!Directory.Exists(directory))
            {
                Log.Verbose($"Directory '{directory}' does not exist.");
                return Enumerable.Empty<string>();
            }

            var razorFiles = Directory.GetFiles(directory, "*.razor", SearchOption.AllDirectories);

            if (!razorFiles.Any())
            {
                Log.Verbose("No .razor files found.");
                return Enumerable.Empty<string>();
            }

            Log.Verbose($"Found {razorFiles.Length} .razor file(s) in the directory.");

            return razorFiles;
        }

        /// <summary>
        /// Deletes files with the word "log" in the name and ending in .txt from the current directory.
        /// </summary>
        private static void DeleteLogFiles()
        {
            var currentDirectory = Directory.GetCurrentDirectory();
            var logFiles = Directory.GetFiles(currentDirectory, "*log*.txt");

            foreach (var logFile in logFiles)
            {
                try
                {
                    File.Delete(logFile);
                    Log.Information($"Deleted log file: {logFile}");
                }
                catch (Exception ex)
                {
                    Log.Error(ex, $"Failed to delete log file: {logFile}");
                }
            }
        }
    }
}