using BlazorGraph;
using Microsoft.CodeAnalysis;

namespace BlazorComponentAnalyzer
{
    class Program
    {
        static void Main(string[] args)
        {
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

            Console.WriteLine("Mermaid dependency graph generated in 'blazorDependencyGraph.mmd'.");
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
                    Console.WriteLine("Please provide the directory path where the .razor files are located:");
                    directory = Console.ReadLine();
                }
            }

            if (!Directory.Exists(directory))
            {
                Console.WriteLine($"Directory '{directory}' does not exist.");
                return Enumerable.Empty<string>();
            }

            var razorFiles = Directory.GetFiles(directory, "*.razor", SearchOption.AllDirectories);

            if (!razorFiles.Any())
            {
                Console.WriteLine("No .razor files found.");
                return Enumerable.Empty<string>();
            }

            Console.WriteLine($"Found {razorFiles.Length} .razor file(s) in the directory.");

            return razorFiles;
        }

    }
}
