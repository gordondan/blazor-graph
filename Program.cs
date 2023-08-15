using BlazorGraph;
using Microsoft.AspNetCore.Razor.Language;
using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.CSharp;
using Microsoft.CodeAnalysis.CSharp.Syntax;
using Microsoft.CodeAnalysis.Options;
using Microsoft.Extensions.Configuration;
using System.Text;
using System.Text.RegularExpressions;

namespace BlazorComponentAnalyzer
{
    class Program
    {
        static void Main(string[] args)
        {
            var configurationBuilder = new ConfigurationBuilder()
                .AddCommandLine(args);

            var configuration = configurationBuilder.Build();

            string directory = configuration["directory"];
            string node = configuration["starting-node"];

            Run(directory, node);
        }

        static void Run(string directory, string node)
        {
            var razorFiles = GetRazorFiles(directory);
            var componentRelations = new Dictionary<string, List<string>>();

            foreach (var razorFilePath in razorFiles)
            {
                var razorContent = File.ReadAllText(razorFilePath);
                var components = ExtractBlazorComponents(razorContent);

                // Assuming the filename itself (without path or extension) represents the component's name
                var currentComponent = Path.GetFileNameWithoutExtension(razorFilePath);

                componentRelations[currentComponent] = components;
            }
            componentRelations = SortDictionary(componentRelations,node);

            var mermaidGraph = GenerateMermaidGraph(componentRelations);
            SaveToMermaidFile(mermaidGraph);

            Console.WriteLine("Mermaid dependency graph generated in 'blazorDependencyGraph.mmd'.");
        }

        private static Dictionary<string, List<string>> SortDictionary(Dictionary<string, List<string>> componentRelations, string startingNode)
        {
            if (componentRelations.ContainsKey(startingNode))
            {
                var startingNodeRelations = new Dictionary<string, List<string>>
            {
                { startingNode, componentRelations[startingNode] }
            };
                componentRelations.Remove(startingNode);
                componentRelations = startingNodeRelations.Concat(componentRelations).ToDictionary(k => k.Key, v => v.Value);
            }
            return componentRelations;
        }


        public static List<string> ExtractBlazorComponents(string razorContent)
        {
            var csharpCode = GenerateCSharpFromRazor(razorContent);
            var componentNames = ExtractComponentUsagesFromGeneratedCSharp(csharpCode);

            return componentNames;
        }

        private static string GenerateCSharpFromRazor(string razorContent)
        {
            var engine = RazorProjectEngine.Create(RazorConfiguration.Default, RazorProjectFileSystem.Create("/"), (builder) => { });

            var sourceDocument = RazorSourceDocument.Create(razorContent, "RazorFile");
            var projectItem = new CustomRazorProjectItem(sourceDocument);
            var codeDocument = engine.Process(projectItem);

            var csharp = codeDocument.GetCSharpDocument().GeneratedCode;
            return csharp;
        }

        private static List<string> ExtractComponentUsagesFromGeneratedCSharp(string csharpCode)
        {
            var tree = CSharpSyntaxTree.ParseText(csharpCode);
            var root = tree.GetCompilationUnitRoot();

            var componentNames = new HashSet<string>();

            var invocations = root.DescendantNodes().OfType<InvocationExpressionSyntax>().ToList();
            foreach (var invocation in invocations)
            {
                if (invocation.Expression is IdentifierNameSyntax methodName
                    && methodName.Identifier.Text == "WriteLiteral")
                {
                    var argument = invocation.ArgumentList.Arguments.FirstOrDefault();
                    if (argument != null)
                    {
                        var argValue = argument.Expression.NormalizeWhitespace().ToFullString().Trim('\"');
                        componentNames.UnionWith(ExtractComponentNamesFromLiteral(argValue));
                    }
                }
            }

            return componentNames.ToList();
        }
        private static string SanitizeComponentName(string componentName)
        {
            // List of keywords/tags that might conflict with Mermaid or HTML.
            var reservedKeywords = new List<string> { "style", "strong", /* ... add others as needed ... */ };

            if (reservedKeywords.Contains(componentName))
            {
                return $"tag_{componentName}"; // Prefix with "tag_" or any other suitable prefix.
            }

            return componentName;
        }

        private static string GenerateMermaidGraph(Dictionary<string, List<string>> componentRelations)
        {
            StringBuilder mermaidStringBuilder = new StringBuilder();
            mermaidStringBuilder.AppendLine("graph TD");  // TD denotes top-down layout

            var processedNodes = new HashSet<string>();

            foreach (var component in componentRelations)
            {
                ProcessComponent(component, processedNodes, mermaidStringBuilder);
            }

            return mermaidStringBuilder.ToString();
        }

        private static void ProcessComponent(KeyValuePair<string, List<string>> component, HashSet<string> processedNodes, StringBuilder mermaidStringBuilder)
        {
            string sanitizedParent = SanitizeComponentName(component.Key);
            bool isNewSubgraph = !processedNodes.Contains(sanitizedParent);

            if (isNewSubgraph)
            {
                StartNewSubgraph(sanitizedParent, mermaidStringBuilder, processedNodes);
            }

            ProcessRelatedComponents(component.Value, sanitizedParent, mermaidStringBuilder, processedNodes);

            if (isNewSubgraph)
            {
                CloseSubgraph(mermaidStringBuilder);
            }
        }

        private static void StartNewSubgraph(string sanitizedParent, StringBuilder mermaidStringBuilder, HashSet<string> processedNodes)
        {
            mermaidStringBuilder.AppendLine($"subgraph {sanitizedParent}_g");
            mermaidStringBuilder.AppendLine(sanitizedParent);
            processedNodes.Add(sanitizedParent);
        }

        private static void ProcessRelatedComponents(List<string> relatedComponents, string sanitizedParent, StringBuilder mermaidStringBuilder, HashSet<string> processedNodes)
        {
            foreach (var relatedComponent in relatedComponents)
            {
                LinkParentToChild(sanitizedParent, relatedComponent, mermaidStringBuilder, processedNodes);
            }
        }

        private static void LinkParentToChild(string sanitizedParent, string relatedComponent, StringBuilder mermaidStringBuilder, HashSet<string> processedNodes)
        {
            string sanitizedChild = SanitizeComponentName(relatedComponent);
            mermaidStringBuilder.AppendLine($"{sanitizedParent}--> {sanitizedChild}");
            processedNodes.Add(sanitizedChild); // Add the child to the processed nodes
        }

        private static void CloseSubgraph(StringBuilder mermaidStringBuilder)
        {
            mermaidStringBuilder.AppendLine("end");
        }



        private static void SaveToMermaidFile(string mermaidContent, string filename = "dependencyGraph.mmd")
        {
            File.WriteAllText(filename, mermaidContent);
        }


        private static IEnumerable<string> ExtractComponentNamesFromLiteral(string literalValue)
        {
            var matches = Regex.Matches(literalValue, @"<(\w+)(\s|>)");
            return matches.Cast<Match>().Select(match => match.Groups[1].Value)
                          .Where(name => !IsHtmlTag(name) && IsValidComponentName(name));
        }

        private static bool IsValidComponentName(string name)
        {
            // Example criteria: Component names start with a capital letter
            return Char.IsUpper(name[0]);
        }


        private static bool IsHtmlTag(string tagName)
        {
            // You can expand this list based on common HTML tags you encounter.
            var htmlTags = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
    {
        "a", "div", "span", "h1", "h2", "h3", "p", "br", "input", "button", "form",
        "img", "ul", "li", "ol", "table", "thead", "tbody", "tr", "td", "th",
        "style", "strong", "em", "b", "i", "u", "script", "link"
        // ... add more HTML tags if needed
    };

            return htmlTags.Contains(tagName);
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
