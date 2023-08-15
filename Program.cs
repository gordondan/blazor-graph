﻿using BlazorGraph;
using Microsoft.AspNetCore.Razor.Language;
using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.CSharp;
using Microsoft.CodeAnalysis.CSharp.Syntax;
using System.Text.RegularExpressions;

namespace BlazorComponentAnalyzer
{
    class Program
    {
        static void Main(string[] args)
        {
            var razorFiles = GetRazorFiles(args);

            foreach (var razorFilePath in razorFiles)
            {
                var razorContent = File.ReadAllText(razorFilePath);
                var components = ExtractBlazorComponents(razorContent);

                Console.WriteLine($"\nFor {razorFilePath}, found {components.Count} component(s):");

                foreach (var component in components)
                {
                    Console.WriteLine($"- {component}");
                }
            }
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

        private static IEnumerable<string> ExtractComponentNamesFromLiteral(string literalValue)
        {
            var matches = Regex.Matches(literalValue, @"<(\w+)(\s|>)");
            return matches.Cast<Match>().Select(match => match.Groups[1].Value).Where(name => !IsHtmlTag(name));
        }

        private static bool IsHtmlTag(string tagName)
        {
            // You can expand this list based on common HTML tags you encounter.
            var htmlTags = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
    {
        "a", "div", "span", "h1", "h2", "h3", "p", "br", "input", "button", "form",
        "img", "ul", "li", "ol", "table", "thead", "tbody", "tr", "td", "th"
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
        private static IEnumerable<string> GetRazorFiles(string[] args)
        {
            string directory;

            // Check if a command-line argument was provided.
            if (args.Length > 0)
            {
                directory = args[0];
            }
            else
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
