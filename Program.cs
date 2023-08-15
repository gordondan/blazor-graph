using BlazorGraph;
using Microsoft.AspNetCore.Razor.Language;
using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.CSharp;
using Microsoft.CodeAnalysis.CSharp.Syntax;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text.RegularExpressions;

namespace BlazorComponentAnalyzer
{
    class Program
    {
        static void Main(string[] args)
        {
            if (args.Length == 0)
            {
                Console.WriteLine("Please provide the path to the .razor file as an argument.");
                return;
            }

            var filePath = args[0];

            if (!File.Exists(filePath))
            {
                Console.WriteLine($"File '{filePath}' does not exist.");
                return;
            }

            var razorContent = File.ReadAllText(filePath);
            var components = ExtractBlazorComponents(razorContent);

            Console.WriteLine($"Found {components.Count} component(s) in {filePath}:");

            foreach (var component in components)
            {
                Console.WriteLine($"- {component}");
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


    }
}