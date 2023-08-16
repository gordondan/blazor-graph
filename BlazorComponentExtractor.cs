using Microsoft.AspNetCore.Razor.Language;
using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.CSharp;
using Microsoft.CodeAnalysis.CSharp.Syntax;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Xml.Linq;

namespace BlazorGraph
{
    public class BlazorComponentExtractor
    {
        private static AppSettings _appSettings;
        public BlazorComponentExtractor(AppSettings appSettings)
        {
            _appSettings = appSettings;
        }
        public Dictionary<string,List<string>> ExtractComponentRelationsFromRazorFiles(IEnumerable<string> razorFiles)
        {
            var componentRelations = new Dictionary<string,List<string>>();
            foreach (var razorFilePath in razorFiles)
            {
                var razorContent = File.ReadAllText(razorFilePath);
                var components = ExtractBlazorComponents(razorContent);

                // Assuming the filename itself (without path or extension) represents the component's name
                var currentComponent = Path.GetFileNameWithoutExtension(razorFilePath);

                componentRelations[currentComponent] = components;
            }
            componentRelations = SortDictionary(componentRelations, _appSettings.StartingNode);

            return componentRelations;
        }
        public static List<string> ExtractBlazorComponents(string razorContent)
        {
            var csharpCode = GenerateCSharpFromRazor(razorContent);
            var componentNames = ExtractComponentUsagesFromGeneratedCSharp(csharpCode);
            componentNames = componentNames.Where(x=>!_appSettings.Skips.Contains(x)).ToList();
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

        private static Dictionary<string, List<string>> SortDictionary(Dictionary<string, List<string>> componentRelations, string startingNode)
        {
            if (startingNode != null)
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
            }
                return componentRelations;
            
        }
    }
}
