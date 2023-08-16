using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BlazorGraph
{
    public class MermaidGraphGenerator
    {
        private static AppSettings _appSettings;
        public MermaidGraphGenerator(AppSettings appSettings)
        {
            _appSettings = appSettings;
        }
        public string GenerateMermaidGraph(Dictionary<string, List<string>> componentRelations)
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

            // Apply styling rules to the component
            ColorComponent(sanitizedChild, mermaidStringBuilder);
        }


        private static void CloseSubgraph(StringBuilder mermaidStringBuilder)
        {
            mermaidStringBuilder.AppendLine("end");
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
        private static bool IsVendorComponent(string componentName)
        {
            return _appSettings.Vendors.Any(vendor => componentName.Contains(vendor));
        }

        private static void ColorComponent(string componentName, StringBuilder mermaidStringBuilder)
        {
            // Apply vendor color if it's a vendor component
            if (IsVendorComponent(componentName))
            {
                mermaidStringBuilder.AppendLine($"style {componentName} fill:{_appSettings.VendorComponentColor}");
            }
        }

        public void SaveToMermaidFile(string mermaidContent, string filename = "dependencyGraph.mmd")
        {
            File.WriteAllText(filename, mermaidContent);
        }
    }
}
