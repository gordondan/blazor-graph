using Microsoft.Office.Interop.Visio;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BlazorGraph
{
    public class ComponentGraphProcessor
    {
        private Dictionary<string, List<string>> componentRelations;
        private HashSet<string> processedNodes = new HashSet<string>();
        private HashSet<string> stateNodes = new HashSet<string>();
        private List<List<string>> grid = new List<List<string>>();
        private AppSettings _appSettings;

        public ComponentGraphProcessor(Dictionary<string, List<string>> relations,AppSettings appSettings)
        {
            this.componentRelations = relations;
            this._appSettings = appSettings;
        }

        // Move methods related to data processing like GetRootNodes, PreprocessRelationsBFS,
        // IsRootNode, AddNodeToGrid, GetParent, DFS, and other related methods to this class.

        public Dictionary<string, List<string>> Relations => componentRelations;
        public HashSet<string> ProcessedNodes => processedNodes;
        public List<List<string>> Grid => grid;



        public HashSet<string> GetRootNodes(Dictionary<string, List<string>> componentRelations)
        {
            HashSet<string> allNodes = new HashSet<string>(componentRelations.Keys);
            foreach (var relatedComponents in componentRelations.Values)
            {
                foreach (var relatedComponent in relatedComponents)
                {
                    allNodes.Remove(relatedComponent); // removes child nodes, leaving only the root nodes
                }
            }
            return allNodes;
        }

        public bool IsRootNode(string nodeName, Dictionary<string, List<string>> componentRelations)
        {
            return !componentRelations.Values.Any(relList => relList.Contains(nodeName));
        }

        public bool IsVendorComponent(string componentName)
        {
            return _appSettings.Vendors.Any(vendor => componentName.Contains(vendor));
        }

        public void AddNodeToGrid(string nodeName, Dictionary<string, List<string>> componentRelations)
        {
            // Check if node is already in the grid
            if (grid.Any(row => row.Contains(nodeName)))
                return;

            if (nodeName.EndsWith("State"))
            {
                if (grid.Count == 0)
                    grid.Add(new List<string>());

                grid[0].Add(nodeName);
            }
            else if (IsRootNode(nodeName, componentRelations))
            {
                while (grid.Count < 2)
                    grid.Add(new List<string>());

                grid[1].Add(nodeName);
            }
            else
            {
                var parent = GetParent(nodeName, componentRelations);
                if (parent == null) return;

                int parentRowIdx = grid.FindIndex(row => row.Contains(parent));
                if (parentRowIdx == -1) return;

                // Ensure next row exists
                while (grid.Count <= parentRowIdx + 1)
                    grid.Add(new List<string>());

                grid[parentRowIdx + 1].Add(nodeName);
            }
        }

        public List<string> PreprocessRelationsBFS(Dictionary<string, List<string>> componentRelations)
        {
            HashSet<string> visited = new HashSet<string>();
            List<string> order = new List<string>();
            Queue<string> queue = new Queue<string>();

            // Add all root nodes to the queue to start BFS from them
            foreach (var key in componentRelations.Keys)
            {
                if (IsRootNode(key, componentRelations))
                {
                    queue.Enqueue(key);
                    visited.Add(key);
                }
            }

            while (queue.Count > 0)
            {
                var currentNode = queue.Dequeue();
                order.Add(currentNode);

                if (componentRelations.ContainsKey(currentNode))
                {
                    foreach (var child in componentRelations[currentNode])
                    {
                        if (!visited.Contains(child))
                        {
                            visited.Add(child);
                            queue.Enqueue(child);
                        }
                    }
                }
            }

            return order;
        }

        private string GetParent(string childNode, Dictionary<string, List<string>> componentRelations)
        {
            foreach (var pair in componentRelations)
            {
                if (pair.Value.Contains(childNode))
                {
                    return pair.Key;  // The parent node
                }
            }
            return null;  // No parent found
        }

        private void DFS(string node, Dictionary<string, List<string>> componentRelations, HashSet<string> visited, List<string> order)
        {
            if (visited.Contains(node))
            {
                return;
            }

            visited.Add(node);
            order.Add(node);

            if (componentRelations.ContainsKey(node))
            {
                foreach (var child in componentRelations[node])
                {
                    DFS(child, componentRelations, visited, order);
                }
            }
        }
    }

}
