namespace BlazorGraph
{
    public static class ComponentGraphProcessor
    {
        public static Dictionary<string, List<string>> Relations { get; set; }
        public static HashSet<string> ProcessedNodes { get; set; } = new HashSet<string>();
        public static List<List<string>> Grid { get; set; } = new List<List<string>>();
        public static AppSettings AppSettings { get; set; }

        public static void Configure(AppSettings appSettings)
        {
            AppSettings = appSettings;
        }

        public static List<List<string>> GetNewGrid()
        {
            return new List<List<string>>();
        }

        public static HashSet<string> GetRootNodes(Dictionary<string, List<string>> componentRelations)
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

        public static bool IsRootNode(string nodeName, Dictionary<string, List<string>> componentRelations)
        {
            return !componentRelations.Values.Any(relList => relList.Contains(nodeName));
        }

        public static bool IsVendorComponent(string componentName)
        {
            return AppSettings.Vendors.Any(vendor => componentName.Contains(vendor));
        }

        public static List<List<string>> AddNodeToGrid(string nodeName, Dictionary<string, List<string>> componentRelations, List<List<string>> inputGrid)
        {
            List<List<string>> grid = new List<List<string>>(inputGrid); // Create a new instance to avoid side effects on the original grid

            // Check if node is already in the grid
            if (grid.Any(row => row.Contains(nodeName)))
                return grid;

            if (nodeName.EndsWith("State"))
            {
                if (grid.Count == 0)
                    grid.Add(new List<string>());

                grid[0].Add(nodeName);
            }
            else
            {
                var parent = GetParent(nodeName, componentRelations);
                if (parent == null)
                {
                    // Handle the node as a root node if no parent is found
                    while (grid.Count < 2)
                        grid.Add(new List<string>());

                    grid[1].Add(nodeName);
                    return grid;
                }

                int parentRowIdx = grid.FindIndex(row => row.Contains(parent));
                if (parentRowIdx == -1)
                    return grid;

                // Ensure next row exists
                while (grid.Count <= parentRowIdx + 1)
                    grid.Add(new List<string>());

                grid[parentRowIdx + 1].Add(nodeName);
            }

            return grid;
        }

        public static List<string> PreprocessRelationsBFS(Dictionary<string, List<string>> componentRelations)
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

        private static string GetParent(string childNode, Dictionary<string, List<string>> componentRelations)
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

        private static void DFS(string node, Dictionary<string, List<string>> componentRelations, HashSet<string> visited, List<string> order)
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

        public static void ConsoleWriteGrid(List<List<string>> grid)
        {
            if (grid == null || grid.Count == 0) return;

            // 1. Calculate the width of each column.
            int columnCount = grid.Max(row => row.Count);
            int[] columnWidths = new int[columnCount];

            foreach (var row in grid)
            {
                for (int col = 0; col < row.Count; col++)
                {
                    columnWidths[col] = Math.Max(columnWidths[col], row[col].Length);
                }
            }

            // Helper function to print border
            void PrintBorder()
            {
                Console.Write('+');
                foreach (var width in columnWidths)
                {
                    Console.Write(new string('-', width + 2)); // +2 for space padding on both sides
                    Console.Write('+');
                }
                Console.WriteLine();
            }

            // 2. Print the top border.
            PrintBorder();

            // 3. For each row in the grid, print the content.'
            var rowIndex = 0;
            foreach (var row in grid)
            {
                Console.Write($"row: [{rowIndex++}]");
                for (int col = 0; col < columnCount; col++)
                {
                    if (col < row.Count)
                    {
                        string cellContent = row[col];
                        int padding = (columnWidths[col] - cellContent.Length) / 2;
                        int extraPadding = (columnWidths[col] - cellContent.Length) % 2;
                        Console.Write(new string(' ', padding + 1)); // +1 for space padding
                        Console.Write($"[{col}] {cellContent}");
                        Console.Write(new string(' ', padding + 1 + extraPadding)); // +1 for space padding
                    }
                    else
                    {
                        // In case some rows have fewer columns than others, fill with spaces
                        Console.Write(new string(' ', columnWidths[col] + 2)); // +2 for space padding on both sides
                    }
                }
                Console.WriteLine();
            }

            // 4. Print the bottom border.
            PrintBorder();
        }

    }
}
