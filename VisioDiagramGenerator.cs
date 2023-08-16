using Microsoft.Office.Interop.Visio;
using System.Runtime.InteropServices;

namespace BlazorGraph
{
    public class VisioDiagramGenerator
    {
        private const double init_y = 10;
        private AppSettings _appSettings;
        private List<List<string>> _grid = new List<List<string>>();
        private double currentX = .5;
        private double currentY = 10;
        private HashSet<string> processedNodes = new HashSet<string>();
        private HashSet<string> stateNodes = new HashSet<string>();
        private double x_offset = 2.25;
        private double y_offset = -1.25;
        private List<GraphNode> graphNodes = new List<GraphNode>();
        private Dictionary<GraphNode, (double X, double Y)> nodePositions = new Dictionary<GraphNode, (double X, double Y)>();

        public VisioDiagramGenerator(AppSettings appSettings)
        {
            _appSettings = appSettings;
            ComponentGraphProcessor.Configure(_appSettings);
            _grid = ComponentGraphProcessor.GetNewGrid();
        }

        public void GenerateVisioDiagram(Dictionary<string, List<string>> componentRelations)
        {
            var application = new Microsoft.Office.Interop.Visio.Application();
            application.Visible = true;
            var document = application.Documents.Add("");
            var page = application.ActivePage;

            List<GraphNode> graphNodes = LoadGrid(componentRelations);
            
            WriteToVisio(graphNodes, page);

            application.ActiveWindow.ViewFit = (short)VisWindowFit.visFitPage;
            application.ActiveWindow.Zoom = 1;  // Sets zoom to 100%

            string currentDirectory = Directory.GetCurrentDirectory();
            string fullPath = System.IO.Path.Combine(currentDirectory, _appSettings.VisioFileName);
            document.SaveAs(fullPath);

            document.Close();
            application.Quit();
        }

        private List<GraphNode> LoadGrid(Dictionary<string, List<string>> componentRelations)
        {
            List<GraphNode> graphNodes = new List<GraphNode>();
            List<string> orderedNodes = ComponentGraphProcessor.PreprocessRelationsBFS(componentRelations);

            foreach (var rootNode in orderedNodes)
            {
                Queue<string> nodesToProcess = new Queue<string>();
                nodesToProcess.Enqueue(rootNode);


                while (nodesToProcess.Count > 0)
                {
                    string currentComponent = nodesToProcess.Dequeue();

                    // Get or create the graph node for the current component
                    GraphNode node = graphNodes.FirstOrDefault(n => n.ComponentName == currentComponent);
                    if (node == null)
                    {
                        node = new GraphNode(currentComponent);
                        graphNodes.Add(node);
                    }

                    _grid = ComponentGraphProcessor.AddNodeToGrid(currentComponent, componentRelations, _grid);
                    //ComponentGraphProcessor.ConsoleWriteGrid(_grid);

                    //SetPositionFromGrid(node, currentComponent); // Update this method to work with GraphNode
                    if (componentRelations.ContainsKey(currentComponent))
                    {
                        foreach (string relatedComponent in componentRelations[currentComponent])
                        {
                            if (!processedNodes.Contains(relatedComponent))
                            {
                                nodesToProcess.Enqueue(relatedComponent);

                                // Add related node to the current graph node
                                GraphNode relatedNode = graphNodes.FirstOrDefault(n => n.ComponentName == relatedComponent);
                                if (relatedNode == null)
                                {
                                    relatedNode = new GraphNode(relatedComponent);
                                    graphNodes.Add(relatedNode);
                                }

                                node.AddRelatedComponent(relatedNode);
                            }
                        }
                    }
                }
            }
            return graphNodes;
        }
        private void ConnectShapes(Shape shape1, Shape shape2, Page page)
        {
            if (shape1 == null || shape2 == null)
            {
                Console.WriteLine($"Warning: Attempted to connect a null shape. Shape1: {shape1?.Text}, Shape2: {shape2?.Text}");
                return;
            }
            shape1.AutoConnect(shape2, VisAutoConnectDir.visAutoConnectDirNone);
        }
        private Shape CreateShape(Page page, GraphNode node)
        {
            var x = node.X;
            var y = node.Y;
            var componentName = node.ComponentName;

            Console.WriteLine($"Creating shape for: {componentName} at X: {x}, Y: {y}");

            EnsurePageSize(page);

            Shape shape = page.DrawRectangle(x,y, x + 2, y + 1);
            shape.Text = componentName;

            // Setting shape rounding for rounded rectangle
            shape.CellsU["Rounding"].ResultIU = 0.1;

            // Default to navy blue with white text
            shape.CellsU["FillForegnd"].FormulaU = "RGB(0, 0, 128)";
            shape.CellsU["Char.Color"].FormulaU = "RGB(255, 255, 255)";  // White color for text

            // If it's a vendor component, set to lime green with white text
            {
            if (ComponentGraphProcessor.IsVendorComponent(componentName))
                try
                {
                    shape.CellsU["FillForegnd"].FormulaU = "RGB(50, 205, 50)";  // Lime green
                    shape.CellsU["Char.Color"].FormulaU = "RGB(255, 255, 255)";  // White color for text
                }
                catch (COMException ex)
                {
                    // Handle the error, perhaps logging it or notifying the user
                    Console.WriteLine($"Error setting color for component {componentName}: {ex.Message}");
                }
            }
            if (ComponentGraphProcessor.IsStateComponent(componentName))
            {
                try
                {
                    shape.CellsU["FillForegnd"].FormulaU = "RGB(255, 165, 0)";  // Orange
                    shape.CellsU["Char.Color"].FormulaU = "RGB(255, 255, 255)";  // White color for text
                }
                catch (COMException ex)
                {
                    // Handle the error
                    Console.WriteLine($"Error setting color for state component {componentName}: {ex.Message}");
                }
            }
            return shape;
        }

        private void EnsurePageSize(Page page)
        {
            const double margin = 1; // some space on all sides

            // Ensure height
            if (currentY - margin < 0)
            {
                double heightIncrease = Math.Abs(currentY) + margin;
                page.PageSheet.CellsU["PageHeight"].ResultIU += heightIncrease;
                currentY += heightIncrease;
            }

            // Ensure width
            if (currentX + 2 + margin > page.PageSheet.CellsU["PageWidth"].ResultIU)
            {
                page.PageSheet.CellsU["PageWidth"].ResultIU = currentX + margin;
            }
        }

        private Shape GetShapeByName(Page page, string name)
        {
            foreach (Shape shape in page.Shapes)
            {
                if (shape.Text == name)
                {
                    return shape;
                }
            }
            return null;
        }
        private List<GraphNode> GetPositionsForNodes(List<GraphNode> graphNodes)
        {
            return graphNodes.Select(node => GetPositionForNode(node)).ToList();
        }

        private void WriteToVisio(List<GraphNode> graphNodes, Page page)
        {
            var positionedNodes = GetPositionsForNodes(graphNodes);
            ConsoleWriteGrid(positionedNodes);  // This method will print grid for debugging

            // 1. Draw all the shapes first
            foreach (var positionedNode in positionedNodes)
            {
                Console.WriteLine($"Processing: {positionedNode.ComponentName} at X: {positionedNode.X}, Y: {positionedNode.Y}");

                EnsurePageSize(page);
                CreateShape(page, positionedNode);
            }

            // 2. Now make the connections between the shapes
            foreach (var positionedNode in positionedNodes)
            {
                Shape currentShape = GetShapeByName(page, positionedNode.ComponentName);
                foreach (var relatedNode in positionedNode.RelatedComponents)
                {
                    Shape relatedShape = GetShapeByName(page, relatedNode.ComponentName);
                    if (currentShape != null && relatedShape != null) // Safety check
                    {
                        ConnectShapes(currentShape, relatedShape, page);
                    }
                }
            }
        }


        private GraphNode GetPositionForNode(GraphNode node)
        {
            int rowIndex = -1;
            int columnIndex = -1;

            for (int i = 0; i < _grid.Count; i++)
            {
                columnIndex = _grid[i].IndexOf(node.ComponentName);
                if (columnIndex != -1)
                {
                    rowIndex = i;
                    break;
                }
            }

            if (rowIndex != -1 && columnIndex != -1)
            {
                double x = columnIndex * x_offset + 0.5;  // added offset
                double y = init_y - rowIndex * y_offset;

                // Here, instead of modifying the existing node, we create a new one and return it.
                GraphNode newNode = new GraphNode(node.ComponentName)
                {
                    X = x,
                    Y = y,
                    RelatedComponents = new List<GraphNode>(node.RelatedComponents)  // shallow copy to maintain the same list of related nodes
                };

                return newNode;
            }

            return node;  // If there's no change in the position, return the original node.
        }

        private void ConsoleWriteGrid(List<GraphNode> graphNodes)
        {
            foreach (var node in graphNodes)
            {
                Console.WriteLine($"Node: {node.ComponentName}, X: {node.X}, Y: {node.Y}");
                foreach (var related in node.RelatedComponents)
                {
                    Console.WriteLine($"\tRelated: {related.ComponentName}");
                }
            }
        }


    }
}