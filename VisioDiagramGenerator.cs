using Microsoft.Office.Interop.Visio;
using System.Runtime.InteropServices;

namespace BlazorGraph
{
    public class VisioDiagramGenerator
    {
        private const double init_y = 8;
        private const double margin = 1;
        private double headerHeight = 0.5;

        private AppSettings _appSettings;
        private List<List<string>> _grid = new List<List<string>>();

        private HashSet<string> processedNodes = new HashSet<string>();

        private double x_offset = 2.25;
        private double y_offset = 1.25;
        private double card_height = 1;
        private double card_width = 2;

        private List<GraphNode> graphNodes = new List<GraphNode>();
        private Dictionary<GraphNode, (double X, double Y)> nodePositions = new Dictionary<GraphNode, (double X, double Y)>();

        public VisioDiagramGenerator(AppSettings appSettings, double customHeaderHeight = 0.5)
        {
            _appSettings = appSettings;
            ComponentGraphProcessor.Configure(_appSettings);
            _grid = ComponentGraphProcessor.GetNewGrid();
            headerHeight = customHeaderHeight;
        }

        public void GenerateVisioDiagram(Dictionary<string, List<string>> componentRelations)
        {
            var application = new Microsoft.Office.Interop.Visio.Application();
            application.Visible = true;
            var document = application.Documents.Add("");
            var page = application.ActivePage;
            if (page.PageSheet.CellsU["PageWidth"].ResultIU < page.PageSheet.CellsU["PageHeight"].ResultIU)
            {
                double temp = page.PageSheet.CellsU["PageWidth"].ResultIU;
                page.PageSheet.CellsU["PageWidth"].ResultIU = page.PageSheet.CellsU["PageHeight"].ResultIU;
                page.PageSheet.CellsU["PageHeight"].ResultIU = temp;
            }

            List<GraphNode> graphNodes = LoadGrid(componentRelations);
            var positionedNodes = GetPositionsForNodes(graphNodes);
            var (maxX, maxY) = GetMaxCoordinates(positionedNodes);

            EnsurePageSize(page, maxX, maxY);

            // ... rest of the WriteToVisio method ...

            WriteToVisio(graphNodes, page);

            application.ActiveWindow.ViewFit = (short)VisWindowFit.visFitPage;
            application.ActiveWindow.Zoom = 1;  // Sets zoom to 100%

            string currentDirectory = Directory.GetCurrentDirectory();
            string fullPath = System.IO.Path.Combine(currentDirectory, _appSettings.VisioFileName);
            document.SaveAs(fullPath);

            document.Close();
            application.Quit();
        }
        private (double maxX, double maxY) GetMaxCoordinates(List<GraphNode> graphNodes)
        {
            double maxX = 0;
            double maxY = 0;

            foreach (var node in graphNodes)
            {
                if (node.X > maxX) maxX = node.X;
                if (node.Y > maxY) maxY = node.Y;
            }

            return (maxX, maxY);
        }

        private void WriteToVisio(List<GraphNode> graphNodes, Page page)
        {
            var positionedNodes = GetPositionsForNodes(graphNodes);
            ConsoleWriteGrid(positionedNodes);

            Dictionary<int, List<Shape>> shapesByRow = new Dictionary<int, List<Shape>>();

            foreach (var positionedNode in positionedNodes)
            {
                Console.WriteLine($"Processing: {positionedNode.ComponentName} at X: {positionedNode.X}, Y: {positionedNode.Y}");

                Shape currentShape = CreateShape(page, positionedNode);

                int row = (int)Math.Floor(positionedNode.Y);
                if (!shapesByRow.ContainsKey(row))
                    shapesByRow[row] = new List<Shape>();

                shapesByRow[row].Add(currentShape);
            }

            foreach (var node in graphNodes)
            {
                Shape sourceShape = GetShapeByName(page, node.ComponentName);
                foreach (var relatedNode in node.RelatedComponents)
                {
                    Shape targetShape = GetShapeByName(page, relatedNode.ComponentName);
                    ConnectShapes(sourceShape, targetShape, page);
                }
            }

            // Group shapes by row
            foreach (var rowShapes in shapesByRow.Values)
            {
                GroupShapesByRow(rowShapes, page);
            }

            // Apply auto layout after grouping
            //ApplyAutoLayout(page);
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
            if(ComponentGraphProcessor.IsStateComponent(shape1.Text) || ComponentGraphProcessor.IsStateComponent(shape2.Text))
            {
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

            Shape header = CreateHeader(page, x, y, componentName);
            Shape body = CreateBody(page, x, y - headerHeight);  
            LabelDependencies(body, node);

            // You can group the shapes into one if required
            // Shape group = page.Group(new Shape[] { header, body });
            return header;  // or return group if you grouped the shapes
        }

        private Shape CreateHeader(Page page, double x, double y, string componentName)
        {
            Console.WriteLine($"Creating Header ({componentName}): {x} {y} {x+card_width} {y-headerHeight}");
            Shape header = page.DrawRectangle(x, y, x + card_width, y - headerHeight);  // Using headerHeight for the header
            header.Text = componentName;

            // Setting shape rounding for rounded rectangle
            //header.CellsU["Rounding"].ResultIU = 0.1;

            SetShapeColor(header, componentName);
            return header;
        }

        private Shape CreateBody(Page page, double x, double y)
        {
            Console.WriteLine($"Creating Body {x} {y} {x + card_width} {y - card_height}");
            Shape body = page.DrawRectangle(x, y, x + card_width, y - card_height);
            body.CellsU["FillForegnd"].FormulaU = "RGB(255, 255, 255)"; // White body
            body.CellsU["Char.Color"].FormulaU = "RGB(0, 0, 0)";  // Black text for body details
            return body;
        }
        private void SetShapeColor(Shape shape, string componentName)
        {
            try
            {
                if (ComponentGraphProcessor.IsVendorComponent(componentName))
                {
                    shape.CellsU["FillForegnd"].FormulaU = "RGB(50, 205, 50)";  // Lime green
                    shape.CellsU["Char.Color"].FormulaU = "RGB(255, 255, 255)";  // White text
                }
                else if (ComponentGraphProcessor.IsStateComponent(componentName))
                {
                    shape.CellsU["FillForegnd"].FormulaU = "RGB(255, 165, 0)";  // Orange
                    shape.CellsU["Char.Color"].FormulaU = "RGB(255, 255, 255)";  // White text
                }
                else
                {
                    shape.CellsU["FillForegnd"].FormulaU = "RGB(0, 0, 128)"; // Default navy blue
                    shape.CellsU["Char.Color"].FormulaU = "RGB(255, 255, 255)";  // White text
                }
            }
            catch (COMException ex)
            {
                // Handle the error, perhaps logging it or notifying the user
                Console.WriteLine($"Error setting color for component {componentName}: {ex.Message}");
            }
        }

        private void LabelDependencies(Shape body, GraphNode node)
        {
            int stateDeps = node.RelatedComponents.Where(x => ComponentGraphProcessor.IsStateComponent(x.ComponentName)).Count();
            int outgoingDeps = node.RelatedComponents.Count();
            int incomingDeps = 0; // Get from node

            body.Text = $"State Dep: {stateDeps}\nOutgoing: {outgoingDeps}\nIncoming: {incomingDeps}";
        }

        private void EnsurePageSize(Page page, double maxX, double maxY)
        {
            // Ensure height
            if (maxY + margin > page.PageSheet.CellsU["PageHeight"].ResultIU)
            {
                page.PageSheet.CellsU["PageHeight"].ResultIU = maxY + margin;
            }

            // Ensure width
            if (maxX + margin > page.PageSheet.CellsU["PageWidth"].ResultIU)
            {
                page.PageSheet.CellsU["PageWidth"].ResultIU = maxX + margin;
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
        private Shape GroupShapesByRow(List<Shape> shapesInRow, Page page)
        {
            if (shapesInRow == null || !shapesInRow.Any()) return null;

            // Get the active window
            Window window = page.Application.ActiveWindow;

            // Clear current selection
            window.DeselectAll();

            // Add each shape in the list to the selection
            foreach (Shape shape in shapesInRow)
            {
                window.Select(shape,(short)VisSelectArgs.visSelect);
            }

            // The active window's selection should now contain the shapes you added
            Selection selection = window.Selection;

            // Check if there's any selected shape before grouping
            if (selection.Count == 0) return null;

            // Group the selection
            Shape groupShape = selection.Group();

            return groupShape;
        }



        //private void ApplyAutoLayout(Page page)
        //{
        //    // Layout solution might vary depending on the specific requirements
        //    page.Layout(
        //        Microsoft.Office.Interop.Visio.layout   VisLayoutStyles .visLORouteCenter,
        //        Microsoft.Office.Interop.Visio.VisLayoutHorzAlignTypes.visLOHAlignCenter,
        //        Microsoft.Office.Interop.Visio.VisLayoutLineAdjustFrom.visLOLineAdjustFromAll,
        //        Microsoft.Office.Interop.Visio.VisLayoutLineAdjustTo.visLOLineAdjustToAll,
        //        Microsoft.Office.Interop.Visio.VisLayoutLineRouteExt.visLORouteExtStraight,
        //        0.125, 0.125, true
        //    );

        //}
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