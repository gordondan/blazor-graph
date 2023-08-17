using Microsoft.Office.Interop.Visio;
using Serilog;
using System.Runtime.InteropServices;

namespace BlazorGraph
{
    public class VisioDiagramGenerator
    {
        private AppSettings _appSettings;
        private List<List<string>> _grid = new List<List<string>>();

        private HashSet<string> processedNodes = new HashSet<string>();
        private List<GraphNode> graphNodes = new List<GraphNode>();
        private Dictionary<GraphNode, (double X, double Y)> nodePositions = new Dictionary<GraphNode, (double X, double Y)>();

        private double init_y => _appSettings.VisioConfig.InitY;
        private double headerHeight => _appSettings.VisioConfig.HeaderHeight;
        private double x_offset => _appSettings.VisioConfig.HorizontalPageOffset; 
        private double y_offset => _appSettings.VisioConfig.VerticalPageOffset;  
        private int cardsPerRow => _appSettings.VisioConfig.CardsPerRow;
        private int rowsPerPage => _appSettings.VisioConfig.RowsPerPage;
        private double pageWidth => _appSettings.VisioConfig.PageWidth;
        private double pageHeight => _appSettings.VisioConfig.PageHeight;
        private double horizontalPageMargin => _appSettings.VisioConfig.HorizontalPageMargin;
        private double verticalPageMargin => _appSettings.VisioConfig.VerticalPageMargin;
        private double horizontalMargin => _appSettings.VisioConfig.HorizontalMargin;
        private double verticalMargin => _appSettings.VisioConfig.VerticalMargin;
        private double availableDrawingWidth => _appSettings.VisioConfig.AvailableDrawingWidth;
        private double availableDrawingHeight => _appSettings.VisioConfig.AvailableDrawingHeight;
        private double maxCardWidth => _appSettings.VisioConfig.MaxCardWidth;
        private double maxCardHeight => _appSettings.VisioConfig.MaxCardHeight;
        private double cardWidth => _appSettings.VisioConfig.CardWidth;
        private double cardHeight => _appSettings.VisioConfig.CardHeight;
        private double horizontalPageOffset => _appSettings.VisioConfig.HorizontalPageOffset;
        private double verticalPageOffset => _appSettings.VisioConfig.VerticalPageOffset;


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
            List<Shape> createdShapes = new List<Shape>();

            var positionedNodes = GetPositionsForNodes(graphNodes);
            ConsoleWriteGrid(positionedNodes);

            Dictionary<int, List<Shape>> shapesByRow = new Dictionary<int, List<Shape>>();

            foreach (var positionedNode in positionedNodes)
            {
                Log.Verbose($"Processing: {positionedNode.ComponentName} at X: {positionedNode.X}, Y: {positionedNode.Y}");

                Shape currentShape = CreateShape(page, positionedNode);
                createdShapes.Add(currentShape);

                int row = (int)Math.Floor(positionedNode.Y);
                if (!shapesByRow.ContainsKey(row))
                    shapesByRow[row] = new List<Shape>();

                shapesByRow[row].Add(currentShape);
            }

            foreach (var node in graphNodes)
            {
                Shape sourceShape = GetShapeByName(page, node.ComponentName);
                if (sourceShape == null)
                {
                    Log.Verbose($"No shape found for ComponentName: {node.ComponentName}");
                }
                foreach (var relatedNode in node.Children)
                {
                    Shape targetShape = GetShapeByName(page, relatedNode.ComponentName);
                    ConnectShapes(sourceShape, targetShape, page);
                }
            }
        }

        private List<GraphNode> LoadGrid(Dictionary<string, List<string>> componentRelations)
        {
            List<GraphNode> graphNodes = new List<GraphNode>();
            HashSet<string> distinctNodes = new HashSet<string>();
            List<string> orderedNodes = ComponentGraphProcessor.PreprocessRelationsBFS(componentRelations);

            foreach (var rootNode in orderedNodes)
            {
                Queue<string> nodesToProcess = new Queue<string>();
                nodesToProcess.Enqueue(rootNode);

                while (nodesToProcess.Count > 0)
                {
                    string currentComponent = nodesToProcess.Dequeue();

                    // Avoid adding duplicate nodes
                    if (distinctNodes.Contains(currentComponent))
                        continue;

                    distinctNodes.Add(currentComponent);

                    GraphNode node = graphNodes.FirstOrDefault(n => n.ComponentName == currentComponent);
                    if (node == null)
                    {
                        node = new GraphNode(currentComponent);
                        graphNodes.Add(node);
                    }

                    _grid = ComponentGraphProcessor.AddNodeToGrid(currentComponent, componentRelations, _grid);

                    if (componentRelations.ContainsKey(currentComponent))
                    {
                        foreach (string childComponent in componentRelations[currentComponent])
                        {
                            nodesToProcess.Enqueue(childComponent);

                            GraphNode childNode = graphNodes.FirstOrDefault(n => n.ComponentName == childComponent);
                            if (childNode == null)
                            {
                                childNode = new GraphNode(childComponent);
                                graphNodes.Add(childNode);
                            }

                            node.AddChild(childNode);          // Add childNode as a child to the current node
                            childNode.AddParent(node);         // Mark the current node as a parent of childNode
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
                Log.Verbose($"Warning: Attempted to connect a null shape. Shape1: {shape1?.Text}, Shape2: {shape2?.Text}");
                return;
            }
            if (ComponentGraphProcessor.IsStateComponent(shape1.Text) || ComponentGraphProcessor.IsStateComponent(shape2.Text))
            {
                return;
            }

            // Get current count before the AutoConnect
            int beforeConnectCount = page.Shapes.Count;
            Documents docs = page.Application.Documents;
            for (int i = 1; i <= docs.Count; i++)
            {
                Document doc = docs[i];
                Log.Verbose(doc.FullName);
            }

            shape1.AutoConnect(shape2, VisAutoConnectDir.visAutoConnectDirNone);
        }

        private Shape CreateShape(Page page, GraphNode node)
        {
            var x = node.X;
            var y = node.Y;
            var componentName = node.ComponentName;

            Log.Verbose($"Creating shape for: {componentName} at X: {x}, Y: {y}");

            Shape header = CreateHeader(page, x, y, componentName);
            Shape body = CreateBody(page, x, y - headerHeight);
            LabelDependencies(body, node);
            // Get or create the temporary layer
            Layer tempLayer;
            try
            {
                tempLayer = page.Layers["TempLayer"];
            }
            catch
            {
                tempLayer = page.Layers.Add("TempLayer");
            }

            // Add shapes to the temporary layer
            tempLayer.Add(header, 0);  // 0 means do not delete shape if layer is deleted
            tempLayer.Add(body, 0);

            // Select all shapes on the temporary layer
            var selection = page.CreateSelection(VisSelectionTypes.visSelTypeByLayer, VisSelectMode.visSelModeOnlySuper | VisSelectMode.visSelModeOnlySub, tempLayer);

            // Group the selected shapes
            var groupedShape = selection.Group();

            // Remove shapes from the temporary layer (this doesn't delete the shapes)
            tempLayer.Remove(header, 0);
            tempLayer.Remove(body, 0);

            return groupedShape;
        }

        private Shape CreateHeader(Page page, double x, double y, string componentName)
        {
            Log.Verbose($"Creating Header ({componentName}): {x} {y} {x + cardWidth} {y - headerHeight}");
            Shape header = page.DrawRectangle(x, y, x + cardWidth, y - headerHeight);  // Using headerHeight for the header
            header.Text = componentName;

            // Setting shape rounding for rounded rectangle
            //header.CellsU["Rounding"].ResultIU = 0.1;

            SetShapeColor(header, componentName);
            return header;
        }

        private Shape CreateBody(Page page, double x, double y)
        {
            Log.Verbose($"Creating Body {x} {y} {x + cardWidth} {y - cardHeight}");
            Shape body = page.DrawRectangle(x, y, x + cardWidth, y - cardHeight);
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
                Log.Verbose($"Error setting color for component {componentName}: {ex.Message}");
            }
        }

        private void LabelDependencies(Shape body, GraphNode node)
        {
            int stateDeps = node.Children.Where(x => ComponentGraphProcessor.IsStateComponent(x.ComponentName)).Count();
            stateDeps += node.Parents.Where(x => ComponentGraphProcessor.IsStateComponent(x.ComponentName)).Count(); ;
            int outgoingDeps = node.Children.Count();
            int incomingDeps = node.Parents.Count();

            body.Text = $"State Dep: {stateDeps}\nOutgoing: {outgoingDeps}\nIncoming: {incomingDeps}";
        }

        private void EnsurePageSize(Page page, double maxX, double maxY)
        {
            // Ensure height
            if (maxY + verticalPageMargin > page.PageSheet.CellsU["PageHeight"].ResultIU)
            {
                page.PageSheet.CellsU["PageHeight"].ResultIU = maxY + verticalPageMargin;
            }

            // Ensure width
            if (maxX + horizontalPageMargin > page.PageSheet.CellsU["PageWidth"].ResultIU)
            {
                page.PageSheet.CellsU["PageWidth"].ResultIU = maxX + horizontalPageMargin;
            }
        }

        private Shape GetShapeByName(Page page, string name)
        {
            foreach (Shape shape in page.Shapes)
            {
                Shape foundShape = SearchShapeByName(shape, name);
                if (foundShape != null)
                {
                    return foundShape;
                }
            }
            return null;
        }

        private Shape SearchShapeByName(Shape shape, string name)
        {
            if (shape.Text == name)
            {
                return shape;
            }

            if (shape.Type == (short)VisShapeTypes.visTypeGroup)
            {
                // It's a grouped shape, so search through its contained shapes.
                foreach (Shape subShape in shape.Shapes)
                {
                    Shape foundShape = SearchShapeByName(subShape, name);
                    if (foundShape != null)
                    {
                        return foundShape;
                    }
                }
            }

            return null;
        }

        private List<GraphNode> GetPositionsForNodes(List<GraphNode> graphNodes)
        {
            return graphNodes.Select(node => GetPositionForNode(node, _grid)).ToList();
        }

        private GraphNode GetPositionForNode(GraphNode node, List<List<string>> nodeGrid)
        {
            var (rowIndex, columnIndex) = FindNodeInGrid(node, nodeGrid);

            if (rowIndex != -1 && columnIndex != -1)
            {
                double x = CalculateXPosition(columnIndex);
                if (x + cardWidth > pageWidth)
                {
                    x += horizontalPageOffset;
                    rowIndex += 1;  // moving to the next row as it's a new page now
                }


                double y = CalculateYPosition(rowIndex);

                // Create a new node and return it.
                GraphNode newNode = new GraphNode(node.ComponentName)
                {
                    X = x,
                    Y = y,
                    Children = new List<GraphNode>(node.Children),
                    Parents = new List<GraphNode>(node.Parents)
                };

                return newNode;
            }

            return node;  // If there's no change in the position, return the original node.
        }

        private (int rowIndex, int columnIndex) FindNodeInGrid(GraphNode node, List<List<string>> nodeGrid)
        {
            for (int i = 0; i < nodeGrid.Count; i++)
            {
                int columnIndex = nodeGrid[i].IndexOf(node.ComponentName);
                if (columnIndex != -1)
                {
                    return (i, columnIndex);
                }
            }
            return (-1, -1);
        }

        private double CalculateXPosition(int columnIndex)
        {
            // Calculate initial X position within the page without considering page transitions.
            double x = columnIndex * (cardWidth + horizontalMargin);

            // Determine the horizontal page number.
            int horizontalPageNumber = columnIndex / cardsPerRow;

            // Adjust the X position for the nodes on subsequent pages.
            x -= horizontalPageNumber * cardsPerRow * (cardWidth + horizontalMargin);

            // Adjust the position to start from the left margin of the correct page.
            x += horizontalPageMargin + (horizontalPageNumber * availableDrawingWidth);

            return x;
        }



        private double CalculateYPosition(int rowIndex)
        {
            int verticalPageNumber = rowIndex / rowsPerPage; // Gives vertical page (top to bottom)
            int localRowIndex = rowIndex % rowsPerPage; // Position within the current vertical page

            double y = init_y
                       - localRowIndex * (cardHeight + verticalMargin)
                       - verticalPageNumber * (rowsPerPage * (cardHeight + verticalMargin) + verticalPageMargin);

            return y;
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
                window.Select(shape, (short)VisSelectArgs.visSelect);
            }

            // The active window's selection should now contain the shapes you added
            Selection selection = window.Selection;

            // Check if there's any selected shape before grouping
            if (selection.Count == 0) return null;

            // Group the selection
            Shape groupShape = selection.Group();

            return groupShape;
        }

        private void ConsoleWriteGrid(List<GraphNode> graphNodes)
        {
            foreach (var node in graphNodes)
            {
                Log.Verbose($"Node: {node.ComponentName}, X: {node.X}, Y: {node.Y}");
                foreach (var related in node.Children)
                {
                    Log.Verbose($"\tRelated: {related.ComponentName}");
                }
            }
        }
    }
}