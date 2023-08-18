using Microsoft.Office.Interop.Visio;
using Serilog;
using System.Text;

namespace BlazorGraph
{
    public partial class VisioDiagramGenerator
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
        private double effectivePageWidth => _appSettings.VisioConfig.EffectivePageWidth;
        private double effectivePageHeight => _appSettings.VisioConfig.EffectivePageHeight;
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
            LogConfigurationValues();

            List<GraphNode> graphNodes = LoadGrid(componentRelations);         

            var positionedNodes = GetPositionsForNodes(graphNodes);
            var (maxX, maxY) = GetMaxCoordinates(positionedNodes);
            var (minX, minY) = GetMinCoordinates(positionedNodes);
            
            Log.Verbose("*******************************************************");
            positionedNodes = ShiftNodesRightAndUp(positionedNodes, minX - horizontalPageMargin, Math.Abs(minY));
            Log.Verbose("-------------------------------------------------------");
            
            (maxX, maxY) = GetMaxCoordinates(positionedNodes);
            (minX, minY) = GetMinCoordinates(positionedNodes);


            var application = new Microsoft.Office.Interop.Visio.Application();
            application.Visible = true;
            var document = application.Documents.Add("");
            var page = application.ActivePage;
            var pageSheet = page.PageSheet;

            pageSheet.get_CellsU("PageWidth").ResultIU = pageWidth;
            pageSheet.get_CellsU("PageHeight").ResultIU = pageHeight;
            EnsurePageSize(page, maxX, maxY);

            WriteToVisio(positionedNodes, page);

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

        private (double maxX, double maxY) GetMinCoordinates(List<GraphNode> graphNodes)
        {
            double minX = 100000000;
            double minY = 100000000;

            foreach (var node in graphNodes)
            {
                if (node.X < minX) minX = node.X;
                if (node.Y < minY) minY = node.Y;
            }

            return (minX, minY);
        }

        private void WriteToVisio(List<GraphNode> positionedNodes, Page page)
        {
            List<Shape> createdShapes = new List<Shape>();

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

        public static void LogGrid(Dictionary<string, List<string>> grid)
        {
            if (grid == null || grid.Count == 0)
            {
                Log.Verbose("The grid is empty.");
                return;
            }

            foreach (var entry in grid)
            {
                Log.Verbose($"Key: {entry.Key}");
                Log.Verbose("Values:");
                foreach (var value in entry.Value)
                {
                    Log.Verbose($"- {value}");
                }
                Log.Verbose("");  // Empty line for better separation between entries
            }
        }
        public static void LogGraphNodes(List<GraphNode> nodes)
        {
            if (nodes == null || !nodes.Any())
            {
                Log.Verbose("The list of nodes is empty.");
                return;
            }

            foreach (var node in nodes)
            {
                StringBuilder sb = new StringBuilder();
                sb.AppendLine($"ComponentName: {node.ComponentName}");
                sb.AppendLine($"Position - X: {node.X}, Y: {node.Y}");
                sb.AppendLine("Parents:");
                foreach (var parent in node.Parents)
                {
                    sb.AppendLine($"- {parent.ComponentName}");
                }
                sb.AppendLine("Children:");
                foreach (var child in node.Children)
                {
                    sb.AppendLine($"- {child.ComponentName}");
                }
                Log.Verbose(sb.ToString());
            }
        }
        private void LogConfigurationValues()
        {
            Log.Verbose("Logging Visio Configuration Values:");
            Log.Verbose("init_y: {InitY}", init_y);
            Log.Verbose("headerHeight: {HeaderHeight}", headerHeight);
            Log.Verbose("x_offset: {XOffset}", x_offset);
            Log.Verbose("y_offset: {YOffset}", y_offset);
            Log.Verbose("cardsPerRow: {CardsPerRow}", cardsPerRow);
            Log.Verbose("rowsPerPage: {RowsPerPage}", rowsPerPage);
            Log.Verbose("pageWidth: {PageWidth}", pageWidth);
            Log.Verbose("pageHeight: {PageHeight}", pageHeight);
            Log.Verbose("horizontalPageMargin: {HorizontalPageMargin}", horizontalPageMargin);
            Log.Verbose("verticalPageMargin: {VerticalPageMargin}", verticalPageMargin);
            Log.Verbose("horizontalMargin: {HorizontalMargin}", horizontalMargin);
            Log.Verbose("verticalMargin: {VerticalMargin}", verticalMargin);
            Log.Verbose("availableDrawingWidth: {AvailableDrawingWidth}", availableDrawingWidth);
            Log.Verbose("availableDrawingHeight: {AvailableDrawingHeight}", availableDrawingHeight);
            Log.Verbose("maxCardWidth: {MaxCardWidth}", maxCardWidth);
            Log.Verbose("maxCardHeight: {MaxCardHeight}", maxCardHeight);
            Log.Verbose("cardWidth: {CardWidth}", cardWidth);
            Log.Verbose("cardHeight: {CardHeight}", cardHeight);
            Log.Verbose("horizontalPageOffset: {HorizontalPageOffset}", horizontalPageOffset);
            Log.Verbose("verticalPageOffset: {VerticalPageOffset}", verticalPageOffset);
            Log.Verbose(Environment.NewLine);
        }
    }
}