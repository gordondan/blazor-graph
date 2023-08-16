﻿using Microsoft.Office.Interop.Visio;
using System.Runtime.InteropServices;

namespace BlazorGraph
{
    public class VisioDiagramGenerator
    {
        private const double init_y = 10;
        private AppSettings _appSettings;
        private double currentX = .5;
        private double currentY = 10;
        private List<List<string>> grid = new List<List<string>>();
        private HashSet<string> processedNodes = new HashSet<string>();
        private HashSet<string> stateNodes = new HashSet<string>();
        private double x_offset = 2.25;
        private double y_offset = -1.25;
        private ComponentGraphProcessor _graphProcessor;

        public VisioDiagramGenerator(AppSettings appSettings)
        {
            _appSettings = appSettings;
        }

        public void GenerateVisioDiagram(Dictionary<string, List<string>> componentRelations)
        {
            var application = new Microsoft.Office.Interop.Visio.Application();
            _graphProcessor = new ComponentGraphProcessor(componentRelations, _appSettings);
            application.Visible = true;
            var document = application.Documents.Add("");
            var page = application.ActivePage;

            var rootNodes = _graphProcessor.GetRootNodes(componentRelations);
            List<string> orderedNodes = _graphProcessor.PreprocessRelationsBFS(componentRelations);

            foreach (var rootNode in orderedNodes)
            {
                ProcessComponent(rootNode, page, componentRelations, isRoot: true);
            }

            application.ActiveWindow.ViewFit = (short)VisWindowFit.visFitPage;
            application.ActiveWindow.Zoom = 1; // Sets zoom to 100%

            // Save the document to the specified file in AppSettings
            string currentDirectory = Directory.GetCurrentDirectory();
            string fullPath = System.IO.Path.Combine(currentDirectory, _appSettings.VisioFileName);
            document.SaveAs(fullPath);

            // Close the document and Visio application
            document.Close();
            application.Quit();
        }
        private void ProcessComponent(string rootNode, Page page, Dictionary<string, List<string>> componentRelations, bool isRoot)
        {
            Queue<string> nodesToProcess = new Queue<string>();
            nodesToProcess.Enqueue(rootNode);

            while (nodesToProcess.Count > 0)
            {
                string currentComponent = nodesToProcess.Dequeue();

                // Decide the position
                _graphProcessor.AddNodeToGrid(currentComponent, componentRelations);

                SetPositionFromGrid(currentComponent);

                Console.WriteLine($"Processing: {currentComponent} at X: {currentX}, Y: {currentY}");

                EnsurePageSize(page);

                if (!processedNodes.Contains(currentComponent))
                {
                    Shape currentShape = CreateShape(page, currentComponent);
                    processedNodes.Add(currentComponent);

                    if (componentRelations.ContainsKey(currentComponent))
                    {
                        foreach (string relatedComponent in componentRelations[currentComponent])
                        {
                            if (!processedNodes.Contains(relatedComponent))
                            {
                                nodesToProcess.Enqueue(relatedComponent);
                            }

                            // Connecting nodes
                            Shape relatedShape = GetShapeByName(page, relatedComponent);
                            if (relatedShape != null)
                            {
                                ConnectShapes(currentShape, relatedShape, page);
                            }
                        }
                    }
                }
            }
        }

        private void SetPositionFromGrid(string nodeName)
        {
            int rowIndex = -1;
            int columnIndex = -1;

            for (int i = 0; i < grid.Count; i++)
            {
                columnIndex = grid[i].IndexOf(nodeName);
                if (columnIndex != -1)
                {
                    rowIndex = i;
                    break;
                }
            }

            if (rowIndex != -1 && columnIndex != -1)
            {
                currentX = columnIndex * x_offset + 0.5;  // added offset
                currentY = init_y - rowIndex * y_offset;
            }
        }
        private void ConnectShapes(Shape shape1, Shape shape2, Page page)
        {
            shape1.AutoConnect(shape2, VisAutoConnectDir.visAutoConnectDirNone);
        }

        private Shape CreateShape(Page page, string componentName)
        {
            Console.WriteLine($"Creating shape for: {componentName} at X: {currentX}, Y: {currentY}");

            EnsurePageSize(page);

            Shape shape = page.DrawRectangle(currentX, currentY, currentX + 2, currentY + 1);
            shape.Text = componentName;

            // Setting shape rounding for rounded rectangle
            shape.CellsU["Rounding"].ResultIU = 0.25;  // Adjust the value as necessary for the desired rounding

            // Default to navy blue with white text
            shape.CellsU["FillForegnd"].FormulaU = "RGB(0, 0, 128)";
            shape.CellsU["Char.Color"].FormulaU = "RGB(255, 255, 255)";  // White color for text

            // If it's a vendor component, set to lime green with white text
            if (_graphProcessor.IsVendorComponent(componentName))
            {
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

        private void UpdatePositionForRootNode(double originalY, bool isRoot)
        {
            if (isRoot)
            {
                currentY = originalY + y_offset;
            }
            else
            {
                currentY += y_offset;
            }
        }
        private void UpdatePositionForStateNode(string nodeName)
        {
            if (!nodeName.EndsWith("State") || stateNodes.Contains(nodeName))
                return;

            stateNodes.Add(nodeName);
            currentY = init_y - (1.5 * y_offset);  // Two row heights above the root node row
            currentX = stateNodes.Count > 1 ? x_offset * stateNodes.Count() : x_offset;
        }


    }
}