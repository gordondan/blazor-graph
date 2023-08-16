using System.Collections.Generic;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Visio;
using static System.Net.Mime.MediaTypeNames;

namespace BlazorGraph
{
    public class VisioDiagramGenerator
    {
        private AppSettings _appSettings;
        private double currentX = .5;
        private double currentY = 10;
        private const double init_y = 10;
        private double y_offset = -1.25;
        private double x_offset = 2.25;

        private HashSet<string> processedNodes = new HashSet<string>();
        private HashSet<string> stateNodes = new HashSet<string>();

        public VisioDiagramGenerator(AppSettings appSettings)
        {
            _appSettings = appSettings;
        }

        public void GenerateVisioDiagram(Dictionary<string, List<string>> componentRelations)
        {
            var application = new Microsoft.Office.Interop.Visio.Application();
            application.Visible = true;
            var document = application.Documents.Add("");
            var page = application.ActivePage;

            var rootNodes = GetRootNodes(componentRelations);

            foreach (var rootNode in rootNodes)
            {
                ProcessComponent(new KeyValuePair<string, List<string>>(rootNode, componentRelations[rootNode]), page, componentRelations, isRoot: true);
            }

            foreach (var component in componentRelations)
            {
                if (!rootNodes.Contains(component.Key)) // This check ensures root nodes are not processed again.
                    ProcessComponent(component, page, componentRelations);
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

        private void ProcessComponent(KeyValuePair<string, List<string>> component, Page page, Dictionary<string, List<string>> componentRelations, int depth = 0, bool isRoot = false)
        {
            Console.WriteLine($"Processing: {component.Key} at X: {currentX}, Y: {currentY}");

            EnsurePageSize(page);

            string nodeName = component.Key;

            // Store original positions
            double originalX = currentX;
            double originalY = currentY;

            // Check if the current node is a State node
            if (nodeName.EndsWith("State") && !stateNodes.Contains(nodeName))
            {
                stateNodes.Add(nodeName);
                currentY = init_y + (1.5 * y_offset);  // Two row heights above the root node row

                if (stateNodes.Count > 1)
                {
                    currentX = x_offset * stateNodes.Count();  // Move to the right for subsequent state nodes
                }
                else
                {
                    currentX = x_offset;  // For the first state node, be one x_offset width in
                }
            }
            if (processedNodes.Contains(nodeName))
                return;  // skip nodes that have already been processed

            processedNodes.Add(nodeName);



            Shape parentShape = CreateShape(page, nodeName);

            foreach (var relatedComponent in component.Value)
            {
                if (isRoot) // For direct descendants of a root node
                {
                    currentY = originalY + y_offset;
                }
                else
                {
                    currentY += y_offset * depth; // For deeper descendants
                }
                // Check if the current node is a State node
                if (relatedComponent.EndsWith("State") && !stateNodes.Contains(nodeName))
                {
                    stateNodes.Add(nodeName);
                    currentY = init_y + (1.5 * y_offset);  // Two row heights above the root node row
                    if (stateNodes.Count > 1)
                    {
                        currentX = x_offset * stateNodes.Count();  // Move to the right for subsequent state nodes
                    }
                    else
                    {
                        currentX = x_offset;  // For the first state node, be one x_offset width in
                    }
                }
                if (componentRelations.ContainsKey(relatedComponent))
                {
                    var childComponentPair = new KeyValuePair<string, List<string>>(relatedComponent, componentRelations[relatedComponent]);
                    ProcessComponent(childComponentPair, page, componentRelations, depth + 1);
                }
                else
                {
                    Shape childShape = GetShapeByName(page, relatedComponent);
                    if (childShape == null)
                    {
                        childShape = CreateShape(page, relatedComponent);
                    }

                    ConnectShapes(parentShape, childShape, page);
                }

                currentX += x_offset; // Move to the right for the next child (or sibling)
            }

            if (isRoot)
            {
                currentY = originalY; // Only reset the Y position for root nodes (their siblings should be on the same level)
            }

            if (stateNodes.Contains(nodeName))
            {
                currentX = originalX;
                currentY = originalY;
            }
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
            if (IsVendorComponent(componentName))
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

        private void ConnectShapes(Shape shape1, Shape shape2, Page page)
        {

            shape1.AutoConnect(shape2, VisAutoConnectDir.visAutoConnectDirNone);
        }
        private HashSet<string> GetRootNodes(Dictionary<string, List<string>> componentRelations)
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

        private bool IsVendorComponent(string componentName)
        {
            return _appSettings.Vendors.Any(vendor => componentName.Contains(vendor));
        }
    }
}
