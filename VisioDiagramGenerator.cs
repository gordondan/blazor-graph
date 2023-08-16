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
        private double x_offset = 3;

        private HashSet<string> processedNodes = new HashSet<string>();

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

            foreach (var component in componentRelations)
            {
                ProcessComponent(component, page, componentRelations);
            }

            application.ActiveWindow.ViewFit = (short)VisWindowFit.visFitPage;
            application.ActiveWindow.Zoom = 1; // Sets zoom to 100%
            application.ActiveWindow.SetViewRect(0, page.PageSheet.CellsU["PageHeight"].ResultIU,0,0);

            // Save the document to the specified file in AppSettings
            string currentDirectory = Directory.GetCurrentDirectory();
            string fullPath = System.IO.Path.Combine(currentDirectory, _appSettings.VisioFileName);
            document.SaveAs(fullPath);

            // Close the document and Visio application
            document.Close();
            application.Quit();
        }

        private void ProcessComponent(KeyValuePair<string, List<string>> component, Page page, Dictionary<string, List<string>> componentRelations, int depth = 0)
        {
            Console.WriteLine($"Processing: {component.Key} at X: {currentX}, Y: {currentY}");

            EnsurePageSize(page);

            string parentName = component.Key;

            if (processedNodes.Contains(parentName))
                return;  // skip nodes that have already been processed

            processedNodes.Add(parentName);

            Shape parentShape = CreateShape(page, parentName);

            depth += 1; // Increase the depth for child components

            double originalX = currentX; // remember the starting X position
            double originalY = currentY; // remember the starting Y position

            foreach (var relatedComponent in component.Value)
            {
                currentY += y_offset * depth; // Move down by 'depth' times the 'y_offset' for each child

                // If the relatedComponent has its own children, process them
                if (componentRelations.ContainsKey(relatedComponent))
                {
                    var childComponentPair = new KeyValuePair<string, List<string>>(relatedComponent, componentRelations[relatedComponent]);
                    ProcessComponent(childComponentPair, page, componentRelations, depth);
                }
                else
                {// Inside the foreach loop in ProcessComponent
                    Shape childShape = GetShapeByName(page, relatedComponent);
                    if (childShape == null)
                    {
                        childShape = CreateShape(page, relatedComponent);
                    }

                    ConnectShapes(parentShape, childShape, page);
                }

                currentX += x_offset; // Move to the right for the next child (or sibling)
            }

            currentX = originalX; // reset the X position back to the parent's X position
            currentY = originalY; // reset the Y position back to the parent's Y position
        }

        private void EnsurePageSize(Page page)
        {
            const double margin = 2; // some space on all sides

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
                page.PageSheet.CellsU["PageWidth"].ResultIU = currentX + 2 + margin;
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
    
        private bool IsVendorComponent(string componentName)
        {
            return _appSettings.Vendors.Any(vendor => componentName.Contains(vendor));
        }
    }
}
