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
        private double currentY = 20;
        private double y_offset = -1;
        private double x_offset = 2.25;
        public VisioDiagramGenerator(AppSettings appSettings)
        {
            _appSettings = appSettings;
        }

        public void GenerateVisioDiagram(Dictionary<string, List<string>> componentRelations)
        {
            var application = new Microsoft.Office.Interop.Visio.Application();
            application.Visible = false;
            var document = application.Documents.Add("");
            var page = application.ActivePage;

            foreach (var component in componentRelations)
            {
                ProcessComponent(component, page);
            }

            // Save the document to the specified file in AppSettings
            string currentDirectory = Directory.GetCurrentDirectory();
            string fullPath = System.IO.Path.Combine(currentDirectory, _appSettings.VisioFileName);
            document.SaveAs(fullPath);


            // Close the document and Visio application
            document.Close();
            application.Quit();
        }

        private void ProcessComponent(KeyValuePair<string, List<string>> component, Page page)
        {
            EnsurePageSize(page);

            string parentName = component.Key;
            Shape parentShape = CreateShape(page, parentName);

            // Save current x value and adjust y for children
            double originalX = currentX;
            currentY += y_offset;

            foreach (var relatedComponent in component.Value)
            {
                Shape childShape = CreateShape(page, relatedComponent);
                ConnectShapes(parentShape, childShape, page);
            }

            // Reset y for next parent and set x to where it was after processing the parent
            currentY -= y_offset;
            currentX = originalX;
        }

        private void EnsurePageSize(Page page)
        {
            const double margin = 2; // some space on all sides

            if (currentY < margin)
            {
                page.PageSheet.CellsU["PageHeight"].ResultIU += Math.Abs(currentY) + margin;
                currentY += Math.Abs(currentY) + margin;
            }
        }


        private Shape CreateShape(Page page, string componentName)
        {
            Shape shape = page.DrawRectangle(currentX, currentY, currentX + 2, currentY + 1);
            shape.Text = componentName;

            // Increment the y-value for the next shape. Adjust this value as needed for spacing.
            currentY += y_offset;

            // Add color to vendor components
            if (IsVendorComponent(componentName))
            {
                // Assuming VendorComponentColor is in the format "RGB(255,0,0)" for red.
                try
                {
                    //shape.CellsU["FillForegnd"].FormulaU = _appSettings.VendorComponentColor;
                    shape.CellsU["FillForegnd"].FormulaU = "RGB(0, 0, 128)";  // This should turn the shape red

                }
                catch (COMException ex)
                {
                    // Handle the error, perhaps logging it or notifying the user
                    Console.WriteLine($"Error setting color for component {componentName}: {ex.Message}");
                }
            }


            return shape;
        }

        //private void ConnectShapes(Shape parentShape, Shape childShape, Page page)
        //{
        //    // Ensure the shapes are valid before attempting to connect them.
        //    if (parentShape == null || childShape == null)
        //    {
        //        throw new ArgumentNullException("One or both of the shapes are null.");
        //    }

        //    var flowchartStencil = page.Application.Documents.OpenEx("BASFLO_M.VSSX", (short)VisOpenSaveArgs.visOpenRO + (short)VisOpenSaveArgs.visOpenDocked);
        //    foreach (Master master in flowchartStencil.Masters)
        //    {
        //        Console.WriteLine(master.Name);
        //    }
        //    var connectorMaster = flowchartStencil.Masters["Dynamic Connector"];



        //    var connector = page.Drop(connectorMaster, 0, 0);

        //    connector.CellsU["BeginX"].GlueTo(parentShape.CellsU["PinX"]);
        //    connector.CellsU["BeginY"].GlueTo(parentShape.CellsU["PinY"]);
        //    connector.CellsU["EndX"].GlueTo(childShape.CellsU["PinX"]);
        //    connector.CellsU["EndY"].GlueTo(childShape.CellsU["PinY"]);

        //}
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
