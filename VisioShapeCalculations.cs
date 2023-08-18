using Microsoft.Office.Interop.Visio;
using Serilog;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace BlazorGraph
{
    public partial class VisioDiagramGenerator
    {
        private Shape CreateShape(Page page, GraphNode node)
        {
            var x = node.X;
            var y = node.Y;
            var componentName = node.ComponentName;

            Log.Debug($"Creating shape for: {componentName} at X: {x}, Y: {y}");

            Log.Debug($"Creating a header for component '{componentName}' at position X: {x}, Y: {y}. Card Height: {cardHeight}");
            Shape header = CreateHeader(page, x, y, componentName);

            Log.Debug($"Creating a body for component '{componentName}' directly below the header at position X: {x}, Y: {y - headerHeight}.");
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
            Log.Debug($"Creating Body {x} {y} {x + cardWidth} {y - cardHeight + headerHeight}");
            Shape body = page.DrawRectangle(x, y, x + cardWidth, y - cardHeight + headerHeight);
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
        private GraphNode ShiftAndLog(GraphNode node, double xShift, double yShift)
        {
            var shiftedNode = node with { X = xShift + node.X, Y = yShift + node.Y };
            Log.Verbose($"Original X: {node.X}, Original Y: {node.Y}, Shifted X: {shiftedNode.X} ({xShift}), Shifted Y: {shiftedNode.Y} ({yShift})");
            return shiftedNode;
        }

        private List<GraphNode> ShiftNodesRightAndUp(List<GraphNode> positionedNodes, double xShift, double yShift)
        {
            return positionedNodes.Select(node => ShiftAndLog(node, xShift, yShift)).ToList();
        }


        private List<GraphNode> GetPositionsForNodes(List<GraphNode> graphNodes)
        {
            return graphNodes.Select(node => GetPositionForNode(node, _grid)).ToList();
        }

        private GraphNode GetPositionForNode(GraphNode node, List<List<string>> nodeGrid)
        {
            Log.Verbose($"{Environment.NewLine}--------------------------------------------{Environment.NewLine}", node.ComponentName);
            Log.Verbose("Getting position for node: {NodeName}", node.ComponentName);

            var (rowIndex, columnIndex) = FindNodeInGrid(node, nodeGrid);
            Log.Verbose("Node found in grid at Row: {RowIndex}, Column: {ColumnIndex}", rowIndex, columnIndex);

            if (rowIndex != -1 && columnIndex != -1)
            {
                double x = CalculateXPosition(columnIndex);
                Log.Verbose("Calculated X position: {CalculatedX}", x);

                double y = CalculateYPosition(rowIndex);
                Log.Verbose("Calculated Y position: {CalculatedY}", y);

                // Create a new node and return it.
                GraphNode newNode = new GraphNode(node.ComponentName)
                {
                    X = x,
                    Y = y,
                    Children = new List<GraphNode>(node.Children),
                    Parents = new List<GraphNode>(node.Parents)
                };

                Log.Verbose("Created new node with updated positions: {NewNodeName} at X: {NewNodeX}, Y: {NewNodeY}",
                    newNode.ComponentName, newNode.X, newNode.Y);

                return newNode;
            }

            Log.Verbose("No position change detected for node: {NodeName}. Returning original node.{return}", node.ComponentName, Environment.NewLine);
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
            Log.Verbose("Calculating X position for column index {ColumnIndex}", columnIndex);

            // Determine the horizontal page number.
            int horizontalPageNumber = columnIndex / cardsPerRow;
            Log.Verbose("Determined horizontal page number: {HorizontalPageNumber}", horizontalPageNumber);

            // Calculate layoutColumnIndex, which represents the position of the card within its current page.
            int layoutColumnIndex = columnIndex % cardsPerRow;
            Log.Verbose($"Layout Column Index: {layoutColumnIndex}");

            // Calculate the initial X position for this card, using the layoutColumnIndex.
            double x = layoutColumnIndex * (cardWidth + horizontalMargin);
            Log.Verbose($"Initial calculated X position (within current page): {x}");

            // Now adjust for the entire width of the previous pages the card has spanned.
            x += horizontalPageNumber * effectivePageWidth;
            Log.Verbose($"X position after adding width of previous pages: {x}");

            // Finally, adjust the position to start from the left margin of the current page.
            x += horizontalPageMargin;
            Log.Verbose($" (considering left margin): {x}");

            //x += (cardWidth / 2);  // Adjust for center-center pin position
            //Log.Verbose($"X position (considering center pin pos cardWidth/2={cardWidth / 2}): {x}");
            return x;
        }

        private double CalculateYPosition(int rowIndex)
        {
            Log.Verbose("-----------------------------------------------------------------------------");
            Log.Verbose("Calculating Y position for row index {RowIndex}", rowIndex);

            // Determine the vertical page number.
            int verticalPageNumber = rowIndex / rowsPerPage;
            Log.Verbose("Determined vertical page number: {VerticalPageNumber}", verticalPageNumber);

            // Calculate localRowIndex, which represents the position of the card within its current page.
            int localRowIndex = rowIndex % rowsPerPage;
            Log.Verbose($"Local Row Index: {localRowIndex}");

            // Calculate the initial Y position for this card, using the localRowIndex.
            double y = init_y - localRowIndex * (cardHeight + verticalMargin);
            Log.Verbose($"Initial calculated Y position (within current page): {y}");

            // Now adjust for the entire height of the previous pages the card has spanned.
            y -= verticalPageNumber * effectivePageHeight;
            Log.Verbose($"Y position after subtracting height ({verticalPageNumber * effectivePageHeight}) of previous pages: {y}");

            // Finally, adjust the position to start from the top margin of the current page.
            y -= verticalPageMargin;
            Log.Verbose($"Y position (considering top margin): {y}");

            y -= (cardHeight / 2);  // Adjust for center-center pin position
            Log.Verbose($"Y position (considering center pin pos cardHeight/2={cardHeight / 2}): {y}");

            Log.Verbose("-----------------------------------------------------------------------------");
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
