using System.Collections.Generic;

namespace BlazorGraph
{
    public class GraphNode
    {
        public string ComponentName { get; set; }

        public List<GraphNode> Parents { get; set; } = new List<GraphNode>();
        public List<GraphNode> Children { get; set; } = new List<GraphNode>();

        public double X { get; set; }
        public double Y { get; set; }

        public GraphNode(string componentName)
        {
            ComponentName = componentName;
        }

        // Add a child to this node and set this node as the parent for the child
        public void AddChild(GraphNode child)
        {
            if (!Children.Contains(child))
            {
                Children.Add(child);
            }

            if (!child.Parents.Contains(this))
            {
                child.Parents.Add(this);
            }
        }

        // Add a parent to this node and set this node as the child for the parent
        public void AddParent(GraphNode parent)
        {
            if (!Parents.Contains(parent))
            {
                Parents.Add(parent);
            }

            if (!parent.Children.Contains(this))
            {
                parent.Children.Add(this);
            }
        }
    }
}
