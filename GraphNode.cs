using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BlazorGraph
{
    public class GraphNode
    {
        public string ComponentName { get; set; }
        public List<GraphNode> RelatedComponents { get; set; } = new List<GraphNode>();
        public double X { get; set; }
        public double Y { get; set; }

        public GraphNode(string componentName)
        {
            ComponentName = componentName;
        }

        public void AddRelatedComponent(GraphNode component)
        {
            RelatedComponents.Add(component);
        }
    }

}
