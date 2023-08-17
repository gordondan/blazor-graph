namespace BlazorGraph
{
    public class AppSettings
    {
        internal string? OutputFilePath;

        public List<string> Vendors { get; set; }
        public List<string> Skips { get; set; }
        public bool DisplayVendorComponents { get; set; }
        public string VendorComponentColor { get; set; }
        public string Directory { get; set; }
        public string StartingNode { get; set; }
        public string MermaidFileName { get; set; } = "dependencyGraph.mmd";
        public string VisioFileName { get; set; } = "dependencyGraph.vsd"; 
        public VisioDiagramConfig VisioConfig { get; set; } = new VisioDiagramConfig();

    }
}