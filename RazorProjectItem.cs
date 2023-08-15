using Microsoft.AspNetCore.Razor.Language;
using System.Text;

namespace BlazorGraph
{
    public class CustomRazorProjectItem : RazorProjectItem
    {
        private readonly RazorSourceDocument _sourceDocument;

        public CustomRazorProjectItem(RazorSourceDocument sourceDocument)
        {
            _sourceDocument = sourceDocument;
        }

        public override string BasePath => "/";
        public override string FilePath => _sourceDocument.FilePath;
        public override string PhysicalPath => null;
        public override bool Exists => true;

        public override Stream Read()
        {
            var contentBuilder = new StringBuilder();
            char[] buffer = new char[_sourceDocument.Length];
            _sourceDocument.CopyTo(0, buffer, 0, buffer.Length);
            contentBuilder.Append(buffer);
            var contentBytes = Encoding.UTF8.GetBytes(contentBuilder.ToString());
            return new MemoryStream(contentBytes);
        }
    }

}
