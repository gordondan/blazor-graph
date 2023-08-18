using Serilog.Events;
using Serilog.Formatting;

namespace BlazorGraph
{
    public class CustomFileFormatter : ITextFormatter
    {
        public void Format(LogEvent logEvent, TextWriter output)
        {
            output.Write(logEvent.RenderMessage());
            output.WriteLine(); 
        }
    }
}