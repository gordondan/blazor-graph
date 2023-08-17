using Serilog.Events;
using Serilog.Formatting;

namespace BlazorGraph
{
    public class CustomConsoleFormatter : ITextFormatter
    {
        public void Format(LogEvent logEvent, TextWriter output)
        {
            output.Write(logEvent.RenderMessage());

            if (logEvent.Properties.TryGetValue("NoNewLine", out var noNewLineProperty)
                && noNewLineProperty is ScalarValue scalarValue
                && scalarValue.Value is bool noNewLine && noNewLine)
            {
                // Do not append a newline if the NoNewLine property is true.
                return;
            }

            output.WriteLine(); // Append a newline for all other cases.
        }
    }
}
