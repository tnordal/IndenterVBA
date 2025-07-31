using System;
using VBIDE = Microsoft.Vbe.Interop;
using System.Text;

namespace IndenterVBA
{
    public class VbaIndenter
    {
        public void IndentAllModules(VBIDE.VBProject vbProject)
        {
            foreach (VBIDE.VBComponent component in vbProject.VBComponents)
            {
                if (component.CodeModule != null)
                {
                    int lineCount = component.CodeModule.CountOfLines;
                    if (lineCount > 0)
                    {
                        string code = component.CodeModule.Lines[1, lineCount];
                        string indentedCode = IndentCode(code);
                        component.CodeModule.DeleteLines(1, lineCount);
                        component.CodeModule.AddFromString(indentedCode);
                    }
                }
            }
        }

        public string IndentCode(string code)
        {
            var lines = code.Split(new[] { "\r\n", "\n" }, StringSplitOptions.None);
            int indentLevel = 0;
            var sb = new StringBuilder();

            foreach (var line in lines)
            {
                string trimmed = line.Trim();
                if (trimmed.StartsWith("End", StringComparison.OrdinalIgnoreCase) ||
                    trimmed.Equals("Next", StringComparison.OrdinalIgnoreCase) ||
                    trimmed.Equals("Loop", StringComparison.OrdinalIgnoreCase))
                {
                    indentLevel = Math.Max(0, indentLevel - 1);
                }

                sb.AppendLine(new string(' ', indentLevel * 4) + trimmed);

                if (trimmed.StartsWith("Sub ", StringComparison.OrdinalIgnoreCase) ||
                    trimmed.StartsWith("Function ", StringComparison.OrdinalIgnoreCase) ||
                    (trimmed.StartsWith("If ", StringComparison.OrdinalIgnoreCase) && trimmed.EndsWith("Then")) ||
                    trimmed.StartsWith("For ", StringComparison.OrdinalIgnoreCase) ||
                    trimmed.StartsWith("Do ", StringComparison.OrdinalIgnoreCase) ||
                    trimmed.StartsWith("Select Case", StringComparison.OrdinalIgnoreCase))
                {
                    indentLevel++;
                }
            }

            return sb.ToString();
        }
    }
}
