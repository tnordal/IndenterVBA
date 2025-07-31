using System;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using System.Collections.Generic;
using VBIDE = Microsoft.Vbe.Interop;

namespace IndenterVBA
{
    public class VbaIndenter
    {
        // For simplicity, we'll directly get the active code pane
        public void IndentAllModules(VBIDE.VBProject vbProject)
        {
            // We'll ignore the project parameter and just focus on active code pane
            IndentActiveCodePane(vbProject.VBE);
        }

        private void IndentActiveCodePane(VBIDE.VBE vbe)
        {
            try
            {
                // Check if we have an active code pane
                if (vbe.ActiveCodePane == null)
                {
                    MessageBox.Show("Please open a code module first.");
                    return;
                }

                // Get the active code module
                VBIDE.CodeModule codeModule = vbe.ActiveCodePane.CodeModule;
                if (codeModule == null)
                {
                    MessageBox.Show("No active code module found.");
                    return;
                }

                // Initialize log
                StringBuilder log = new StringBuilder();
                log.AppendLine($"Module: {codeModule.Name}");
                log.AppendLine($"Total lines: {codeModule.CountOfLines}");

                // Get all code from the module
                string code = "";
                if (codeModule.CountOfLines > 0)
                {
                    code = codeModule.Lines[1, codeModule.CountOfLines];
                }

                // Create directory if needed
                Directory.CreateDirectory("C:\\Temp");

                // Log original code
                File.WriteAllText("C:\\Temp\\Original_Code.txt", code);

                // Find all Sub and Function boundaries and indent them
                log.AppendLine("\nFinding and indenting procedures:");
                string indentedCode = IndentCode(code, log);

                // Write the indented code to a log file for verification
                File.WriteAllText("C:\\Temp\\Indented_Code.txt", indentedCode);

                // Apply the indented code to the module
                if (codeModule.CountOfLines > 0)
                {
                    codeModule.DeleteLines(1, codeModule.CountOfLines);
                    codeModule.InsertLines(1, indentedCode);
                }

                // Write log
                File.WriteAllText("C:\\Temp\\VbaIndentLog.txt", log.ToString());

                MessageBox.Show("Indentation complete. Check C:\\Temp\\VbaIndentLog.txt for details.");
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error: {ex.Message}");
                File.WriteAllText("C:\\Temp\\IndentError.txt", ex.ToString());
            }
        }

        private string IndentCode(string code, StringBuilder log)
        {
            // Split code into lines
            string[] lines = code.Split(new[] { "\r\n", "\n" }, StringSplitOptions.None);
            StringBuilder indentedCode = new StringBuilder();
            
            // Find all procedure boundaries first
            List<Procedure> procedures = FindProcedureBoundaries(code, log);
            
            // Process each line and apply indentation
            for (int i = 0; i < lines.Length; i++)
            {
                // Current line number (0-based)
                int lineIndex = i;
                
                // Check if this line is inside a procedure
                Procedure currentProc = procedures.Find(p => lineIndex >= p.StartLineIndex && lineIndex <= p.EndLineIndex);
                
                if (currentProc != null)
                {
                    // Inside a procedure - apply indentation rules
                    string line = lines[lineIndex];
                    string trimmedLine = line.Trim();
                    
                    // Check if this is the procedure declaration or end
                    if (lineIndex == currentProc.StartLineIndex || lineIndex == currentProc.EndLineIndex)
                    {
                        // No indentation for procedure declaration and End Sub/Function
                        indentedCode.AppendLine(trimmedLine);
                    }
                    else if (string.IsNullOrWhiteSpace(trimmedLine))
                    {
                        // Preserve empty lines
                        indentedCode.AppendLine("");
                    }
                    else 
                    {
                        // Check if this is in the declaration block at the top of the procedure
                        bool isInDeclarationBlock = IsInDeclarationBlock(lineIndex, currentProc, lines);
                        
                        if (isInDeclarationBlock)
                        {
                            // No indentation for declarations
                            indentedCode.AppendLine(trimmedLine);
                        }
                        else
                        {
                            // This is regular code inside the procedure body - indent it
                            indentedCode.AppendLine("    " + trimmedLine);
                        }
                    }
                }
                else
                {
                    // Outside any procedure - no indentation
                    indentedCode.AppendLine(lines[lineIndex].Trim());
                }
            }
            
            return indentedCode.ToString();
        }

        private List<Procedure> FindProcedureBoundaries(string code, StringBuilder log)
        {
            List<Procedure> procedures = new List<Procedure>();
            
            // Split code into lines
            string[] lines = code.Split(new[] { "\r\n", "\n" }, StringSplitOptions.None);

            // Track procedure boundaries
            for (int i = 0; i < lines.Length; i++)
            {
                string trimmedLine = lines[i].Trim();
                
                // Check for procedure start (Sub or Function with optional Public/Private)
                Regex procStartRegex = new Regex(@"^\s*(Public\s+|Private\s+|Friend\s+)?(Sub|Function)\s+", 
                                               RegexOptions.IgnoreCase);
                
                Match match = procStartRegex.Match(trimmedLine);
                if (match.Success && !trimmedLine.StartsWith("End ", StringComparison.OrdinalIgnoreCase))
                {
                    // Found procedure start
                    string name = ExtractProcedureName(trimmedLine);
                    int startLine = i + 1; // 1-based line numbers for display
                    int startLineIndex = i; // 0-based index for array access
                    
                    log.AppendLine($"Procedure: {name}, Start Line: {startLine}");
                    log.AppendLine($"  Declaration: {trimmedLine}");

                    // Find the end of this procedure
                    int endLine = -1;
                    int endLineIndex = -1;
                    
                    for (int j = i + 1; j < lines.Length; j++)
                    {
                        string endLineText = lines[j].Trim();
                        if (endLineText.Equals("End Sub", StringComparison.OrdinalIgnoreCase) || 
                            endLineText.Equals("End Function", StringComparison.OrdinalIgnoreCase))
                        {
                            endLine = j + 1; // 1-based line numbers for display
                            endLineIndex = j; // 0-based index for array access
                            log.AppendLine($"  End Line: {endLine}");
                            
                            // Add procedure to our list
                            procedures.Add(new Procedure
                            {
                                Name = name,
                                StartLineIndex = startLineIndex,
                                EndLineIndex = endLineIndex
                            });
                            
                            break;
                        }
                    }
                    
                    // If no end was found, log that information
                    if (endLine == -1)
                    {
                        log.AppendLine($"  WARNING: No matching End Sub/Function found!");
                    }
                }
            }
            
            return procedures;
        }

        private bool IsInDeclarationBlock(int lineIndex, Procedure procedure, string[] lines)
        {
            // Declarations are at the top of the procedure
            // We consider declaration block to be consecutive declaration statements
            // right after the procedure declaration
            
            List<string> declarationKeywords = new List<string> { 
                "Dim ", "Private ", "Public ", "Static ", "Const ", "ReDim ", "Set ", "Declare " 
            };
            
            // Start from the procedure start line and go down until we hit the current line
            for (int i = procedure.StartLineIndex + 1; i < lineIndex; i++)
            {
                string trimmedLine = lines[i].Trim();
                
                // Skip empty lines and comments
                if (string.IsNullOrWhiteSpace(trimmedLine) || trimmedLine.StartsWith("'"))
                {
                    continue;
                }
                
                // Check if this line is a declaration
                bool isDeclaration = false;
                foreach (var keyword in declarationKeywords)
                {
                    if (trimmedLine.StartsWith(keyword, StringComparison.OrdinalIgnoreCase))
                    {
                        isDeclaration = true;
                        break;
                    }
                }
                
                // If we find a non-declaration line before reaching our target line,
                // then the target line is not in the declaration block
                if (!isDeclaration)
                {
                    return false;
                }
            }
            
            // Check if the current line itself is a declaration
            string currentLine = lines[lineIndex].Trim();
            foreach (var keyword in declarationKeywords)
            {
                if (currentLine.StartsWith(keyword, StringComparison.OrdinalIgnoreCase))
                {
                    return true;
                }
            }
            
            return false;
        }

        private string ExtractProcedureName(string line)
        {
            // Extract name from "(Public|Private) Sub Name(...)" or "(Public|Private) Function Name(...)"
            Regex regex = new Regex(@"(?:Public\s+|Private\s+|Friend\s+)?(?:Sub|Function)\s+([^\s(]+)", RegexOptions.IgnoreCase);
            Match match = regex.Match(line);
            if (match.Success && match.Groups.Count > 1)
            {
                return match.Groups[1].Value;
            }
            return "<unknown>";
        }
    }

    // Helper class to store procedure information
    class Procedure
    {
        public string Name { get; set; }
        public int StartLineIndex { get; set; } // 0-based
        public int EndLineIndex { get; set; }   // 0-based
    }
}
