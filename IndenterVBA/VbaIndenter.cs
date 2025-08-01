using System;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using System.Collections.Generic;
using System.Linq;
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
        
        // Indent only the current procedure (method) at cursor position
        public bool IndentCurrentProcedure(VBIDE.VBE vbe)
        {
            try
            {
                // Check if we have an active code pane
                if (vbe.ActiveCodePane == null)
                {
                    MessageBox.Show("Please open a code module first.");
                    return false;
                }

                // Get the active code module
                VBIDE.CodeModule codeModule = vbe.ActiveCodePane.CodeModule;
                if (codeModule == null)
                {
                    MessageBox.Show("No active code module found.");
                    return false;
                }

                // Get current cursor position
                int currentLine, currentColumn, endLine, endColumn;
                vbe.ActiveCodePane.GetSelection(out currentLine, out currentColumn, out endLine, out endColumn);

                // Initialize log
                StringBuilder log = new StringBuilder();
                log.AppendLine($"Module: {codeModule.Name}");
                log.AppendLine($"Current Position: Line {currentLine}, Column {currentColumn}");

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

                // Split code into lines
                string[] lines = code.Split(new[] { "\r\n", "\n" }, StringSplitOptions.None);
                
                // Find all procedures in the code
                List<ProcedureInfo> procedures = FindProcedures(lines, log);
                
                // Find the procedure containing the current cursor position
                ProcedureInfo currentProcedure = null;
                foreach (var proc in procedures)
                {
                    int procStart = proc.StartLineIndex + 1; // 1-based
                    int procEnd = proc.EndLineIndex + 1; // 1-based
                    
                    if (currentLine >= procStart && currentLine <= procEnd)
                    {
                        currentProcedure = proc;
                        break;
                    }
                }
                
                if (currentProcedure == null)
                {
                    MessageBox.Show("Cursor is not within any procedure. Please position the cursor inside a Sub, Function, or Property.");
                    return false;
                }
                
                log.AppendLine($"\nIndenting procedure: {currentProcedure.Type} {currentProcedure.Name}");
                
                // Get a copy of the original lines for comparison
                string[] originalLines = new string[lines.Length];
                Array.Copy(lines, originalLines, lines.Length);
                
                // Only indent the current procedure
                SimpleMultiPassIndentation(lines, currentProcedure, log);
                
                // Create a new string with the updated code
                StringBuilder updatedCode = new StringBuilder();
                for (int i = 0; i < lines.Length; i++)
                {
                    updatedCode.AppendLine(lines[i]);
                }
                
                // Write the indented code to a log file for verification
                File.WriteAllText("C:\\Temp\\Indented_Current_Procedure.txt", updatedCode.ToString());
                
                // Apply the changes only to the procedure
                int procStartLine = currentProcedure.StartLineIndex + 1; // 1-based line number
                int procEndLine = currentProcedure.EndLineIndex + 1; // 1-based line number
                int procBodyLineCount = procEndLine - procStartLine + 1;
                
                // Get the procedure body from our indented code
                StringBuilder procBody = new StringBuilder();
                for (int i = currentProcedure.StartLineIndex; i <= currentProcedure.EndLineIndex; i++)
                {
                    procBody.AppendLine(lines[i]);
                }
                
                // Replace the procedure in the module
                codeModule.DeleteLines(procStartLine, procBodyLineCount);
                codeModule.InsertLines(procStartLine, procBody.ToString());
                
                // Write log
                File.WriteAllText("C:\\Temp\\VbaIndentLog_CurrentProc.txt", log.ToString());
                
                // Position cursor back to the original procedure but at the start
                vbe.ActiveCodePane.SetSelection(procStartLine, 1, procStartLine, 1);
                
                return true;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error: {ex.Message}");
                File.WriteAllText("C:\\Temp\\IndentError.txt", ex.ToString());
                return false;
            }
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

                // Process and indent the code
                log.AppendLine("\nIndenting procedures:");
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

                MessageBox.Show("Indentation complete. Check C:\\Temp\\VbaIndentLog.txt for details and C:\\Temp\\Pass_*.txt for debugging information.");
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
            
            // Find all procedures in the code
            List<ProcedureInfo> procedures = FindProcedures(lines, log);
            
            // Process each procedure with the simplified level-based indentation
            foreach (var procedure in procedures)
            {
                SimpleMultiPassIndentation(lines, procedure, log);
            }
            
            // Combine the indented lines back into a single string
            StringBuilder indentedCode = new StringBuilder();
            foreach (string line in lines)
            {
                indentedCode.AppendLine(line);
            }
            
            return indentedCode.ToString();
        }
        
        private void SimpleMultiPassIndentation(string[] lines, ProcedureInfo procedure, StringBuilder log)
        {
            log.AppendLine($"Processing {procedure.Type} {procedure.Name}");
            
            // Directory for debug files
            string debugDir = "C:\\Temp";
            Directory.CreateDirectory(debugDir);
            
            // Pass 1: Reset all lines in the procedure to have no indentation
            StringBuilder resetText = new StringBuilder();
            for (int i = procedure.StartLineIndex; i <= procedure.EndLineIndex; i++)
            {
                lines[i] = lines[i].TrimStart();
                resetText.AppendLine(lines[i]);
            }
            File.WriteAllText(Path.Combine(debugDir, $"Pass_1_Reset_{procedure.Name}.txt"), resetText.ToString());
            
            // Pass 2: Identify blocks and their structure
            List<Tuple<int, int>> blockRanges = IdentifyBlockRanges(lines, procedure);
            
            // Create debug output for block structure
            StringBuilder blockDebug = new StringBuilder();
            blockDebug.AppendLine($"Identified {blockRanges.Count} code blocks in {procedure.Name}:");
            foreach (var range in blockRanges)
            {
                string startLine = lines[range.Item1].Trim();
                string endLine = lines[range.Item2].Trim();
                blockDebug.AppendLine($"Block from line {range.Item1 + 1} to {range.Item2 + 1}");
                blockDebug.AppendLine($"  Start: {startLine}");
                blockDebug.AppendLine($"  End: {endLine}");
            }
            File.WriteAllText(Path.Combine(debugDir, $"Pass_2_Blocks_{procedure.Name}.txt"), blockDebug.ToString());
            
            // Pass 3: Indent procedure-level statements (level 0)
            ApplyInitialIndentation(lines, procedure);
            StringBuilder level0Text = new StringBuilder();
            for (int i = procedure.StartLineIndex; i <= procedure.EndLineIndex; i++)
            {
                level0Text.AppendLine(lines[i]);
            }
            File.WriteAllText(Path.Combine(debugDir, $"Pass_3_Level0_{procedure.Name}.txt"), level0Text.ToString());
            
            // Create a mapping of each line to its indentation level
            Dictionary<int, int> lineIndentLevels = new Dictionary<int, int>();
            
            // Pass 4: Process each block level by level
            int maxLevel = blockRanges.Count > 0 ? blockRanges.Count : 0;
            
            // First, organize blocks by their nesting level
            Dictionary<int, List<Tuple<int, int>>> blocksByLevel = OrganizeBlocksByLevel(blockRanges, lines);
            
            // Process blocks level by level
            for (int level = 1; level <= maxLevel; level++)
            {
                if (blocksByLevel.ContainsKey(level))
                {
                    // Apply indentation for this level's blocks
                    ApplyLevelIndentation(lines, blocksByLevel[level], level);
                    
                    // Output the state after this level
                    StringBuilder levelText = new StringBuilder();
                    for (int i = procedure.StartLineIndex; i <= procedure.EndLineIndex; i++)
                    {
                        levelText.AppendLine(lines[i]);
                    }
                    File.WriteAllText(Path.Combine(debugDir, $"Pass_4_Level{level}_{procedure.Name}.txt"), levelText.ToString());
                }
            }
            
            // Pass 5: Special cases (error handling)
            ProcessErrorHandling(lines, procedure);
            
            // Final output
            StringBuilder finalText = new StringBuilder();
            for (int i = procedure.StartLineIndex; i <= procedure.EndLineIndex; i++)
            {
                finalText.AppendLine(lines[i]);
            }
            File.WriteAllText(Path.Combine(debugDir, $"Pass_5_Final_{procedure.Name}.txt"), finalText.ToString());
        }

        private List<Tuple<int, int>> IdentifyBlockRanges(string[] lines, ProcedureInfo procedure)
        {
            List<Tuple<int, int>> blockRanges = new List<Tuple<int, int>>();
            Stack<int> blockStarts = new Stack<int>();
            Dictionary<string, string> blockEndKeywords = new Dictionary<string, string>
            {
                { "for", "next" },
                { "for each", "next" },
                { "do", "loop" },
                { "while", "wend" },
                { "if", "end if" },
                { "select case", "end select" },
                { "with", "end with" }
            };
            
            for (int i = procedure.StartLineIndex + 1; i < procedure.EndLineIndex; i++)
            {
                string line = lines[i].Trim().ToLower();
                if (string.IsNullOrWhiteSpace(line)) continue;
                
                // Check for block start
                bool isStart = false;
                string startKeyword = "";
                
                foreach (var keyword in blockEndKeywords.Keys)
                {
                    if (line.StartsWith(keyword + " ") || line == keyword)
                    {
                        isStart = true;
                        startKeyword = keyword;
                        break;
                    }
                }
                
                // Handle If statements separately
                if (line.StartsWith("if ") && line.EndsWith(" then") && !line.Contains("then:"))
                {
                    isStart = true;
                    startKeyword = "if";
                }
                
                if (isStart)
                {
                    blockStarts.Push(i);
                }
                
                // Check for block end
                bool isEnd = false;
                string matchingStart = "";
                
                foreach (var pair in blockEndKeywords)
                {
                    if (line.StartsWith(pair.Value))
                    {
                        isEnd = true;
                        matchingStart = pair.Key;
                        break;
                    }
                }
                
                if (isEnd && blockStarts.Count > 0)
                {
                    int startLine = blockStarts.Pop();
                    string startLineText = lines[startLine].Trim().ToLower();
                    
                    // Verify this is a matching pair
                    bool isMatchingPair = false;
                    foreach (var pair in blockEndKeywords)
                    {
                        if ((startLineText.StartsWith(pair.Key + " ") || startLineText == pair.Key) && 
                            line.StartsWith(pair.Value))
                        {
                            isMatchingPair = true;
                            break;
                        }
                    }
                    
                    // If this is a matching pair, add to block ranges
                    if (isMatchingPair)
                    {
                        blockRanges.Add(new Tuple<int, int>(startLine, i));
                    }
                    else
                    {
                        // Push the start back, as this wasn't a match
                        blockStarts.Push(startLine);
                    }
                }
            }
            
            return blockRanges;
        }
        
        private Dictionary<int, List<Tuple<int, int>>> OrganizeBlocksByLevel(List<Tuple<int, int>> blockRanges, string[] lines)
        {
            var result = new Dictionary<int, List<Tuple<int, int>>>();
            
            // First, sort all blocks by their start line
            blockRanges = blockRanges.OrderBy(r => r.Item1).ToList();
            
            // Determine the nesting level for each block
            foreach (var block in blockRanges)
            {
                int level = 1; // Default level is 1
                
                // Check if this block is nested inside any other blocks
                foreach (var outerBlock in blockRanges)
                {
                    // If the current block is completely contained in another block,
                    // and it's not the same block, increase its level
                    if (block.Item1 > outerBlock.Item1 && block.Item2 < outerBlock.Item2)
                    {
                        level++;
                    }
                }
                
                // Add this block to the appropriate level list
                if (!result.ContainsKey(level))
                {
                    result[level] = new List<Tuple<int, int>>();
                }
                result[level].Add(block);
            }
            
            return result;
        }
        
        private void ApplyInitialIndentation(string[] lines, ProcedureInfo procedure)
        {
            // Declarations are at the start and get no indentation
            bool inDeclarationSection = true;
            
            for (int i = procedure.StartLineIndex + 1; i < procedure.EndLineIndex; i++)
            {
                string line = lines[i].Trim();
                if (string.IsNullOrWhiteSpace(line)) continue;
                
                // Check if this is a declaration
                if (inDeclarationSection && IsDeclarationLine(line))
                {
                    // No indentation for declarations
                    lines[i] = line;
                }
                else
                {
                    // Once we hit a non-declaration, all following lines are indented
                    inDeclarationSection = false;
                    
                    // Add basic indentation (4 spaces) to the procedure body
                    lines[i] = "    " + line;
                }
            }
        }

        private void ApplyLevelIndentation(string[] lines, List<Tuple<int, int>> blocks, int level)
        {
            foreach (var block in blocks)
            {
                int startLine = block.Item1;
                int endLine = block.Item2;
                string startLineText = lines[startLine].Trim();
                string endLineText = lines[endLine].Trim();
                
                // Calculate indentation for this level
                string indent = new string(' ', level * 4);
                
                // Apply indentation to block start and end lines
                lines[startLine] = indent + startLineText;
                lines[endLine] = indent + endLineText;
                
                // Apply indentation to lines inside the block
                ApplyContentIndentation(lines, startLine, endLine, level);
            }
        }
        
        private void ApplyContentIndentation(string[] lines, int startLine, int endLine, int level)
        {
            // Get the block type from the start line
            string blockStart = lines[startLine].Trim().ToLower();
            bool isForEachBlock = blockStart.StartsWith("for each ");
            
            // Apply indentation to content lines
            for (int i = startLine + 1; i < endLine; i++)
            {
                string line = lines[i].Trim();
                if (string.IsNullOrWhiteSpace(line)) continue;
                
                // Special case handling for block midpoints (else, elseif, case)
                if (IsBlockMidpoint(line))
                {
                    // Block midpoints get the same indentation as the block start
                    lines[i] = new string(' ', level * 4) + line;
                }
                // Lines inside the block that aren't already the start/end of another block
                // and aren't block midpoints get one more level of indentation
                else if (!IsBlockBoundary(line))
                {
                    // Contents get indented one more level
                    lines[i] = new string(' ', (level + 1) * 4) + line;
                }
            }
        }
        
        private bool IsBlockMidpoint(string line)
        {
            line = line.ToLower();
            return line == "else" || 
                   line.StartsWith("elseif ") || 
                   line.StartsWith("case ") || 
                   line == "case else";
        }
        
        private bool IsBlockBoundary(string line)
        {
            line = line.ToLower();
            
            // Check if this is a block start
            if (line.StartsWith("for ") || 
                line.StartsWith("for each ") || 
                line.StartsWith("do ") || 
                line == "do" ||
                (line.StartsWith("if ") && line.EndsWith(" then") && !line.Contains("then:")) ||
                line.StartsWith("while ") ||
                line.StartsWith("select case") ||
                line.StartsWith("with "))
            {
                return true;
            }
            
            // Check if this is a block end
            if (line.StartsWith("next") ||
                line.StartsWith("loop") ||
                line.StartsWith("end if") ||
                line.StartsWith("wend") ||
                line.StartsWith("end select") ||
                line.StartsWith("end with"))
            {
                return true;
            }
            
            return false;
        }
        
        private void ProcessErrorHandling(string[] lines, ProcedureInfo procedure)
        {
            for (int i = procedure.StartLineIndex + 1; i < procedure.EndLineIndex; i++)
            {
                string line = lines[i].Trim().ToLower();
                if (string.IsNullOrWhiteSpace(line)) continue;
                
                // Handle error handling statements
                if (line.StartsWith("on error "))
                {
                    // Get the current indentation level (preserve it)
                    int currentIndent = GetIndentationLevel(lines[i]);
                    string indentation = new string(' ', currentIndent);
                    
                    // Apply indentation to the On Error statement
                    lines[i] = indentation + line;
                    
                    // Check for On Error Resume Next pattern
                    if (line.Contains("resume next"))
                    {
                        // Handle the If Err.Number <> 0 pattern
                        for (int j = i + 1; j < procedure.EndLineIndex; j++)
                        {
                            string checkLine = lines[j].Trim().ToLower();
                            if (string.IsNullOrWhiteSpace(checkLine)) continue;
                            
                            // Found error check statement
                            if (checkLine.Contains("if err.") && checkLine.Contains("<> 0"))
                            {
                                // Apply same indentation as On Error
                                lines[j] = indentation + checkLine;
                                
                                // Find the matching End If
                                for (int k = j + 1; k < procedure.EndLineIndex; k++)
                                {
                                    string endIfCheck = lines[k].Trim().ToLower();
                                    
                                    if (endIfCheck.StartsWith("end if"))
                                    {
                                        // End If gets same indentation as If
                                        lines[k] = indentation + endIfCheck;
                                        
                                        // Lines between If and End If get one more level
                                        for (int m = j + 1; m < k; m++)
                                        {
                                            string content = lines[m].Trim();
                                            if (!string.IsNullOrWhiteSpace(content))
                                            {
                                                lines[m] = indentation + "    " + content;
                                            }
                                        }
                                        
                                        // Check for On Error GoTo 0 after End If
                                        if (k + 1 < procedure.EndLineIndex)
                                        {
                                            string nextLine = lines[k + 1].Trim().ToLower();
                                            if (nextLine.StartsWith("on error goto 0"))
                                            {
                                                lines[k + 1] = indentation + nextLine;
                                            }
                                        }
                                        break;
                                    }
                                }
                                break;
                            }
                        }
                    }
                }
            }
        }
        
        private int GetIndentationLevel(string line)
        {
            int spaces = 0;
            foreach (char c in line)
            {
                if (c == ' ')
                    spaces++;
                else
                    break;
            }
            return spaces;
        }
        
        private bool IsDeclarationLine(string line)
        {
            line = line.ToLower();
            return line.StartsWith("dim ") || 
                   line.StartsWith("private ") || 
                   line.StartsWith("public ") || 
                   line.StartsWith("static ") || 
                   line.StartsWith("const ") || 
                   line.StartsWith("redim ");
        }

        private List<ProcedureInfo> FindProcedures(string[] lines, StringBuilder log)
        {
            List<ProcedureInfo> procedures = new List<ProcedureInfo>();
            
            for (int i = 0; i < lines.Length; i++)
            {
                string line = lines[i].Trim();
                
                // Check for procedure start
                if (IsProcedureStart(line, out string name, out string type))
                {
                    int startLine = i;
                    int endLine = -1;
                    
                    // Look for matching End Sub/Function
                    for (int j = i + 1; j < lines.Length; j++)
                    {
                        string endLine1 = lines[j].Trim();
                        if (IsProcedureEnd(endLine1, type))
                        {
                            endLine = j;
                            break;
                        }
                    }
                    
                    if (endLine >= 0)
                    {
                        // Add procedure to list
                        procedures.Add(new ProcedureInfo
                        {
                            Name = name,
                            Type = type,
                            StartLineIndex = startLine,
                            EndLineIndex = endLine
                        });
                        
                        log.AppendLine($"Found {type} '{name}' from line {startLine + 1} to {endLine + 1}");
                        
                        // Skip to end of this procedure
                        i = endLine;
                    }
                }
            }
            
            return procedures;
        }
        
        private bool IsProcedureStart(string line, out string name, out string type)
        {
            name = "";
            type = "";
            
            // Check for Sub or Function declaration with optional access modifiers
            Regex regex = new Regex(@"^\s*(Public\s+|Private\s+|Friend\s+)?(Sub|Function|Property\s+[^\s]+)\s+([^\s(]+)", RegexOptions.IgnoreCase);
            Match match = regex.Match(line);
            
            if (match.Success && !line.StartsWith("End ", StringComparison.OrdinalIgnoreCase))
            {
                type = match.Groups[2].Value;
                name = match.Groups[3].Value;
                return true;
            }
            
            return false;
        }
        
        private bool IsProcedureEnd(string line, string type)
        {
            // Handle Property Let/Get/Set separately
            if (type.StartsWith("Property", StringComparison.OrdinalIgnoreCase))
            {
                return line.StartsWith("End Property", StringComparison.OrdinalIgnoreCase);
            }
            
            return line.Equals($"End {type}", StringComparison.OrdinalIgnoreCase);
        }
        
        #region Helper Classes
        
        // Stores information about a procedure
        private class ProcedureInfo
        {
            public string Name { get; set; }
            public string Type { get; set; } // "Sub" or "Function"
            public int StartLineIndex { get; set; } // 0-based
            public int EndLineIndex { get; set; }   // 0-based
        }
        
        #endregion
    }
}
