using System;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Excel;
using VBIDE = Microsoft.Vbe.Interop;
using System.Windows.Forms;

namespace IndenterVBA
{
    public partial class ThisAddIn
    {
        private Office.CommandBarPopup vbaIndenterMenu;
        private Office.CommandBarButton indentAllButton;
        private Office.CommandBarButton indentCurrentModuleButton;
        private Office.CommandBarButton indentCurrentMethodButton;
        private Office.CommandBarButton settingsButton;
        private VbaIndenter vbaIndenter = new VbaIndenter();

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            AddVbaIndenterMenu();
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            RemoveVbaIndenterMenu();
        }

        private void AddVbaIndenterMenu()
        {
            try
            {
                var vbe = this.Application.VBE;
                Office.CommandBar menuBar = null;
                
                try
                {
                    menuBar = vbe.CommandBars["Menu Bar"];
                }
                catch { }
                
                if (menuBar == null)
                {
                    System.Windows.Forms.MessageBox.Show("VBE Menu Bar not found. Please open the VBA editor (Alt+F11) and restart the add-in.");
                    return;
                }

                // Check for existing menu to avoid duplicates
                foreach (Office.CommandBarControl ctrl in menuBar.Controls)
                {
                    if (ctrl.Caption == "VBA Indenter")
                        return;
                }

                // Find the Tools menu to position our menu after it
                int toolsPosition = -1;
                for (int i = 1; i <= menuBar.Controls.Count; i++)
                {
                    if (menuBar.Controls[i].Caption == "Tools")
                    {
                        toolsPosition = i;
                        break;
                    }
                }

                // Add our main menu after the Tools menu (or at the end if Tools menu not found)
                vbaIndenterMenu = (Office.CommandBarPopup)menuBar.Controls.Add(
                    Office.MsoControlType.msoControlPopup, 
                    Before: toolsPosition > 0 ? toolsPosition + 1 : menuBar.Controls.Count + 1,
                    Temporary: true);
                vbaIndenterMenu.Caption = "VBA Indenter";
                vbaIndenterMenu.Visible = true;

                // Add the four buttons
                indentAllButton = (Office.CommandBarButton)vbaIndenterMenu.Controls.Add(
                    Office.MsoControlType.msoControlButton,
                    Temporary: true);
                indentAllButton.Caption = "Indent All Modules";
                indentAllButton.FaceId = 59;
                indentAllButton.Visible = true;
                indentAllButton.Click += new Office._CommandBarButtonEvents_ClickEventHandler(IndentAllButton_Click);

                indentCurrentModuleButton = (Office.CommandBarButton)vbaIndenterMenu.Controls.Add(
                    Office.MsoControlType.msoControlButton,
                    Temporary: true);
                indentCurrentModuleButton.Caption = "Indent Current Module";
                indentCurrentModuleButton.FaceId = 59;
                indentCurrentModuleButton.Visible = true;
                indentCurrentModuleButton.Click += new Office._CommandBarButtonEvents_ClickEventHandler(IndentCurrentModuleButton_Click);

                indentCurrentMethodButton = (Office.CommandBarButton)vbaIndenterMenu.Controls.Add(
                    Office.MsoControlType.msoControlButton,
                    Temporary: true);
                indentCurrentMethodButton.Caption = "Indent Current Method";
                indentCurrentMethodButton.FaceId = 59;
                indentCurrentMethodButton.Visible = true;
                indentCurrentMethodButton.Click += new Office._CommandBarButtonEvents_ClickEventHandler(IndentCurrentMethodButton_Click);

                settingsButton = (Office.CommandBarButton)vbaIndenterMenu.Controls.Add(
                    Office.MsoControlType.msoControlButton,
                    Temporary: true);
                settingsButton.Caption = "Settings";
                settingsButton.FaceId = 59;
                settingsButton.Visible = true;
                settingsButton.Click += new Office._CommandBarButtonEvents_ClickEventHandler(SettingsButton_Click);
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show("Failed to add VBA Indenter menu: " + ex.Message);
            }
        }

        private void RemoveVbaIndenterMenu()
        {
            try
            {
                if (vbaIndenterMenu != null)
                {
                    vbaIndenterMenu.Delete();
                    vbaIndenterMenu = null;
                    indentAllButton = null;
                    indentCurrentModuleButton = null;
                    indentCurrentMethodButton = null;
                    settingsButton = null;
                }
            }
            catch { }
        }

        private void IndentAllButton_Click(Office.CommandBarButton Ctrl, ref bool CancelDefault)
        {
            try
            {
                var app = this.Application;
                if (app.VBE.ActiveVBProject == null)
                {
                    MessageBox.Show("No active VBA project found.");
                    return;
                }

                var vbProject = app.VBE.ActiveVBProject;
                int moduleCount = 0;
                
                // Count total modules for progress information
                foreach (VBIDE.VBComponent component in vbProject.VBComponents)
                {
                    if (IsCodeModule(component.Type))
                    {
                        moduleCount++;
                    }
                }

                if (moduleCount == 0)
                {
                    MessageBox.Show("No code modules found in the active project.");
                    return;
                }

                if (MessageBox.Show($"This will indent all {moduleCount} modules in the active project. Continue?", 
                                    "Confirm Indentation", MessageBoxButtons.YesNo, 
                                    MessageBoxIcon.Question) != DialogResult.Yes)
                {
                    return;
                }

                int processedCount = 0;
                
                // Process each module in the project
                foreach (VBIDE.VBComponent component in vbProject.VBComponents)
                {
                    if (IsCodeModule(component.Type))
                    {
                        try
                        {
                            // Activate this code module
                            component.CodeModule.CodePane.SetSelection(1, 1, 1, 1);
                            component.CodeModule.CodePane.Show();
                            
                            // Indent the module
                            vbaIndenter.IndentAllModules(vbProject);
                            processedCount++;
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show($"Error processing module {component.Name}: {ex.Message}");
                        }
                    }
                }

                MessageBox.Show($"Indentation complete. Processed {processedCount} of {moduleCount} modules.");
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error: " + ex.Message);
            }
        }

        private void IndentCurrentModuleButton_Click(Office.CommandBarButton Ctrl, ref bool CancelDefault)
        {
            try
            {
                IndentActiveModule();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error: " + ex.Message);
            }
        }

        private void IndentCurrentMethodButton_Click(Office.CommandBarButton Ctrl, ref bool CancelDefault)
        {
            try
            {
                IndentActiveMethod();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error: " + ex.Message);
            }
        }

        private void SettingsButton_Click(Office.CommandBarButton Ctrl, ref bool CancelDefault)
        {
            System.Windows.Forms.MessageBox.Show("Settings button clicked");
        }

        private void IndentActiveModule()
        {
            try
            {
                var app = this.Application;
                if (app.VBE.ActiveVBProject == null)
                {
                    MessageBox.Show("No active VBA project found.");
                    return;
                }
                
                if (app.VBE.ActiveCodePane == null)
                {
                    MessageBox.Show("Please open a code module first.");
                    return;
                }

                vbaIndenter.IndentAllModules(app.VBE.ActiveVBProject);
                MessageBox.Show("Current module indentation complete.");
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error: " + ex.Message);
            }
        }
        
        private void IndentActiveMethod()
        {
            try
            {
                var app = this.Application;
                if (app.VBE.ActiveVBProject == null)
                {
                    MessageBox.Show("No active VBA project found.");
                    return;
                }
                
                if (app.VBE.ActiveCodePane == null)
                {
                    MessageBox.Show("Please open a code module first.");
                    return;
                }

                // Get current selection
                int startLine, startColumn, endLine, endColumn;
                app.VBE.ActiveCodePane.GetSelection(out startLine, out startColumn, out endLine, out endColumn);
                
                // Call the new method that indents only the current method
                bool success = vbaIndenter.IndentCurrentProcedure(app.VBE);
                
                if (success)
                {
                    MessageBox.Show("Current method indentation complete.");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error: " + ex.Message);
            }
        }

        private bool IsCodeModule(VBIDE.vbext_ComponentType type)
        {
            // These are the component types that can contain code
            return type == VBIDE.vbext_ComponentType.vbext_ct_StdModule ||
                   type == VBIDE.vbext_ComponentType.vbext_ct_ClassModule ||
                   type == VBIDE.vbext_ComponentType.vbext_ct_MSForm ||
                   type == VBIDE.vbext_ComponentType.vbext_ct_Document;
        }

        #region VSTO generated code

        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }

        #endregion
    }
}
