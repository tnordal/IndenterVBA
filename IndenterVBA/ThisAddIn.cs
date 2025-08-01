using System;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Excel;
using VBIDE = Microsoft.Vbe.Interop;

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
            System.Windows.Forms.MessageBox.Show("Indent All Modules button clicked");
        }

        private void IndentCurrentModuleButton_Click(Office.CommandBarButton Ctrl, ref bool CancelDefault)
        {
            System.Windows.Forms.MessageBox.Show("Indent Current Module button clicked");
        }

        private void IndentCurrentMethodButton_Click(Office.CommandBarButton Ctrl, ref bool CancelDefault)
        {
            System.Windows.Forms.MessageBox.Show("Indent Current Method button clicked");
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
                    System.Windows.Forms.MessageBox.Show("No active VBA project found.");
                    return;
                }
                
                vbaIndenter.IndentAllModules(app.VBE.ActiveVBProject);
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show("Error: " + ex.Message);
            }
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
