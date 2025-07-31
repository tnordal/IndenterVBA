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
        private Office.CommandBarButton indentButton;
        private VbaIndenter vbaIndenter = new VbaIndenter();

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            AddVbeMenu();
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            RemoveVbeMenu();
        }

        private void AddVbeMenu()
        {
            try
            {
                var vbe = this.Application.VBE;
                Office.CommandBar toolsMenu = null;
                try
                {
                    toolsMenu = vbe.CommandBars["Tools"];
                }
                catch { }
                if (toolsMenu == null)
                {
                    System.Windows.Forms.MessageBox.Show("VBE Tools menu not found. Please open the VBA editor (Alt+F11) and restart the add-in.");
                    return;
                }

                // Avoid duplicate menu
                foreach (Office.CommandBarControl ctrl in toolsMenu.Controls)
                {
                    if (ctrl.Caption == "Indent Active Module")
                        return;
                }

                indentButton = (Office.CommandBarButton)toolsMenu.Controls.Add(
                    Office.MsoControlType.msoControlButton,
                    Temporary: true);
                indentButton.Caption = "Indent Active Module";
                indentButton.FaceId = 59; // Set a standard icon
                indentButton.Visible = true;
                indentButton.Click += new Office._CommandBarButtonEvents_ClickEventHandler(IndentButton_Click);
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show("Failed to add menu: " + ex.Message);
            }
        }

        private void RemoveVbeMenu()
        {
            try
            {
                if (indentButton != null)
                {
                    indentButton.Delete();
                    indentButton = null;
                }
            }
            catch { }
        }

        private void IndentButton_Click(Office.CommandBarButton Ctrl, ref bool CancelDefault)
        {
            IndentActiveModule();
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
