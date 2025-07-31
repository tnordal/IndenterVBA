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
                    if (ctrl.Caption == "Indent VBA Code")
                        return;
                }

                indentButton = (Office.CommandBarButton)toolsMenu.Controls.Add(
                    Office.MsoControlType.msoControlButton,
                    Temporary: true);
                indentButton.Caption = "Indent VBA Code";
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
            IndentAllVbaModules();
            System.Windows.Forms.MessageBox.Show("VBA code indented.");
        }

        private void IndentAllVbaModules()
        {
            var app = this.Application;
            VBIDE.VBProject vbProject = app.VBE.ActiveVBProject;
            vbaIndenter.IndentAllModules(vbProject);
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
