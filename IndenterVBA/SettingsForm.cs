using System;
using System.Windows.Forms;
using System.Reflection;

namespace IndenterVBA
{
    public partial class SettingsForm : Form
    {
        private IndenterSettings _settings;

        private void AddMenu()
        {
            // Create a menu strip
            MenuStrip menuStrip = new MenuStrip();

            // Create the 'About' menu item
            ToolStripMenuItem aboutMenuItem = new ToolStripMenuItem("About");
            aboutMenuItem.Click += AboutMenuItem_Click;

            // Add the 'About' menu item to the menu strip
            menuStrip.Items.Add(aboutMenuItem);

            // Add the menu strip to the form
            this.Controls.Add(menuStrip);
            this.MainMenuStrip = menuStrip;
            menuStrip.Dock = DockStyle.Top; // Ensure the menu strip is docked at the top

            // Force the menu strip to be visible
            menuStrip.Visible = true;
        }

        private void AboutMenuItem_Click(object sender, EventArgs e)
        {
            // Get the application name, version, and description
            string appName = Assembly.GetExecutingAssembly().GetName().Name;
            string appVersion = Assembly.GetExecutingAssembly().GetName().Version.ToString();
            string appDescription = ((AssemblyDescriptionAttribute)Attribute.GetCustomAttribute(
                Assembly.GetExecutingAssembly(), typeof(AssemblyDescriptionAttribute)))?.Description;

            // Show the message box
            MessageBox.Show($"{appName} - Version {appVersion}\n{appDescription}", "About", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        public SettingsForm()
        {
            InitializeComponent();
            _settings = IndenterSettings.Instance;
            AddMenu();
            LoadSettings();
        }

        private void LoadSettings()
        {
            numIndentSpaces.Value = _settings.IndentSpaces;
            chkIndentDeclarations.Checked = _settings.IndentDeclarations;
            chkUseLogging.Checked = _settings.UseLogging;
            chkIndentSelectCaseStatements.Checked = _settings.IndentSelectCaseStatements;
        }

        private void SaveSettings()
        {
            _settings.IndentSpaces = (int)numIndentSpaces.Value;
            _settings.IndentDeclarations = chkIndentDeclarations.Checked;
            _settings.UseLogging = chkUseLogging.Checked;
            _settings.IndentSelectCaseStatements = chkIndentSelectCaseStatements.Checked;
            _settings.Save();
        }

        private void btnSave_Click(object sender, EventArgs e)
        {
            SaveSettings();
            DialogResult = DialogResult.OK;
            Close();
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            DialogResult = DialogResult.Cancel;
            Close();
        }

        private void btnReset_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show(
                "Are you sure you want to reset all settings to their default values?",
                "Reset Settings",
                MessageBoxButtons.YesNo, 
                MessageBoxIcon.Question) == DialogResult.Yes)
            {
                _settings.ResetToDefault();
                LoadSettings();
            }
        }

        private void btnOpenLogsFolder_Click(object sender, EventArgs e)
        {
            try
            {
                System.Diagnostics.Process.Start(IndenterSettings.LogsFolder);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error opening logs folder: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void InitializeComponent()
        {
            this.lblIndentSpaces = new System.Windows.Forms.Label();
            this.numIndentSpaces = new System.Windows.Forms.NumericUpDown();
            this.chkIndentDeclarations = new System.Windows.Forms.CheckBox();
            this.chkUseLogging = new System.Windows.Forms.CheckBox();
            this.btnSave = new System.Windows.Forms.Button();
            this.btnCancel = new System.Windows.Forms.Button();
            this.btnReset = new System.Windows.Forms.Button();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.btnOpenLogsFolder = new System.Windows.Forms.Button();
            this.lblLogsPath = new System.Windows.Forms.Label();
            this.chkIndentSelectCaseStatements = new System.Windows.Forms.CheckBox();
            ((System.ComponentModel.ISupportInitialize)(this.numIndentSpaces)).BeginInit();
            this.groupBox1.SuspendLayout();
            this.groupBox2.SuspendLayout();
            this.SuspendLayout();
            // 
            // lblIndentSpaces
            // 
            this.lblIndentSpaces.AutoSize = true;
            this.lblIndentSpaces.Location = new System.Drawing.Point(15, 30);
            this.lblIndentSpaces.Name = "lblIndentSpaces";
            this.lblIndentSpaces.Size = new System.Drawing.Size(151, 13);
            this.lblIndentSpaces.TabIndex = 0;
            this.lblIndentSpaces.Text = "Number of spaces for indentation:";
            // 
            // numIndentSpaces
            // 
            this.numIndentSpaces.Location = new System.Drawing.Point(172, 28);
            this.numIndentSpaces.Maximum = new decimal(new int[] {
            8,
            0,
            0,
            0});
            this.numIndentSpaces.Minimum = new decimal(new int[] {
            1,
            0,
            0,
            0});
            this.numIndentSpaces.Name = "numIndentSpaces";
            this.numIndentSpaces.Size = new System.Drawing.Size(65, 20);
            this.numIndentSpaces.TabIndex = 1;
            this.numIndentSpaces.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            this.numIndentSpaces.Value = new decimal(new int[] {
            4,
            0,
            0,
            0});
            // 
            // chkIndentDeclarations
            // 
            this.chkIndentDeclarations.AutoSize = true;
            this.chkIndentDeclarations.Location = new System.Drawing.Point(18, 65);
            this.chkIndentDeclarations.Name = "chkIndentDeclarations";
            this.chkIndentDeclarations.Size = new System.Drawing.Size(205, 17);
            this.chkIndentDeclarations.TabIndex = 2;
            this.chkIndentDeclarations.Text = "Indent declarations (Dim, Private, etc.)";
            this.chkIndentDeclarations.UseVisualStyleBackColor = true;
            // 
            // chkUseLogging
            // 
            this.chkUseLogging.AutoSize = true;
            this.chkUseLogging.Location = new System.Drawing.Point(15, 30);
            this.chkUseLogging.Name = "chkUseLogging";
            this.chkUseLogging.Size = new System.Drawing.Size(145, 17);
            this.chkUseLogging.TabIndex = 3;
            this.chkUseLogging.Text = "Enable indentation logging";
            this.chkUseLogging.UseVisualStyleBackColor = true;
            // 
            // btnSave
            // 
            this.btnSave.Location = new System.Drawing.Point(120, 305);
            this.btnSave.Name = "btnSave";
            this.btnSave.Size = new System.Drawing.Size(75, 23);
            this.btnSave.TabIndex = 4;
            this.btnSave.Text = "Save";
            this.btnSave.UseVisualStyleBackColor = true;
            this.btnSave.Click += new System.EventHandler(this.btnSave_Click);
            // 
            // btnCancel
            // 
            this.btnCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.btnCancel.Location = new System.Drawing.Point(210, 305);
            this.btnCancel.Name = "btnCancel";
            this.btnCancel.Size = new System.Drawing.Size(75, 23);
            this.btnCancel.TabIndex = 5;
            this.btnCancel.Text = "Cancel";
            this.btnCancel.UseVisualStyleBackColor = true;
            this.btnCancel.Click += new System.EventHandler(this.btnCancel_Click);
            // 
            // btnReset
            // 
            this.btnReset.Location = new System.Drawing.Point(30, 305);
            this.btnReset.Name = "btnReset";
            this.btnReset.Size = new System.Drawing.Size(75, 23);
            this.btnReset.TabIndex = 6;
            this.btnReset.Text = "Reset";
            this.btnReset.UseVisualStyleBackColor = true;
            this.btnReset.Click += new System.EventHandler(this.btnReset_Click);
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.lblIndentSpaces);
            this.groupBox1.Controls.Add(this.numIndentSpaces);
            this.groupBox1.Controls.Add(this.chkIndentDeclarations);
            this.groupBox1.Controls.Add(this.chkIndentSelectCaseStatements);
            this.groupBox1.Location = new System.Drawing.Point(12, 50);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(285, 130);
            this.groupBox1.TabIndex = 7;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Indentation Settings";
            // 
            // groupBox2
            // 
            this.groupBox2.Controls.Add(this.lblLogsPath);
            this.groupBox2.Controls.Add(this.btnOpenLogsFolder);
            this.groupBox2.Controls.Add(this.chkUseLogging);
            this.groupBox2.Location = new System.Drawing.Point(12, 190);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(285, 110);
            this.groupBox2.TabIndex = 8;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "Logging Settings";
            // 
            // btnOpenLogsFolder
            // 
            this.btnOpenLogsFolder.Location = new System.Drawing.Point(15, 60);
            this.btnOpenLogsFolder.Name = "btnOpenLogsFolder";
            this.btnOpenLogsFolder.Size = new System.Drawing.Size(127, 23);
            this.btnOpenLogsFolder.TabIndex = 4;
            this.btnOpenLogsFolder.Text = "Open Logs Folder";
            this.btnOpenLogsFolder.UseVisualStyleBackColor = true;
            this.btnOpenLogsFolder.Click += new System.EventHandler(this.btnOpenLogsFolder_Click);
            // 
            // lblLogsPath
            // 
            this.lblLogsPath.AutoSize = true;
            this.lblLogsPath.Location = new System.Drawing.Point(15, 90);
            this.lblLogsPath.Name = "lblLogsPath";
            this.lblLogsPath.Size = new System.Drawing.Size(113, 13);
            this.lblLogsPath.TabIndex = 5;
            this.lblLogsPath.Text = "Logs saved in MyDocs\\IndentVBA\\logs";
            // 
            // chkIndentSelectCaseStatements
            // 
            this.chkIndentSelectCaseStatements.AutoSize = true;
            this.chkIndentSelectCaseStatements.Location = new System.Drawing.Point(18, 90);
            this.chkIndentSelectCaseStatements.Name = "chkIndentSelectCaseStatements";
            this.chkIndentSelectCaseStatements.Size = new System.Drawing.Size(250, 17);
            this.chkIndentSelectCaseStatements.TabIndex = 3;
            this.chkIndentSelectCaseStatements.Text = "Indent Select Case statements (aligned or nested)";
            this.chkIndentSelectCaseStatements.UseVisualStyleBackColor = true;
            // 
            // SettingsForm
            // 
            this.AcceptButton = this.btnSave;
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.CancelButton = this.btnCancel;
            this.ClientSize = new System.Drawing.Size(309, 340);
            this.Controls.Add(this.groupBox2);
            this.Controls.Add(this.groupBox1);
            this.Controls.Add(this.btnReset);
            this.Controls.Add(this.btnCancel);
            this.Controls.Add(this.btnSave);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "SettingsForm";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "VBA Indenter Settings";
            ((System.ComponentModel.ISupportInitialize)(this.numIndentSpaces)).EndInit();
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.groupBox2.ResumeLayout(false);
            this.groupBox2.PerformLayout();
            this.ResumeLayout(false);

        }

        private System.Windows.Forms.Label lblIndentSpaces;
        private System.Windows.Forms.NumericUpDown numIndentSpaces;
        private System.Windows.Forms.CheckBox chkIndentDeclarations;
        private System.Windows.Forms.CheckBox chkUseLogging;
        private System.Windows.Forms.Button btnSave;
        private System.Windows.Forms.Button btnCancel;
        private System.Windows.Forms.Button btnReset;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.GroupBox groupBox2;
        private System.Windows.Forms.Button btnOpenLogsFolder;
        private System.Windows.Forms.Label lblLogsPath;
        private System.Windows.Forms.CheckBox chkIndentSelectCaseStatements;
    }
}