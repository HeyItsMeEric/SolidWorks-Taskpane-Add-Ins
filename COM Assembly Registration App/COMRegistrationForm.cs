using System;
using System.IO;
using System.Windows.Forms;

namespace COM_Assembly_Registration_App {
    internal static class Program {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        internal static void Main() {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new COMRegistrationForm());
        }
    }
    
    /// <inheritdoc />
    internal class COMRegistrationForm : Form {
        
        //Form Members
        private System.Windows.Forms.Label headerLabel;
        private System.Windows.Forms.Button installButton;
        private System.Windows.Forms.Button browseForCOMButton;
        private System.Windows.Forms.Button uninstallButton;
        private System.Windows.Forms.Button browseForRegAsm;
        private System.Windows.Forms.TextBox comTextBox;
        private System.Windows.Forms.TextBox regAsmTextBox;
        private System.Windows.Forms.MenuStrip menuStrip;
        private System.Windows.Forms.ToolStripMenuItem aboutToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem howThisWorksToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem aboutTheAuthorToolStripMenuItem;
        private System.Windows.Forms.Label autoDetectionLabel;
        private System.Windows.Forms.Label pathToRegAsmLabel;
        private System.Windows.Forms.Label pathToCOMLabel;
        private System.Windows.Forms.Label descriptionLabel;
        
        /// <summary>
        /// 
        /// </summary>
        public COMRegistrationForm() {
            InitializeComponent();
            if (!File.Exists(@"C:\Windows\Microsoft.NET\Framework64\v4.0.30319\RegAsm.exe")) return;
            regAsmTextBox.Text = @"C:\Windows\Microsoft.NET\Framework64\v4.0.30319\RegAsm.exe";
            autoDetectionLabel.Visible = true;
            
            //Centering the Form in the middle of the screen
            this.Location = new System.Drawing.Point((Screen.FromControl(this).Bounds.Width - this.Width) / 2,
                                                     (Screen.FromControl(this).Bounds.Height / 7));
        }

        /// <summary>
        /// Required designer variable.
        /// </summary>
        private readonly System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing) {
            if (disposing && (components != null)) {
                components.Dispose();
            }

            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent() {
            this.headerLabel = new System.Windows.Forms.Label();
            this.installButton = new System.Windows.Forms.Button();
            this.browseForCOMButton = new System.Windows.Forms.Button();
            this.uninstallButton = new System.Windows.Forms.Button();
            this.browseForRegAsm = new System.Windows.Forms.Button();
            this.comTextBox = new System.Windows.Forms.TextBox();
            this.regAsmTextBox = new System.Windows.Forms.TextBox();
            this.menuStrip = new System.Windows.Forms.MenuStrip();
            this.aboutToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.howThisWorksToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.aboutTheAuthorToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.descriptionLabel = new System.Windows.Forms.Label();
            this.pathToCOMLabel = new System.Windows.Forms.Label();
            this.pathToRegAsmLabel = new System.Windows.Forms.Label();
            this.autoDetectionLabel = new System.Windows.Forms.Label();
            this.menuStrip.SuspendLayout();
            this.SuspendLayout();
            // 
            // headerLabel
            // 
            this.headerLabel.AutoSize = true;
            this.headerLabel.Font = new System.Drawing.Font("Century Gothic", 24F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte) (0)));
            this.headerLabel.Location = new System.Drawing.Point(130, 41);
            this.headerLabel.Name = "headerLabel";
            this.headerLabel.Size = new System.Drawing.Size(446, 38);
            this.headerLabel.TabIndex = 0;
            this.headerLabel.Text = "COM Assembly Registration";
            // 
            // installButton
            // 
            this.installButton.Location = new System.Drawing.Point(19, 251);
            this.installButton.Name = "installButton";
            this.installButton.Size = new System.Drawing.Size(522, 23);
            this.installButton.TabIndex = 1;
            this.installButton.Text = "Install";
            this.installButton.UseVisualStyleBackColor = true;
            this.installButton.Click += new System.EventHandler(this.installButton_Click);
            // 
            // browseForCOMButton
            // 
            this.browseForCOMButton.Location = new System.Drawing.Point(582, 134);
            this.browseForCOMButton.Name = "browseForCOMButton";
            this.browseForCOMButton.Size = new System.Drawing.Size(75, 23);
            this.browseForCOMButton.TabIndex = 2;
            this.browseForCOMButton.Text = "Browse...";
            this.browseForCOMButton.UseVisualStyleBackColor = true;
            this.browseForCOMButton.Click += new System.EventHandler(this.BrowseForComButtonClicked);
            // 
            // uninstallButton
            // 
            this.uninstallButton.ForeColor = System.Drawing.Color.Red;
            this.uninstallButton.Location = new System.Drawing.Point(547, 251);
            this.uninstallButton.Name = "uninstallButton";
            this.uninstallButton.Size = new System.Drawing.Size(110, 23);
            this.uninstallButton.TabIndex = 3;
            this.uninstallButton.Text = "Uninstall";
            this.uninstallButton.UseVisualStyleBackColor = true;
            this.uninstallButton.Click += new System.EventHandler(this.UninstallButtonClicked);
            // 
            // browseForRegAsm
            // 
            this.browseForRegAsm.Location = new System.Drawing.Point(582, 193);
            this.browseForRegAsm.Name = "browseForRegAsm";
            this.browseForRegAsm.Size = new System.Drawing.Size(75, 23);
            this.browseForRegAsm.TabIndex = 4;
            this.browseForRegAsm.Text = "Browse...";
            this.browseForRegAsm.UseVisualStyleBackColor = true;
            this.browseForRegAsm.Click += new System.EventHandler(this.BrowseForRegAsmClicked);
            // 
            // comTextBox
            // 
            this.comTextBox.Location = new System.Drawing.Point(19, 136);
            this.comTextBox.Name = "comTextBox";
            this.comTextBox.Size = new System.Drawing.Size(557, 20);
            this.comTextBox.TabIndex = 5;
            // 
            // regAsmTextBox
            // 
            this.regAsmTextBox.Location = new System.Drawing.Point(19, 196);
            this.regAsmTextBox.Name = "regAsmTextBox";
            this.regAsmTextBox.Size = new System.Drawing.Size(557, 20);
            this.regAsmTextBox.TabIndex = 6;
            // 
            // menuStrip
            // 
            this.menuStrip.ImageScalingSize = new System.Drawing.Size(24, 24);
            this.menuStrip.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {this.aboutToolStripMenuItem});
            this.menuStrip.Location = new System.Drawing.Point(0, 0);
            this.menuStrip.Name = "menuStrip";
            this.menuStrip.Size = new System.Drawing.Size(671, 24);
            this.menuStrip.TabIndex = 7;
            this.menuStrip.Text = "menuStrip1";
            // 
            // aboutToolStripMenuItem
            // 
            this.aboutToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {this.howThisWorksToolStripMenuItem, this.aboutTheAuthorToolStripMenuItem});
            this.aboutToolStripMenuItem.Name = "aboutToolStripMenuItem";
            this.aboutToolStripMenuItem.Size = new System.Drawing.Size(44, 20);
            this.aboutToolStripMenuItem.Text = "Help";
            // 
            // howThisWorksToolStripMenuItem
            // 
            this.howThisWorksToolStripMenuItem.Name = "howThisWorksToolStripMenuItem";
            this.howThisWorksToolStripMenuItem.Size = new System.Drawing.Size(165, 22);
            this.howThisWorksToolStripMenuItem.Text = "How this works";
            this.howThisWorksToolStripMenuItem.Click += new System.EventHandler(this.HowThisWorksToolStripMenuItemClicked);
            // 
            // aboutTheAuthorToolStripMenuItem
            // 
            this.aboutTheAuthorToolStripMenuItem.Name = "aboutTheAuthorToolStripMenuItem";
            this.aboutTheAuthorToolStripMenuItem.Size = new System.Drawing.Size(165, 22);
            this.aboutTheAuthorToolStripMenuItem.Text = "About the author";
            this.aboutTheAuthorToolStripMenuItem.Click += new System.EventHandler(this.AboutTheAuthorToolStripMenuItemClicked);
            // 
            // descriptionLabel
            // 
            this.descriptionLabel.Location = new System.Drawing.Point(143, 79);
            this.descriptionLabel.Name = "descriptionLabel";
            this.descriptionLabel.Size = new System.Drawing.Size(433, 23);
            this.descriptionLabel.TabIndex = 8;
            this.descriptionLabel.Text = "This app registers a COM Assembly in the form of a .dll with the Windows Registry" + " Editor";
            // 
            // pathToCOMLabel
            // 
            this.pathToCOMLabel.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte) (0)));
            this.pathToCOMLabel.Location = new System.Drawing.Point(19, 114);
            this.pathToCOMLabel.Name = "pathToCOMLabel";
            this.pathToCOMLabel.Size = new System.Drawing.Size(182, 19);
            this.pathToCOMLabel.TabIndex = 9;
            this.pathToCOMLabel.Text = "Path to COM Assembly";
            // 
            // pathToRegAsmLabel
            // 
            this.pathToRegAsmLabel.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte) (0)));
            this.pathToRegAsmLabel.Location = new System.Drawing.Point(19, 174);
            this.pathToRegAsmLabel.Name = "pathToRegAsmLabel";
            this.pathToRegAsmLabel.Size = new System.Drawing.Size(182, 19);
            this.pathToRegAsmLabel.TabIndex = 10;
            this.pathToRegAsmLabel.Text = "Path to RegAsm.exe";
            // 
            // autoDetectionLabel
            // 
            this.autoDetectionLabel.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Italic, System.Drawing.GraphicsUnit.Point, ((byte) (0)));
            this.autoDetectionLabel.Location = new System.Drawing.Point(24, 219);
            this.autoDetectionLabel.Name = "autoDetectionLabel";
            this.autoDetectionLabel.Size = new System.Drawing.Size(177, 23);
            this.autoDetectionLabel.TabIndex = 11;
            this.autoDetectionLabel.Text = "RegAsm.exe autodetected";
            this.autoDetectionLabel.Visible = false;
            // 
            // COMRegistrationForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(671, 291);
            this.Controls.Add(this.autoDetectionLabel);
            this.Controls.Add(this.pathToRegAsmLabel);
            this.Controls.Add(this.pathToCOMLabel);
            this.Controls.Add(this.descriptionLabel);
            this.Controls.Add(this.regAsmTextBox);
            this.Controls.Add(this.comTextBox);
            this.Controls.Add(this.browseForRegAsm);
            this.Controls.Add(this.uninstallButton);
            this.Controls.Add(this.browseForCOMButton);
            this.Controls.Add(this.installButton);
            this.Controls.Add(this.headerLabel);
            this.Controls.Add(this.menuStrip);
            this.MainMenuStrip = this.menuStrip;
            this.Name = "COMRegistrationForm";
            this.Text = "COM Assembly Registration";
            this.menuStrip.ResumeLayout(false);
            this.menuStrip.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();
        }
        #endregion

        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void BrowseForComButtonClicked(object sender, EventArgs e) {
            OpenFileDialog openFileDialog = new OpenFileDialog {
                Filter = "Dynamic Linked Libraries (*.dll)|*.dll"
            };

            if (openFileDialog.ShowDialog().Equals(DialogResult.OK)) { //When the user presses "Ok"
                comTextBox.Text = openFileDialog.FileName;
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void BrowseForRegAsmClicked(object sender, EventArgs e) {
            OpenFileDialog openFileDialog = new OpenFileDialog {
                Filter = "Executables (*.exe)|*.exe"
            };

            if (openFileDialog.ShowDialog().Equals(DialogResult.OK)) { //When the user presses "Ok"
                regAsmTextBox.Text = openFileDialog.FileName;
            }
        }
        
        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void installButton_Click(object sender, EventArgs e) {
            //Check if the selected paths are valid paths
            if (!File.Exists(comTextBox.Text) && !File.Exists(regAsmTextBox.Text)) {
                MessageBox.Show("Invalid paths for both a dll and RegAsm.exe!",
                                "Invalid Paths Listed", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            } else if (!File.Exists(comTextBox.Text)) {
                MessageBox.Show("Invalid path a dll!",
                                "Invalid Paths Listed", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            } else if (!File.Exists(regAsmTextBox.Text)) {
                MessageBox.Show("Invalid paths for RegAsm.exe!",
                                "Invalid Paths Listed", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            
            // Run the installation command: "RegAsm.exe /codebase "<absolute or relative path of .dll file>"
            System.Diagnostics.Process process = new System.Diagnostics.Process();
            System.Diagnostics.ProcessStartInfo startInfo = new System.Diagnostics.ProcessStartInfo {
                WindowStyle = System.Diagnostics.ProcessWindowStyle.Hidden, //So Command Prompt doesn't open visibly
                FileName = "cmd.exe",
                Arguments = $@"/C {regAsmTextBox.Text} /codebase {comTextBox.Text}", //DO NOT FORGET "/C" in beginning
                Verb = "runas" //Runs command prompt as an Administrator
            };
            process.StartInfo = startInfo;
            process.Start(); //Execute
            
            //Close this WinForm
            this.Close();
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void UninstallButtonClicked(object sender, EventArgs e) {
            //Check if the selected paths are valid paths
            if (!File.Exists(comTextBox.Text) && !File.Exists(regAsmTextBox.Text)) {
                MessageBox.Show("Invalid paths for both a dll and RegAsm.exe!",
                                "Invalid Paths Listed", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            } else if (!File.Exists(comTextBox.Text)) {
                MessageBox.Show("Invalid path a dll!",
                                "Invalid Paths Listed", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            } else if (!File.Exists(regAsmTextBox.Text)) {
                MessageBox.Show("Invalid paths for RegAsm.exe!",
                                "Invalid Paths Listed", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            
            //Warn the use if they want to uninstall
            DialogResult dialogResult = MessageBox.Show("Are you sure you want to uninstall this dll? Doing so will make whatever application it is intended for not be able to see it",
                                                                         "Explicit permission needed", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
            if (dialogResult == DialogResult.No) {  //if you the user selects no, do nothing
                return;
            }
            
            //Uninstall the COM Assembly
            System.Diagnostics.Process process = new System.Diagnostics.Process();
            System.Diagnostics.ProcessStartInfo startInfo = new System.Diagnostics.ProcessStartInfo {
                WindowStyle = System.Diagnostics.ProcessWindowStyle.Hidden, //So Command Prompt doesn't open visibly
                FileName = "cmd.exe",
                Arguments = $@"/C {regAsmTextBox.Text} /unregister {comTextBox.Text}", //DO NOT FORGET "/C" in beginning
                Verb = "runas" //Runs command prompt as an Administrator
            };
            process.StartInfo = startInfo;
            process.Start(); //Execute
            
            //Close this WinForm
            this.Close();
        }
        
        private void HowThisWorksToolStripMenuItemClicked(object sender, EventArgs e) {
            new HowItWorksForm().ShowDialog();
        }

        private void AboutTheAuthorToolStripMenuItemClicked(object sender, EventArgs e) {
            new AboutTheAuthorForm().ShowDialog();
        }
    }
}