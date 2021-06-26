using System;
using System.Collections.Generic;
using System.Diagnostics.Eventing.Reader;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Security;

namespace SolidWorks_WinForm_Creation {
    partial class CopyDisplayStatesForm {
        private OpenFileDialog openFileDialog;
        private string copiedFilePath;
        private string pastedFilePath;

        public CopyDisplayStatesForm() {
            copiedFilePath = null;
            pastedFilePath = null;
            InitializeComponent();
        }

        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main() {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new CopyDisplayStatesForm());
        }

        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

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

        #region Window Components

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent() {
            this.fromConfigLabel = new System.Windows.Forms.Label();
            this.copiedConfigComboBox = new System.Windows.Forms.ComboBox();
            this.button = new System.Windows.Forms.Button();
            this.toConfigLabel = new System.Windows.Forms.Label();
            this.pastedConfigComboBox = new System.Windows.Forms.ComboBox();
            this.configurationTextFromLabel = new System.Windows.Forms.Label();
            this.configurationTextPastedLabel = new System.Windows.Forms.Label();
            this.browseForPastedConfigDocumentButton = new System.Windows.Forms.Button();
            this.browseForCopiedConfigDocumentButton = new System.Windows.Forms.Button();
            this.nameOfCopiedDocumentLabel = new System.Windows.Forms.Label();
            this.nameOfPastedDocumentLabel = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // fromConfigLabel
            // 
            this.fromConfigLabel.Font = new System.Drawing.Font("Century Gothic", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.fromConfigLabel.Location = new System.Drawing.Point(12, 9);
            this.fromConfigLabel.Name = "fromConfigLabel";
            this.fromConfigLabel.Size = new System.Drawing.Size(425, 33);
            this.fromConfigLabel.TabIndex = 0;
            this.fromConfigLabel.Text = "Part or Assembly file to choose a configuration to copy display states FROM: ";
            this.fromConfigLabel.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // copiedConfigComboBox
            // 
            this.copiedConfigComboBox.FormattingEnabled = true;
            this.copiedConfigComboBox.Location = new System.Drawing.Point(105, 73);
            this.copiedConfigComboBox.Name = "copiedConfigComboBox";
            this.copiedConfigComboBox.Size = new System.Drawing.Size(458, 20);
            this.copiedConfigComboBox.TabIndex = 1;
            // 
            // button
            // 
            this.button.Font = new System.Drawing.Font("Century Gothic", 7.8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.button.Location = new System.Drawing.Point(191, 249);
            this.button.Name = "button";
            this.button.Size = new System.Drawing.Size(197, 35);
            this.button.TabIndex = 2;
            this.button.Text = "Copy Display States";
            this.button.UseVisualStyleBackColor = true;
            // 
            // toConfigLabel
            // 
            this.toConfigLabel.Font = new System.Drawing.Font("Century Gothic", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.toConfigLabel.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.toConfigLabel.Location = new System.Drawing.Point(12, 153);
            this.toConfigLabel.Name = "toConfigLabel";
            this.toConfigLabel.Size = new System.Drawing.Size(469, 26);
            this.toConfigLabel.TabIndex = 3;
            this.toConfigLabel.Text = "Part or Assembly file to choose a configuration to copy display states TO: \r\n";
            this.toConfigLabel.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // pastedConfigComboBox
            // 
            this.pastedConfigComboBox.FormattingEnabled = true;
            this.pastedConfigComboBox.Location = new System.Drawing.Point(105, 214);
            this.pastedConfigComboBox.Name = "pastedConfigComboBox";
            this.pastedConfigComboBox.Size = new System.Drawing.Size(458, 20);
            this.pastedConfigComboBox.TabIndex = 4;
            // 
            // configurationTextFromLabel
            // 
            this.configurationTextFromLabel.AutoSize = true;
            this.configurationTextFromLabel.Font = new System.Drawing.Font("Century Gothic", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.configurationTextFromLabel.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.configurationTextFromLabel.Location = new System.Drawing.Point(12, 74);
            this.configurationTextFromLabel.Name = "configurationTextFromLabel";
            this.configurationTextFromLabel.Size = new System.Drawing.Size(87, 16);
            this.configurationTextFromLabel.TabIndex = 5;
            this.configurationTextFromLabel.Text = "Configuration: ";
            // 
            // configurationTextPastedLabel
            // 
            this.configurationTextPastedLabel.AutoSize = true;
            this.configurationTextPastedLabel.Font = new System.Drawing.Font("Century Gothic", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.configurationTextPastedLabel.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.configurationTextPastedLabel.Location = new System.Drawing.Point(12, 215);
            this.configurationTextPastedLabel.Name = "configurationTextPastedLabel";
            this.configurationTextPastedLabel.Size = new System.Drawing.Size(87, 16);
            this.configurationTextPastedLabel.TabIndex = 6;
            this.configurationTextPastedLabel.Text = "Configuration: ";
            // 
            // browseForPastedConfigDocumentButton
            // 
            this.browseForPastedConfigDocumentButton.Font = new System.Drawing.Font("Century Gothic", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.browseForPastedConfigDocumentButton.Location = new System.Drawing.Point(433, 145);
            this.browseForPastedConfigDocumentButton.Name = "browseForPastedConfigDocumentButton";
            this.browseForPastedConfigDocumentButton.Size = new System.Drawing.Size(130, 27);
            this.browseForPastedConfigDocumentButton.TabIndex = 7;
            this.browseForPastedConfigDocumentButton.Text = "Browse...";
            this.browseForPastedConfigDocumentButton.UseVisualStyleBackColor = true;
            this.browseForPastedConfigDocumentButton.Click += new System.EventHandler(this.browseForPastedConfigDocumentButton_Click);
            // 
            // browseForCopiedConfigDocumentButton
            // 
            this.browseForCopiedConfigDocumentButton.Font = new System.Drawing.Font("Century Gothic", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.browseForCopiedConfigDocumentButton.Location = new System.Drawing.Point(433, 9);
            this.browseForCopiedConfigDocumentButton.Name = "browseForCopiedConfigDocumentButton";
            this.browseForCopiedConfigDocumentButton.Size = new System.Drawing.Size(130, 27);
            this.browseForCopiedConfigDocumentButton.TabIndex = 8;
            this.browseForCopiedConfigDocumentButton.Text = "Browse...";
            this.browseForCopiedConfigDocumentButton.UseVisualStyleBackColor = true;
            this.browseForCopiedConfigDocumentButton.Click += new System.EventHandler(this.browseForCopiedConfigDocumentButton_Click);
            // 
            // nameOfCopiedDocumentLabel
            // 
            this.nameOfCopiedDocumentLabel.Font = new System.Drawing.Font("Century Gothic", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.nameOfCopiedDocumentLabel.ForeColor = System.Drawing.Color.Red;
            this.nameOfCopiedDocumentLabel.Location = new System.Drawing.Point(15, 37);
            this.nameOfCopiedDocumentLabel.Name = "nameOfCopiedDocumentLabel";
            this.nameOfCopiedDocumentLabel.Size = new System.Drawing.Size(548, 33);
            this.nameOfCopiedDocumentLabel.TabIndex = 11;
            this.nameOfCopiedDocumentLabel.Text = "null_copied";
            this.nameOfCopiedDocumentLabel.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            this.nameOfCopiedDocumentLabel.Visible = false;
            // 
            // nameOfPastedDocumentLabel
            // 
            this.nameOfPastedDocumentLabel.Font = new System.Drawing.Font("Century Gothic", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.nameOfPastedDocumentLabel.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(192)))), ((int)(((byte)(0)))));
            this.nameOfPastedDocumentLabel.Location = new System.Drawing.Point(15, 175);
            this.nameOfPastedDocumentLabel.Name = "nameOfPastedDocumentLabel";
            this.nameOfPastedDocumentLabel.Size = new System.Drawing.Size(548, 36);
            this.nameOfPastedDocumentLabel.TabIndex = 12;
            this.nameOfPastedDocumentLabel.Text = "null_pasted";
            this.nameOfPastedDocumentLabel.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            this.nameOfPastedDocumentLabel.Visible = false;
            // 
            // CopyDisplayStatesForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(575, 293);
            this.Controls.Add(this.nameOfPastedDocumentLabel);
            this.Controls.Add(this.nameOfCopiedDocumentLabel);
            this.Controls.Add(this.browseForCopiedConfigDocumentButton);
            this.Controls.Add(this.browseForPastedConfigDocumentButton);
            this.Controls.Add(this.configurationTextPastedLabel);
            this.Controls.Add(this.configurationTextFromLabel);
            this.Controls.Add(this.pastedConfigComboBox);
            this.Controls.Add(this.toConfigLabel);
            this.Controls.Add(this.button);
            this.Controls.Add(this.copiedConfigComboBox);
            this.Controls.Add(this.fromConfigLabel);
            this.Font = new System.Drawing.Font("Times New Roman", 7.8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.Name = "CopyDisplayStatesForm";
            this.Text = "Copy Display States";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        private Button button;
        private ComboBox copiedConfigComboBox;
        private ComboBox pastedConfigComboBox;
        private Label fromConfigLabel;
        private Label toConfigLabel;
        private Label configurationTextFromLabel;
        private Label configurationTextPastedLabel;
        private Button browseForPastedConfigDocumentButton;
        private Button browseForCopiedConfigDocumentButton;
        private Label nameOfCopiedDocumentLabel;
        private Label nameOfPastedDocumentLabel;

        #endregion


        /// <summary>
        /// Creates and openFileDialog for the user to select either a SolidWorks part file or SolidWorks assembly file.
        /// If the pasted config button has a file selected, only the file type of that selection will be available to
        /// the user of this button
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void browseForCopiedConfigDocumentButton_Click(object sender, EventArgs e) {
            openFileDialog = new OpenFileDialog(); //creates a dialog box for the user to select a file
            //We want to make sure that file types for both configurations are the same. If one file is a .SLDASM, then the
            //other should be as well (and the same goes for .SLDPRT). Note ToUpper() might be faster than ToLower()
            if (nameOfPastedDocumentLabel.Visible) {
                if (pastedFilePath.Substring(pastedFilePath.Length - 7).ToUpper().Equals(".SLDPRT")) { //if the other file is a .SLDPRT
                    openFileDialog.Filter = "SolidWorks Parts (*.SLDPRT)|*.SLDPRT";
                } else { //The other file is a .SLDASM
                    openFileDialog.Filter = "SolidWorks Assemblies (*.SLDASM)|*.SLDASM";
                }
            } else { //else nothing has been selected yet.
                openFileDialog.Filter = "SolidWorks Parts (*.SLDPRT)|*.SLDPRT|SolidWorks Assemblies (*.SLDASM)|*.SLDASM";
            }


            if (openFileDialog.ShowDialog().Equals(DialogResult.OK)) { //When the user presses "Ok"
                try {
                    copiedFilePath = openFileDialog.FileName;
                    nameOfCopiedDocumentLabel.Text = copiedFilePath;
                    nameOfCopiedDocumentLabel.Visible = true;
                } catch (SecurityException exception) { //I don't think we'll ever get here
                    MessageBox.Show("You do not have permission to open this file!",
                        "Copy Display States - No Permission", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        /// <summary>
        /// Creates and openFileDialog for the user to select either a SolidWorks part file or SolidWorks assembly file.
        /// If the copied config button has a file selected, only the file type of that selection will be available to
        /// the user of this button
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void browseForPastedConfigDocumentButton_Click(object sender, EventArgs e) {
            openFileDialog = new OpenFileDialog(); //creates a dialog box for the user to select a file

            //We want to make sure that file types for both configurations are the same. If one file is a .SLDASM, then the
            //other should be as well (and the same goes for .SLDPRT). Note ToUpper() might be faster than ToLower()
            if (nameOfCopiedDocumentLabel.Visible) {
                if (copiedFilePath.Substring(copiedFilePath.Length - 7).ToUpper().Equals(".SLDPRT")) { //if the other file is a .SLDPRT
                    openFileDialog.Filter = "SolidWorks Parts (*.SLDPRT)|*.SLDPRT";
                } else { //The other file is a .SLDASM
                    openFileDialog.Filter = "SolidWorks Assemblies (*.SLDASM)|*.SLDASM";
                }
            } else {  //else nothing has been selected yet.
                openFileDialog.Filter = "SolidWorks Parts (*.SLDPRT)|*.SLDPRT|SolidWorks Assemblies (*.SLDASM)|*.SLDASM";
            }


            if (openFileDialog.ShowDialog().Equals(DialogResult.OK)) { //When the user presses "Ok"
                try {
                    pastedFilePath = openFileDialog.FileName;
                    nameOfPastedDocumentLabel.Text = pastedFilePath;
                    nameOfPastedDocumentLabel.Visible = true;
                } catch (SecurityException exception) { //I don't think we'll ever get here
                    MessageBox.Show("You do not have permission to open this file!",
                        "Copy Display States - No Permission", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }

        }
    }
}

