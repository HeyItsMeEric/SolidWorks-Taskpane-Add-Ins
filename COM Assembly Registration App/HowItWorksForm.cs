/*
 * A WinForm class with info about how registering a COM Assembly works.
 *
 * @author        Eric Gustafson
 * @date_created  July 12, 2021, 08:03:08 PM
 * @version       1.0
 */

using System;
using System.ComponentModel;
using System.Windows.Forms;

namespace COM_Assembly_Registration_App {
    public class HowItWorksForm : Form {
        public HowItWorksForm() {
            InitializeComponent();
            
            //Centering the Form in the middle of the screen
            this.Location = new System.Drawing.Point((Screen.FromControl(this).Bounds.Width - this.Width) / 2,
                                                     (Screen.FromControl(this).Bounds.Height / 7));
        }
        
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private IContainer components = null;
        
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(HowItWorksForm));
            this.label = new System.Windows.Forms.Label();
            this.linkLabel = new System.Windows.Forms.LinkLabel();
            this.SuspendLayout();
            // 
            // label
            // 
            this.label.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte) (0)));
            this.label.Location = new System.Drawing.Point(12, 28);
            this.label.Name = "label";
            this.label.Size = new System.Drawing.Size(972, 456);
            this.label.TabIndex = 0;
            this.label.Text = resources.GetString("label.Text");
            // 
            // linkLabel
            // 
            this.linkLabel.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte) (0)));
            this.linkLabel.Location = new System.Drawing.Point(526, 427);
            this.linkLabel.Name = "linkLabel";
            this.linkLabel.Size = new System.Drawing.Size(183, 23);
            this.linkLabel.TabIndex = 1;
            this.linkLabel.TabStop = true;
            this.linkLabel.Text = "this YouTube video.";
            this.linkLabel.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.linkLabel_LinkClicked);
            // 
            // HowItWorksForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(993, 457);
            this.Controls.Add(this.linkLabel);
            this.Controls.Add(this.label);
            this.Name = "HowItWorksForm";
            this.Text = "How It Works";
            this.ResumeLayout(false);
        }

        private System.Windows.Forms.LinkLabel linkLabel;
        private System.Windows.Forms.Label label;

        #endregion
        /// <summary>
        /// Opens my LinkedIn profile in the user's default browser
        /// 
        /// Method Contents from 
        /// https://docs.microsoft.com/en-us/dotnet/api/system.windows.forms.linklabel?view=net-5.0#examples
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void linkLabel_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e) {
            // Specify that the link was visited.
            this.linkLabel.LinkVisited = true;

            // Navigate to a URL.
            System.Diagnostics.Process.Start(@"https://www.youtube.com/watch?v=7DlG6OQeJP0");
        }
    }
}