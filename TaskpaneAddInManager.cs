/*
 * This class creates the UI that goes inside the Taskpane of SolidWorks. The taskpane in question is the pane on
 * the right side of the application (by default) that has 
 *
 * @author        Eric Gustafson
 * @date_created  June 22, 2021, 01:25:37 AM
 * @version       1.0
 */

using System.ComponentModel;
using System;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Gustafson.SolidWorks.TaskpaneAddIns;

namespace Gustafson.SolidWorks.TaskpaneAddIns {

    [ProgId("Gustafson.SolidWorks.TaskpaneAddIns")]
    public class TaskpaneAddInManager : UserControl {
        public TaskpaneAddInManager() {
            InitializeComponent();
        }


        /// <summary> 
        /// Required designer variable.
        /// </summary>
        private readonly IContainer components = null;

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

        #region Component Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent() {
            this.TaskpaneCopyDisplayStatesButton = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // TaskpaneCopyDisplayStatesButton
            // 
            this.TaskpaneCopyDisplayStatesButton.Location = new System.Drawing.Point(9, 12);
            this.TaskpaneCopyDisplayStatesButton.Name = "TaskpaneCopyDisplayStatesButton";
            this.TaskpaneCopyDisplayStatesButton.Size = new System.Drawing.Size(157, 57);
            this.TaskpaneCopyDisplayStatesButton.TabIndex = 0;
            this.TaskpaneCopyDisplayStatesButton.Text = "Copy Display States";
            this.TaskpaneCopyDisplayStatesButton.UseVisualStyleBackColor = true;
            this.TaskpaneCopyDisplayStatesButton.Click += new System.EventHandler(this.CopyDisplayStatesButtonClicked);
            // 
            // TaskpaneIntegration
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.TaskpaneCopyDisplayStatesButton);
            this.Name = "TaskpaneIntegration";
            this.Size = new System.Drawing.Size(444, 811);
            this.ResumeLayout(false);
        }

        private System.Windows.Forms.Button TaskpaneCopyDisplayStatesButton;
        private System.Windows.Forms.Button copyDisplayStateButton;

        #endregion


        private void CopyDisplayStatesButtonClicked(object sender, EventArgs e) {
            //Application.SetHighDpiMode(HighDpiMode.SystemAware);
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new CopyDisplayStatesForm());
        }
    }
}