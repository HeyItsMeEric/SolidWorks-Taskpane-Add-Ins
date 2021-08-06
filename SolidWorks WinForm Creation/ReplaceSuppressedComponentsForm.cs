using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;


namespace SolidWorks_WinForm_Creation {
    class ReplaceSuppressedComponentsForm : Form {

        public ReplaceSuppressedComponentsForm() {
            InitializeComponent();

            //Centering the Form in the middle of the screen
            this.Location = new System.Drawing.Point((Screen.FromControl(this).Bounds.Width - this.Width) / 2,
                (Screen.FromControl(this).Bounds.Height / 7));
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

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent() {
            this.suppressedComponentComboBox = new System.Windows.Forms.ComboBox();
            this.ReplacementComponentComboBox = new System.Windows.Forms.ComboBox();
            this.suppressedComponentLabel = new System.Windows.Forms.Label();
            this.replacementComponentLabel = new System.Windows.Forms.Label();
            this.addReplacementLabel = new System.Windows.Forms.Label();
            this.suppressionTable = new System.Windows.Forms.DataGridView();
            this.runButton = new System.Windows.Forms.Button();
            this.cancelButton = new System.Windows.Forms.Button();
            this.addReplacementButton = new System.Windows.Forms.Button();
            this.deleteReplacementRow = new System.Windows.Forms.Button();
            this.runWithoutThisFormLabel = new System.Windows.Forms.Button();
            this.deleteReplacementLabel = new System.Windows.Forms.Label();
            ((System.ComponentModel.ISupportInitialize)(this.suppressionTable)).BeginInit();
            this.SuspendLayout();
            // 
            // suppressedComponentComboBox
            // 
            this.suppressedComponentComboBox.FormattingEnabled = true;
            this.suppressedComponentComboBox.Location = new System.Drawing.Point(27, 56);
            this.suppressedComponentComboBox.Name = "suppressedComponentComboBox";
            this.suppressedComponentComboBox.Size = new System.Drawing.Size(415, 21);
            this.suppressedComponentComboBox.TabIndex = 0;
            // 
            // ReplacementComponentComboBox
            // 
            this.ReplacementComponentComboBox.FormattingEnabled = true;
            this.ReplacementComponentComboBox.Location = new System.Drawing.Point(537, 56);
            this.ReplacementComponentComboBox.Name = "ReplacementComponentComboBox";
            this.ReplacementComponentComboBox.Size = new System.Drawing.Size(398, 21);
            this.ReplacementComponentComboBox.TabIndex = 1;
            // 
            // suppressedComponentLabel
            // 
            this.suppressedComponentLabel.AutoSize = true;
            this.suppressedComponentLabel.Location = new System.Drawing.Point(24, 38);
            this.suppressedComponentLabel.Name = "suppressedComponentLabel";
            this.suppressedComponentLabel.Size = new System.Drawing.Size(200, 13);
            this.suppressedComponentLabel.TabIndex = 2;
            this.suppressedComponentLabel.Text = "Suppressed component in configuration: ";
            // 
            // replacementComponentLabel
            // 
            this.replacementComponentLabel.AutoSize = true;
            this.replacementComponentLabel.Location = new System.Drawing.Point(534, 38);
            this.replacementComponentLabel.Name = "replacementComponentLabel";
            this.replacementComponentLabel.Size = new System.Drawing.Size(92, 13);
            this.replacementComponentLabel.TabIndex = 3;
            this.replacementComponentLabel.Text = "Replace  in   with ";
            // 
            // addReplacementLabel
            // 
            this.addReplacementLabel.AutoSize = true;
            this.addReplacementLabel.Location = new System.Drawing.Point(24, 106);
            this.addReplacementLabel.Name = "addReplacementLabel";
            this.addReplacementLabel.Size = new System.Drawing.Size(371, 13);
            this.addReplacementLabel.TabIndex = 4;
            this.addReplacementLabel.Text = "Once both drop-down menus above are populated, click \"Add Replacement\"";
            // 
            // suppressionTable
            // 
            this.suppressionTable.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.suppressionTable.Location = new System.Drawing.Point(27, 164);
            this.suppressionTable.Name = "suppressionTable";
            this.suppressionTable.Size = new System.Drawing.Size(908, 447);
            this.suppressionTable.TabIndex = 5;
            // 
            // runButton
            // 
            this.runButton.Location = new System.Drawing.Point(844, 643);
            this.runButton.Name = "runButton";
            this.runButton.Size = new System.Drawing.Size(121, 45);
            this.runButton.TabIndex = 6;
            this.runButton.Text = "Copy Display States";
            this.runButton.UseVisualStyleBackColor = true;
            // 
            // cancelButton
            // 
            this.cancelButton.Location = new System.Drawing.Point(591, 654);
            this.cancelButton.Name = "cancelButton";
            this.cancelButton.Size = new System.Drawing.Size(75, 23);
            this.cancelButton.TabIndex = 7;
            this.cancelButton.Text = "Cancel";
            this.cancelButton.UseVisualStyleBackColor = true;
            // 
            // addReplacementButton
            // 
            this.addReplacementButton.Location = new System.Drawing.Point(537, 101);
            this.addReplacementButton.Name = "addReplacementButton";
            this.addReplacementButton.Size = new System.Drawing.Size(164, 23);
            this.addReplacementButton.TabIndex = 9;
            this.addReplacementButton.Text = "Add Replacement";
            this.addReplacementButton.UseVisualStyleBackColor = true;
            // 
            // deleteReplacementRow
            // 
            this.deleteReplacementRow.Location = new System.Drawing.Point(27, 654);
            this.deleteReplacementRow.Name = "deleteReplacementRow";
            this.deleteReplacementRow.Size = new System.Drawing.Size(75, 23);
            this.deleteReplacementRow.TabIndex = 10;
            this.deleteReplacementRow.Text = "Delete Replacement";
            this.deleteReplacementRow.UseVisualStyleBackColor = true;
            // 
            // runWithoutThisFormLabel
            // 
            this.runWithoutThisFormLabel.Location = new System.Drawing.Point(701, 643);
            this.runWithoutThisFormLabel.Name = "runWithoutThisFormLabel";
            this.runWithoutThisFormLabel.Size = new System.Drawing.Size(113, 45);
            this.runWithoutThisFormLabel.TabIndex = 11;
            this.runWithoutThisFormLabel.Text = "Run Without Replacement Parts";
            this.runWithoutThisFormLabel.UseVisualStyleBackColor = true;
            // 
            // deleteReplacementLabel
            // 
            this.deleteReplacementLabel.AutoSize = true;
            this.deleteReplacementLabel.Location = new System.Drawing.Point(24, 614);
            this.deleteReplacementLabel.Name = "deleteReplacementLabel";
            this.deleteReplacementLabel.Size = new System.Drawing.Size(41, 13);
            this.deleteReplacementLabel.TabIndex = 12;
            this.deleteReplacementLabel.Text = "werwer";
            // 
            // ReplaceSuppressedComponentsForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(977, 700);
            this.Controls.Add(this.deleteReplacementLabel);
            this.Controls.Add(this.runWithoutThisFormLabel);
            this.Controls.Add(this.deleteReplacementRow);
            this.Controls.Add(this.addReplacementButton);
            this.Controls.Add(this.cancelButton);
            this.Controls.Add(this.runButton);
            this.Controls.Add(this.suppressionTable);
            this.Controls.Add(this.addReplacementLabel);
            this.Controls.Add(this.replacementComponentLabel);
            this.Controls.Add(this.suppressedComponentLabel);
            this.Controls.Add(this.ReplacementComponentComboBox);
            this.Controls.Add(this.suppressedComponentComboBox);
            this.Name = "ReplaceSuppressedComponentsForm";
            this.Text = "ReplaceSuppressedComponentsForm";
            ((System.ComponentModel.ISupportInitialize)(this.suppressionTable)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        private System.Windows.Forms.ComboBox suppressedComponentComboBox;
        private System.Windows.Forms.ComboBox ReplacementComponentComboBox;
        private System.Windows.Forms.Label suppressedComponentLabel;
        private System.Windows.Forms.Label replacementComponentLabel;
        private System.Windows.Forms.Label addReplacementLabel;
        private System.Windows.Forms.DataGridView suppressionTable;
        private System.Windows.Forms.Button runButton;
        private System.Windows.Forms.Button cancelButton;
        private System.Windows.Forms.Button addReplacementButton;
        private System.Windows.Forms.Button deleteReplacementRow;
        private System.Windows.Forms.Button runWithoutThisFormLabel;
        private System.Windows.Forms.Label deleteReplacementLabel;

        #endregion
    }
}