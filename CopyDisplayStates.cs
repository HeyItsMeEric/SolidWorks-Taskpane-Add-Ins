/*
 * 
 *
 * @author        Eric Gustafson
 * @date_created  June 20, 2021, 04:37:42 PM
 * @version       1.0
 */

using SolidWorks.Interop.swpublished;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using System.Security;

using System.Runtime.InteropServices;
using Exception = System.Exception;


namespace Gustafson.SolidWorks.TaskpaneAddIns {
    /**
     * For the dialog box that pops-up
     *
     * https://bit.ly/3A3tpQe //Get Display State Names and Visibilities of Components Example
     * https://bit.ly/2TX5npx //Get List of Configurations Example
     */
    internal class CopyDisplayStatesForm : Form {
        //Win Form instance variables
        private OpenFileDialog openFileDialog;
        private string copiedFilePath;
        private string pastedFilePath;

        //SolidWorks API instance variables
        private Configuration[] copiedConfigArr;
        private Configuration[] pastedConfigArr;
        private ModelDoc2 copiedDoc;
        private ModelDoc2 pastedDoc;
        


        public CopyDisplayStatesForm() {
            copiedFilePath = null;
            pastedFilePath = null;
            InitializeComponent();
            //StreamWriter debugger = TaskpaneExtensionsAddIn.debugger;
            #region Debug stuff


           /* try {
                debugger.WriteLine($"{DateTime.Now}: {TaskpaneExtensionsAddIn.mySolidWorks.ToString()}");
            } catch (Exception e) {
                debugger.WriteLine(e.Message);
            }*/

            #endregion
            

            //Check to see if there is an active document in SolidWorks. If so, update the output labels (labels that display which files to look at configs from) to
            //display the file path of the currently loaded document. if there is no current document or the user browses for a new document (file) then the
            //button_pressed methods in this class that take care of that will traverse the file and find the configurations.
            if (TaskpaneExtensionsAddIn.mySolidWorks.ActiveDoc != null) {
                copiedDoc = (ModelDoc2) TaskpaneExtensionsAddIn.mySolidWorks.ActiveDoc;
                pastedDoc = (ModelDoc2)TaskpaneExtensionsAddIn.mySolidWorks.ActiveDoc;
                copiedFilePath = pastedFilePath = copiedDoc.GetPathName(); //Set the file paths of the copied and pasted documents to the active document
                nameOfCopiedDocumentLabel.Text = copiedFilePath; //Set the text of the output labels
                nameOfPastedDocumentLabel.Text = pastedFilePath;
                nameOfCopiedDocumentLabel.Visible = nameOfPastedDocumentLabel.Visible = true; //Set the labels' visibility to true

                #region Traverse document for configuration names
                    string[] listOfConfigNames = (string[]) copiedDoc.GetConfigurationNames(); //Get string array where each element is a configuration name
                    copiedConfigArr = new Configuration[listOfConfigNames.Length]; //instantiate private class member copiedConfigArr to hold the configurations of
                    pastedConfigArr = new Configuration[listOfConfigNames.Length];
                    int index = -1;
                    foreach (string configName in listOfConfigNames) {
                        copiedConfigArr[++index] = (Configuration) copiedDoc.GetConfigurationByName(configName); //add config to both config arrays
                        pastedConfigArr[index] = copiedConfigArr[index]; 
                        copiedConfigComboBox.Items.Add(configName); //update the combo boxes (drop down menus)
                        pastedConfigComboBox.Items.Add(configName);
                    }
                #endregion
            } //else keep the output labels not visible. Make the user browse for files.
            //Default selections of combo boxes
            copiedConfigComboBox.SelectedIndex = 0;
            pastedConfigComboBox.SelectedIndex = 0;
        }

        /// <summary>
        ///  Required designer variable.
        /// </summary>
        private readonly System.ComponentModel.IContainer components = null;

        /// <summary>
        ///  Clean up any resources being used.
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
            this.copyButton = new System.Windows.Forms.Button();
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
            // copyButton
            // 
            this.copyButton.Font = new System.Drawing.Font("Century Gothic", 7.8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.copyButton.Location = new System.Drawing.Point(191, 249);
            this.copyButton.Name = "copyButton";
            this.copyButton.Size = new System.Drawing.Size(197, 35);
            this.copyButton.TabIndex = 2;
            this.copyButton.Text = "Copy Display States";
            this.copyButton.UseVisualStyleBackColor = true;
            this.copyButton.Click += new System.EventHandler(this.ButtonClicked);
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
            this.browseForPastedConfigDocumentButton.Click += new System.EventHandler(this.BrowseForPastedConfigDocumentButtonClick);
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
            this.browseForCopiedConfigDocumentButton.Click += new System.EventHandler(this.BrowseForCopiedConfigDocumentButtonClick);
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
            this.Controls.Add(this.copyButton);
            this.Controls.Add(this.copiedConfigComboBox);
            this.Controls.Add(this.fromConfigLabel);
            this.Font = new System.Drawing.Font("Times New Roman", 7.8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.Name = "CopyDisplayStatesForm";
            this.Text = "Copy Display States";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        private Button copyButton;
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
        /// If the pasted config copyButton has a file selected, only the file type of that selection will be available to
        /// the user of this copyButton
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void BrowseForCopiedConfigDocumentButtonClick(object sender, EventArgs e) {




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
                } catch (SecurityException exception) { //I don't think we'll ever get here, but this is thrown if we don't have access to open this file
                    MessageBox.Show("You do not have permission to open this file!",
                        "Copy Display States - No Permission", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }

            //Copy the configurations of the new file into the combo boxes
            DocumentSpecification swDocSpecification = (DocumentSpecification)TaskpaneExtensionsAddIn.mySolidWorks.GetOpenDocSpec(copiedFilePath); 
            copiedDoc = TaskpaneExtensionsAddIn.mySolidWorks.OpenDoc7(swDocSpecification); //

            string[] listOfConfigNames = (string[]) copiedDoc.GetConfigurationNames(); //Get string array where each element is a configuration name
            copiedConfigArr = new Configuration[listOfConfigNames .Length]; //instantiate private class member copiedConfigArr to hold the configurations of
            int index = -1;
            foreach (string configName in listOfConfigNames) {
                copiedConfigArr[++index] = (Configuration) copiedDoc.GetConfigurationByName(configName); //add config to copied config arrays
                copiedConfigComboBox.Items.Add(configName); //update the combo box (drop down menus)
            }
            copiedConfigComboBox.SelectedIndex = 0;
        }

        /// <summary>
        /// Creates and openFileDialog for the user to select either a SolidWorks part file or SolidWorks assembly file.
        /// If the copied config copyButton has a file selected, only the file type of that selection will be available to
        /// the user of this copyButton
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void BrowseForPastedConfigDocumentButtonClick(object sender, EventArgs e) {
            openFileDialog = new OpenFileDialog(); //creates a dialog box for the user to select a file

            //We want to make sure that file types for both configurations are the same. If one file is a .SLDASM, then the
            //other should be as well (and the same goes for .SLDPRT). Note ToUpper() might be faster than ToLower()
            if (nameOfCopiedDocumentLabel.Visible) {
                if (copiedFilePath.Substring(copiedFilePath.Length - 7).ToUpper().Equals(".SLDPRT")) {
                    //if the other file is a .SLDPRT
                    openFileDialog.Filter = "SolidWorks Parts (*.SLDPRT)|*.SLDPRT";
                }
                else { //The other file is a .SLDASM
                    openFileDialog.Filter = "SolidWorks Assemblies (*.SLDASM)|*.SLDASM";
                }
            }
            else { //else nothing has been selected yet.
                openFileDialog.Filter =
                    "SolidWorks Parts (*.SLDPRT)|*.SLDPRT|SolidWorks Assemblies (*.SLDASM)|*.SLDASM";
            }


            if (openFileDialog.ShowDialog().Equals(DialogResult.OK)) { //When the user presses "Ok"
                try {
                    pastedFilePath = openFileDialog.FileName;
                    nameOfPastedDocumentLabel.Text = pastedFilePath;
                    nameOfPastedDocumentLabel.Visible = true;
                }
                catch (SecurityException exception) { //I don't think we'll ever get here
                    MessageBox.Show("You do not have permission to open this file!",
                        "Copy Display States - No Permission", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }

            //Copy the configurations of the new file into the combo boxes
            DocumentSpecification swDocSpecification = (DocumentSpecification) TaskpaneExtensionsAddIn.mySolidWorks.GetOpenDocSpec(pastedFilePath);
            pastedDoc = TaskpaneExtensionsAddIn.mySolidWorks.OpenDoc7(swDocSpecification); //

            string[] listOfConfigNames = (string[]) pastedDoc.GetConfigurationNames(); //Get string array where each element is a configuration name
            pastedConfigArr = new Configuration[listOfConfigNames.Length]; //instantiate private class member copiedConfigArr to hold the configurations of
            int index = -1;
            foreach (string configName in listOfConfigNames) {
                pastedConfigArr[++index] = (Configuration) pastedDoc.GetConfigurationByName(configName); //add config to copied config arrays
                pastedConfigComboBox.Items.Add(configName); //update the combo box (drop down menus)
            }
            pastedConfigComboBox.SelectedIndex = 0;
        }

        /// <summary>
        /// This method copies the display states from one configuration to another.
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void ButtonClicked(object sender, EventArgs e) {

            StreamWriter writer = new StreamWriter(@"C:\Users\eric.gustafson\Documents\Code\SolidWorks\bin\Release\New output text.txt");
            writer.AutoFlush = true;
            try {
                StreamWriter test = new StreamWriter(@"C:\Users\eric.gustafson\Documents\Code\SolidWorks\bin\Release\output part names.txt", true);
                test.AutoFlush = true;

                Configuration copiedConfig = (Configuration)copiedDoc.GetConfigurationByName(copiedConfigComboBox.SelectedItem.ToString());
                Configuration pastedConfig = (Configuration)copiedDoc.GetConfigurationByName(pastedConfigComboBox.SelectedItem.ToString());
                string[] copiedDisplayStateNames = (string[]) copiedConfig.GetDisplayStates();
                object oComponents = null;

                copiedDoc.ShowConfiguration2(copiedConfig.Name);
                int[] copiedConfigDisplayStateVisibilities = (int[])copiedConfig.GetDisplayStateComponentVisibility(copiedDisplayStateNames[3], false, false, out oComponents);
                object[] copiedComponents = (object[])oComponents;
                int[] copiedConfigSuppressedComponents = new int[copiedComponents.Length];
                for (int i = 0; i < copiedComponents.Length; i++) {
                    copiedConfigSuppressedComponents[i] = ((Component2) copiedComponents[i]).GetSuppression2();
                    test.WriteLine($"{DateTime.Now}: {copiedConfigSuppressedComponents[i]}");
                }


                //Get components of the pasted config
                pastedDoc.ShowConfiguration2(pastedConfig.Name);
                pastedConfig.CreateDisplayState($"{copiedDisplayStateNames[3]} copy"); //create a new display state in the pasted config
                pastedConfig.ApplyDisplayState($"{copiedDisplayStateNames[3]} copy");

                //Now that the pastedConfig is open, we can get the components of the pasted config and put them into a dictionary by name
                int[] pastedConfigDisplayStateVisibilities = (int[])pastedConfig.GetDisplayStateComponentVisibility(((string[])(pastedConfig.GetDisplayStates()))[0], false, false, out oComponents);
                Dictionary<string, object> pastedConfigDictionaryOfComponents = ((object[])oComponents).ToDictionary(key => ((Component2)key).Name2, value => value);


                try {
                    //else do nothing
                    for (int i = 0; i < copiedComponents.Length; i++) {
                        test.WriteLine(DateTime.Now + ": " + i.ToString() + " ---> " + ((Component2)copiedComponents[i]).Name2);

                        //if the component is hidden
                        if (pastedConfigDictionaryOfComponents.ContainsKey(((Component2)copiedComponents[i]).Name2)) {
                            ((Component2)pastedConfigDictionaryOfComponents[((Component2)copiedComponents[i]).Name2]).Visible = copiedConfigDisplayStateVisibilities[i];
                            ((Component2) pastedConfigDictionaryOfComponents[((Component2) copiedComponents[i]).Name2])
                                .SetSuppression2(copiedConfigSuppressedComponents[i]);
                        } //else do nothing

                        //If the component is also suppressed
                        /*if () {

                        }*/
                    }
                } catch (Exception exe) {
                    test.WriteLine($"{DateTime.Now}: {exe.ToString()}");
                }
                

                

               

                //else do nothing
                /*

                test.Close();
                copiedDoc.ShowConfiguration2(copiedConfig.Name);
                /*
                foreach (string displayStateName in copiedDisplayStateNames) {
                    pastedDoc.ShowConfiguration2(pastedConfig.Name);
                    pastedConfig.CreateDisplayState($"{displayStateName} copy"); //create a new display state in the pasted config
                    pastedConfig.ApplyDisplayState($"{displayStateName} copy");

                    componentDisplayStateVisibilities = (int[])copiedConfig.GetDisplayStateComponentVisibility(displayStateName, false, false, out oComponents);
                    for (int i = 0; i < copiedComponents.Length; i++) {
                        if (pastedConfigDictionaryOfComponents.ContainsKey(((Component2)copiedComponents[i]).Name)) {
                            ((Component2)pastedConfigDictionaryOfComponents[((Component2)copiedComponents[i]).Name]).Visible = componentDisplayStateVisibilities[i];
                        } //else do nothing
                    }
                    copiedDoc.ShowConfiguration2(copiedConfig.Name);
                }*/
            } catch (Exception ex) {
                writer.WriteLine($"{DateTime.Now}: {ex.ToString()}");
            } finally {
                writer.Close();
            }
        }
    }
}