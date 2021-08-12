/*
 * 
 *
 * @author        Eric Gustafson
 * @date_created  June 20, 2021, 04:37:42 PM
 * @version       1.0
 */

using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Media;
using System.Windows.Forms;
using System.Security;

namespace Gustafson.SolidWorks.TaskpaneAddIns {

    /**
     * The dialog box that pops-up
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




        /// <summary>
        /// Initializes the Win Form GUI
        /// </summary>
        public CopyDisplayStatesForm() {
            copiedFilePath = null;
            pastedFilePath = null;
            InitializeComponent();

            //Centering the Form in the middle of the screen
            this.Location = new System.Drawing.Point((Screen.FromControl(this).Bounds.Width - this.Width) / 2,
                (Screen.FromControl(this).Bounds.Height / 7));

            //Check to see if there is an active document in SolidWorks. If so, update the output labels (labels that display which files to look at configs from) to
            //display the file path of the currently loaded document. if there is no current document or the user browses for a new document (file) then the
            //button_pressed methods in this class that take care of that will traverse the file and find the configurations.
            if (TaskpaneExtensionsAddIn.mySolidWorks.ActiveDoc != null) {
                copiedDoc = (ModelDoc2)TaskpaneExtensionsAddIn.mySolidWorks.ActiveDoc;
                pastedDoc = (ModelDoc2)TaskpaneExtensionsAddIn.mySolidWorks.ActiveDoc;
                copiedFilePath = pastedFilePath = copiedDoc.GetPathName(); //Set the file paths of the copied and pasted documents to the active document
                nameOfCopiedDocumentLabel.Text = copiedFilePath; //Set the text of the output labels
                nameOfPastedDocumentLabel.Text = pastedFilePath;
                nameOfCopiedDocumentLabel.Visible = nameOfPastedDocumentLabel.Visible = true; //Set the labels' visibility to true

                #region Traverse document for configuration names
                string[] listOfConfigNames = (string[])copiedDoc.GetConfigurationNames(); //Get string array where each element is a configuration name
                copiedConfigArr = new Configuration[listOfConfigNames.Length]; //instantiate private class member copiedConfigArr to hold the configurations of
                pastedConfigArr = new Configuration[listOfConfigNames.Length];
                int index = -1;
                foreach (string configName in listOfConfigNames) {
                    copiedConfigArr[++index] = (Configuration)copiedDoc.GetConfigurationByName(configName); //add config to both config arrays
                    pastedConfigArr[index] = copiedConfigArr[index];
                    copiedConfigComboBox.Items.Add(configName); //update the combo boxes (drop down menus)
                    pastedConfigComboBox.Items.Add(configName);
                }
                #endregion

                //Default selections of combo boxes
                copiedConfigComboBox.SelectedIndex = 0;
                pastedConfigComboBox.SelectedIndex = 0;
            } //else keep the output labels not visible. Make the user browse for files.

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
            this.fromConfigLabel.Size = new System.Drawing.Size(551, 33);
            this.fromConfigLabel.TabIndex = 0;
            this.fromConfigLabel.Text = "Part or Assembly file to choose a configuration to copy display states FROM: ";
            this.fromConfigLabel.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // copiedConfigComboBox
            // 
            this.copiedConfigComboBox.FormattingEnabled = true;
            this.copiedConfigComboBox.Location = new System.Drawing.Point(129, 74);
            this.copiedConfigComboBox.Name = "copiedConfigComboBox";
            this.copiedConfigComboBox.Size = new System.Drawing.Size(458, 20);
            this.copiedConfigComboBox.TabIndex = 1;
            // 
            // copyButton
            // 
            this.copyButton.Font = new System.Drawing.Font("Century Gothic", 7.8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.copyButton.Location = new System.Drawing.Point(483, 257);
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
            this.toConfigLabel.Size = new System.Drawing.Size(523, 26);
            this.toConfigLabel.TabIndex = 3;
            this.toConfigLabel.Text = "Part or Assembly file to choose a configuration to copy display states TO: \r\n";
            this.toConfigLabel.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // pastedConfigComboBox
            // 
            this.pastedConfigComboBox.FormattingEnabled = true;
            this.pastedConfigComboBox.Location = new System.Drawing.Point(129, 211);
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
            this.browseForPastedConfigDocumentButton.Location = new System.Drawing.Point(569, 152);
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
            this.browseForCopiedConfigDocumentButton.Location = new System.Drawing.Point(569, 15);
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
            this.nameOfCopiedDocumentLabel.Location = new System.Drawing.Point(12, 38);
            this.nameOfCopiedDocumentLabel.Name = "nameOfCopiedDocumentLabel";
            this.nameOfCopiedDocumentLabel.Size = new System.Drawing.Size(535, 36);
            this.nameOfCopiedDocumentLabel.TabIndex = 11;
            this.nameOfCopiedDocumentLabel.Text = "null_copied";
            this.nameOfCopiedDocumentLabel.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            this.nameOfCopiedDocumentLabel.Visible = false;
            // 
            // nameOfPastedDocumentLabel
            // 
            this.nameOfPastedDocumentLabel.Font = new System.Drawing.Font("Century Gothic", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.nameOfPastedDocumentLabel.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(192)))), ((int)(((byte)(0)))));
            this.nameOfPastedDocumentLabel.Location = new System.Drawing.Point(12, 172);
            this.nameOfPastedDocumentLabel.Name = "nameOfPastedDocumentLabel";
            this.nameOfPastedDocumentLabel.Size = new System.Drawing.Size(551, 36);
            this.nameOfPastedDocumentLabel.TabIndex = 12;
            this.nameOfPastedDocumentLabel.Text = "null_pasted";
            this.nameOfPastedDocumentLabel.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            this.nameOfPastedDocumentLabel.Visible = false;
            // 
            // CopyDisplayStatesForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(724, 304);
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

        #region Open Document Methods
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
                openFileDialog.Filter = "SolidWorks Parts (SolidWorks Assemblies (*.SLDASM)|*.SLDASM|*.SLDPRT)|*.SLDPRT";
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

            string[] listOfConfigNames = (string[])copiedDoc.GetConfigurationNames(); //Get string array where each element is a configuration name
            copiedConfigArr = new Configuration[listOfConfigNames.Length]; //instantiate private class member copiedConfigArr to hold the configurations of
            int index = -1;
            foreach (string configName in listOfConfigNames) {//for each configuration in this document, get the names of the configurations and save them in the combobox
                copiedConfigArr[++index] = (Configuration)copiedDoc.GetConfigurationByName(configName); //add config to copied config arrays
                copiedConfigComboBox.Items.Add(configName); //update the combo box (drop down menus)
            }
            copiedConfigComboBox.SelectedIndex = 0; //set selection to the first element
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
                } else { //The other file is a .SLDASM
                    openFileDialog.Filter = "SolidWorks Assemblies (*.SLDASM)|*.SLDASM";
                }
            } else { //else nothing has been selected yet.
                openFileDialog.Filter = "SolidWorks Parts (SolidWorks Assemblies (*.SLDASM)|*.SLDASM|*.SLDPRT)|*.SLDPRT";
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

            //Copy the configurations of the new file into the combo boxes
            DocumentSpecification swDocSpecification = (DocumentSpecification)TaskpaneExtensionsAddIn.mySolidWorks.GetOpenDocSpec(pastedFilePath);
            pastedDoc = TaskpaneExtensionsAddIn.mySolidWorks.OpenDoc7(swDocSpecification); //

            string[] listOfConfigNames = (string[])pastedDoc.GetConfigurationNames(); //String array where each element is a configuration name
            pastedConfigArr = new Configuration[listOfConfigNames.Length]; //instantiate private class member copiedConfigArr to hold the configurations of
            int index = -1;
            foreach (string configName in listOfConfigNames) { //for each configuration in this document, get the names of the configurations and save them in the combobox
                pastedConfigArr[++index] = (Configuration)pastedDoc.GetConfigurationByName(configName); //add config to copied config arrays
                pastedConfigComboBox.Items.Add(configName); //update the combo box (drop down menus)
            }
            pastedConfigComboBox.SelectedIndex = 0; //set selection to the first element
        }
        #endregion

        #region Copy Display State Methods
        /// <summary>
        /// This method is invoked when the "Copy Display States" button is pressed. It simply looks to see if it is copying
        /// display states from two part configurations or assembly configurations, and then calls the appropriate method to do that.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void ButtonClicked(object sender, EventArgs e) {
            if (copiedFilePath.Substring(copiedFilePath.Length - 7).ToUpper().Equals(".SLDPRT")) {
                throw new NotImplementedException("Copying display states from parts has not yet been implemented");
            } else { //else it is an assembly
               // StreamWriter debugger =new StreamWriter(@"C:\Users\eric.gustafson\Documents\Code\SolidWorks\bin\Release\New output text.txt");
               // debugger.AutoFlush = true;
                try {
                    CopyDisplayStatesFromAssemblies();
                } catch (Exception ex) {
                //    debugger.WriteLine($"{DateTime.Now}: {ex.ToString()}");
                } finally {
                //    debugger.Close();
                }
            }
        }

        internal struct DisplayStateSuppressionStateValues {
            int DisplayStateValue { get; set; }
            int SuppressionStateValue { get; set; }

            public DisplayStateSuppressionStateValues(int display, int suppression) {
                DisplayStateValue = display;
                SuppressionStateValue = suppression;
            }
        }


        /// <summary>
        /// 
        /// </summary>
        private void CopyDisplayStatesFromAssemblies() {
            //get the configurations we're copying display states from and too
            Configuration copiedConfig = (Configuration)copiedDoc.GetConfigurationByName(copiedConfigComboBox.SelectedItem.ToString());
            Configuration pastedConfig = (Configuration)pastedDoc.GetConfigurationByName(pastedConfigComboBox.SelectedItem.ToString());
            string[] copiedDisplayStateNames = (string[])copiedConfig.GetDisplayStates(); //obtain the names of the display states we're going to copy

            #region Deal with suppressed component replacements
            copiedDoc.ShowConfiguration2(copiedConfig.Name); //show copied config
            copiedConfig.ApplyDisplayState(copiedDisplayStateNames[0]); //show any display state (array is guaranteed to have index 0)
            Dictionary<string, int> dictOfSuppressionStates = new Dictionary<string, int>();
            GetSuppressionStates(copiedConfig.GetRootComponent3(true).GetChildren());

            //Search through the pasted configuration for a list of all parts
            pastedDoc.ShowConfiguration2(pastedConfig.Name);
            ReplaceSuppressedComponentsForm replaceForm = new ReplaceSuppressedComponentsForm(copiedConfig, pastedConfig, copiedDoc.GetTitle());
            GetComponents(pastedConfig.GetRootComponent3(true).GetChildren());
            foreach (string name in dictOfSuppressionStates.Keys) {
                replaceForm.AddItemToCombobox(name, "Suppressed Components");
            }
            Dictionary<string, string> suppressedComponentReplacement = new Dictionary<string, string>(); //<k, v> = <name of component in pastedDoc, name of component to apply visibility of>
            replaceForm.ShowDialog();
            replaceForm.GetTableData(ref suppressedComponentReplacement);
            if (replaceForm.UserQuit) return;


            Component2 tempCompSup;
            void GetSuppressionStates(object[] children) {
                foreach (object component in children) {
                    tempCompSup = (Component2)component;
                    if (tempCompSup.GetSuppression() == (int) swComponentSuppressionState_e.swComponentSuppressed) {  //if the component is suppressed then add it
                        dictOfSuppressionStates.Add(tempCompSup.Name2, tempCompSup.GetSuppression());
                    } else if (IsAssemblyFile(tempCompSup)) { //if the component is NOT suppressed and is an assembly file, then go deeper.
                        GetSuppressionStates(tempCompSup.GetChildren());
                    }
                }
            }

            Component2 tempCompSup2;
            void GetComponents(object[] children) {
                foreach (object component in children) {
                    tempCompSup2 = (Component2) component;
                    if (tempCompSup2.GetSuppression() == (int) swComponentSuppressionState_e.swComponentSuppressed) {
                        if (dictOfSuppressionStates.ContainsKey(tempCompSup2.Name2)) {
                            dictOfSuppressionStates.Remove(tempCompSup2.Name2);
                        }
                    } else if (IsAssemblyFile(tempCompSup2)) { //if not suppressed and an assembly file
                        replaceForm.AddItemToCombobox(tempCompSup2.Name2, "Replacement Components");
                        GetComponents(tempCompSup2.GetChildren());
                    } else {
                        replaceForm.AddItemToCombobox(tempCompSup2.Name2, "Replacement Components");
                    }
                }
            }
            #endregion









            DateTime now = DateTime.Now; //Get current time for timer purposes

            //Get Selection Tools from the pasted Document
            SelectionMgr swSelectionMgr = (SelectionMgr)pastedDoc.SelectionManager;
            SelectData swSelectData = (SelectData)swSelectionMgr.CreateSelectData();

            Component2 tempComp;
            Dictionary<string, int> dictOfCopiedComps = new Dictionary<string, int>(); //(key, value) = <string, int> = <component name, component visibility>
            HashSet<Component2> componentsToShow = new HashSet<Component2>(); //to keep track of components that need to be made visible
            foreach (string displayStateName in copiedDisplayStateNames) { //for each display state
                copiedDoc.ShowConfiguration2(copiedConfig.Name); //switch to the configuration that we're copying display states from.
                copiedConfig.ApplyDisplayState(displayStateName); //apply new display state to copy

                //Get component Visibilities of copying configuration
                object[] temp = copiedConfig.GetRootComponent3(true).GetChildren();
                AddComponents(temp); //recursively add to dictOfComponents

                
                pastedDoc.ShowConfiguration2(pastedConfig.Name); //show pasted configuration so we may create a new display state for it
                pastedConfig.CreateDisplayState($"{displayStateName} copy"); //create a new display state in the pasted config
                pastedConfig.ApplyDisplayState($"{displayStateName} copy"); //apply this new display state so we may apply the appropriate visibilities to its components
                //If we do NOT switch to the pasted configuration and stay on the copied configuration, when we go to get the components in
                //the pasted configuration, even though we've listed the name of a pastedConfig display state, we will still be getting
                //components and their visibilities from the active configuration (which would be the copied one).
                 


                //recursively select components to be hidden by passing the root component of the pasted configuration. Any component that needs to be visible
                //instead of hidden is added to the componentsToShow HashSet.
                TraverseAssemblyAndSelectComponentsToBeHidden(pastedConfig.GetRootComponent3(true).GetChildren());

                pastedDoc.HideComponent2(); //Hide all of the selected components in this display state
                pastedDoc.ClearSelection2(true); //Deselect all components

                //Now we must show all the components with a visibility state of 1
                if (componentsToShow.Count != 0) {
                    foreach (Component2 component in componentsToShow) { //select all components in the HashSet
                        component.Select4(true, swSelectData, false);
                    }

                    pastedDoc.ShowComponent2(); //Show all components that are selected
                    pastedDoc.ClearSelection2(true); //Deselect all components
                }

                //clear data structures
                dictOfCopiedComps.Clear();
                componentsToShow.Clear();
            } //end foreach loop
            
            try { //play a nice little jingle when it is done in case anybody has this running in the background.
                SoundPlayer completedSound = new SoundPlayer(@"C:\Windows\Media\Windows Proximity Notification.wav");
                completedSound.Play();
            } catch (Exception ex) {
                Debug.Print("Cannot locate sound file");
            }
        
            MessageBox.Show($"\t\t\t\tSuccess!\nOperation took {(DateTime.Now - now).TotalSeconds} seconds ({(DateTime.Now - now).Minutes} minutes and {(DateTime.Now - now).TotalSeconds % 60} seconds)");
            this.Close(); //close this WinForm GUI.






            void AddComponents(object[] children) {
                foreach (object component in children) { //traverse children
                    tempComp = (Component2)component;
                    if (IsAssemblyFile(tempComp) && tempComp.Visible == (int)swComponentVisibilityState_e.swComponentVisible) {  //if children[i] is an assembly file AND it is visible,
                        dictOfCopiedComps.Add(tempComp.Name2, tempComp.Visible);
                        AddComponents(tempComp.GetChildren()); //then we need to go deeper :)
                    } else {
                        dictOfCopiedComps.Add(tempComp.Name2, tempComp.Visible);
                    }
                }
            }

            int visibility;
            void TraverseAssemblyAndSelectComponentsToBeHidden(object[] children) {
                foreach (object component in children) { //traverse children

                    if (IsAssemblyFile((Component2)component)) {                     //if children[i] is an assembly file AND it is visible,
                        TraverseAssemblyAndSelectComponentsToBeHidden(((Component2)component).GetChildren()); //then we need to go deeper :)
                    }
                    //If the copiedConfig contains the current, pastedConfig's part
                    if (dictOfCopiedComps.ContainsKey(((Component2)component).Name2)) {
                        visibility = suppressedComponentReplacement.ContainsKey(((Component2)component).Name2) && dictOfCopiedComps.ContainsKey(suppressedComponentReplacement[((Component2)component).Name2]) ?
                            dictOfCopiedComps[suppressedComponentReplacement[((Component2)component).Name2]] : dictOfCopiedComps[((Component2)component).Name2];
                        if (visibility == (int) swComponentVisibilityState_e.swComponentHidden) { //if the part in the copied config is hidden, then select it to be hidden later
                            ((Component2)component).Select4(true, swSelectData, false);
                        } else { //if the part is showing, mark it to be shown later
                            componentsToShow.Add((Component2)component);
                        }
                    }
                }
            }
        }

        /*
         * We have to check if the components children array is greater than zero because, for some odd reason,
         * sometimes part files DO have non-null GetChildren arrays of length zero???? Not sure why, but that's why the
         * tempComp.GetChildren().Length > 0       is in the if-statement
         *
         * We have to check if the children are null first because sometimes the part files do return a null GetChildren(),
         * and a NullReferenceException will be thrown if we try and do component.GetChildren().Length, so that first expression
         * avoids that.
         */
        public static bool IsAssemblyFile(Component2 component) {
            return component.GetChildren() != null && component.GetChildren().Length > 0;
        }
        #endregion
    }




    

    class ReplaceSuppressedComponentsForm : Form {
        private Configuration copiedConfig;
        private Configuration pastedConfig;
        private Dictionary<string, int> rowItems; //<key, value> = <string, int> = <suppressed component name, row index>
        public bool UserQuit { get; set; } = false;

        public ReplaceSuppressedComponentsForm(Configuration copiedConfig, Configuration pastedConfig, string documentName) {
            this.copiedConfig = copiedConfig;
            this.pastedConfig = pastedConfig;
            this.rowItems = new Dictionary<string, int>();
            InitializeComponent();
            InitializeTable(documentName);
            suppressionTable.AutoResizeColumns();
            //Centering the Form in the middle of the screen
            this.Location = new System.Drawing.Point((Screen.FromControl(this).Bounds.Width - this.Width) / 2,
                (Screen.FromControl(this).Bounds.Height / 7));
        }

        private void InitializeTable(string documentName) {
            this.suppressionTable.ColumnCount = 2;
            this.suppressionTable.Columns[0].Name = $"Suppressed Component in {documentName}";
            this.suppressionTable.Columns[1].Name = "Component to replace the suppressed component with";

            this.suppressionTable.Columns[0].Width = this.suppressionTable.Columns[1].Width = suppressionTable.Width / 2 - 10;
            this.suppressionTable.Columns[0].Resizable = this.suppressionTable.Columns[1].Resizable = DataGridViewTriState.False;
            this.suppressionTable.MultiSelect = false;
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
            this.suppressedComponentComboBox.SelectedIndexChanged += new System.EventHandler(this.SuppressedComponentComboBoxSelectedIndexChanged);
            this.suppressedComponentComboBox.AutoCompleteMode = AutoCompleteMode.SuggestAppend;
            this.suppressedComponentComboBox.AutoCompleteSource = AutoCompleteSource.ListItems;
            // 
            // ReplacementComponentComboBox
            // 
            this.ReplacementComponentComboBox.FormattingEnabled = true;
            this.ReplacementComponentComboBox.Location = new System.Drawing.Point(537, 56);
            this.ReplacementComponentComboBox.Name = "ReplacementComponentComboBox";
            this.ReplacementComponentComboBox.Size = new System.Drawing.Size(398, 21);
            this.ReplacementComponentComboBox.TabIndex = 1;
            this.ReplacementComponentComboBox.AutoCompleteMode = AutoCompleteMode.SuggestAppend;
            this.ReplacementComponentComboBox.AutoCompleteSource = AutoCompleteSource.ListItems;
            // 
            // suppressedComponentLabel
            // 
            this.suppressedComponentLabel.AutoSize = true;
            this.suppressedComponentLabel.Location = new System.Drawing.Point(24, 38);
            this.suppressedComponentLabel.Name = "suppressedComponentLabel";
            this.suppressedComponentLabel.Size = new System.Drawing.Size(200, 13);
            this.suppressedComponentLabel.TabIndex = 2;
            this.suppressedComponentLabel.Text = $"Suppressed component in configuration: {copiedConfig.Name}";
            // 
            // replacementComponentLabel
            // 
            this.replacementComponentLabel.AutoSize = true;
            this.replacementComponentLabel.Location = new System.Drawing.Point(534, 38);
            this.replacementComponentLabel.Name = "replacementComponentLabel";
            this.replacementComponentLabel.Size = new System.Drawing.Size(92, 13);
            this.replacementComponentLabel.TabIndex = 3;
            this.replacementComponentLabel.Text = "Please select component to replace";
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
            this.runButton.Click += new System.EventHandler(this.RunButtonClicked);
            // 
            // cancelButton
            // 
            this.cancelButton.Location = new System.Drawing.Point(701, 643);
            this.cancelButton.Name = "cancelButton";
            this.cancelButton.Size = new System.Drawing.Size(75, 23);
            this.cancelButton.TabIndex = 7;
            this.cancelButton.Text = "Cancel";
            this.cancelButton.UseVisualStyleBackColor = true;
            this.cancelButton.Click += new System.EventHandler(this.CancelButtonClicked);
            // 
            // addReplacementButton
            // 
            this.addReplacementButton.Location = new System.Drawing.Point(537, 101);
            this.addReplacementButton.Name = "addReplacementButton";
            this.addReplacementButton.Size = new System.Drawing.Size(164, 23);
            this.addReplacementButton.TabIndex = 9;
            this.addReplacementButton.Text = "Add Replacement";
            this.addReplacementButton.UseVisualStyleBackColor = true;
            this.addReplacementButton.Click += new System.EventHandler(this.AddReplacementButtonClicked);
            // 
            // deleteReplacementRow
            // 
            this.deleteReplacementRow.Location = new System.Drawing.Point(27, 654);
            this.deleteReplacementRow.Name = "deleteReplacementRow";
            this.deleteReplacementRow.Size = new System.Drawing.Size(118, 23);
            this.deleteReplacementRow.TabIndex = 10;
            this.deleteReplacementRow.Text = "Delete Selected Rows";
            this.deleteReplacementRow.UseVisualStyleBackColor = true;
            this.deleteReplacementRow.Click += new System.EventHandler(this.DeleteReplacementRowButtonClicked);
            // 
            // deleteReplacementLabel
            // 
            this.deleteReplacementLabel.AutoSize = true;
            this.deleteReplacementLabel.Location = new System.Drawing.Point(24, 638);
            this.deleteReplacementLabel.Name = "deleteReplacementLabel";
            this.deleteReplacementLabel.Size = new System.Drawing.Size(284, 13);
            this.deleteReplacementLabel.TabIndex = 12;
            this.deleteReplacementLabel.Text = "If rows in the table are selected, this button will delete them";
            // 
            // ReplaceSuppressedComponentsForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(977, 700);
            this.Controls.Add(this.deleteReplacementLabel);
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
        private System.Windows.Forms.Label deleteReplacementLabel;

        #endregion

        public void AddItemToCombobox(string itemToAdd, string comboBoxToAddTo) {
            switch (comboBoxToAddTo.ToUpper()) {
                case "SUPPRESSED COMPONENTS":
                    suppressedComponentComboBox.Items.Add(itemToAdd);
                    break;
                case "REPLACEMENT COMPONENTS":
                    ReplacementComponentComboBox.Items.Add(itemToAdd);
                    break;
                default:
                    throw new ArgumentException("Combo box string must be either \"Suppressed Components\" or \"Replacement Components\"");
            }
        }

        public void GetTableData(ref Dictionary<string, string> dictToAddTo) {
            for (int i = 0; i < suppressionTable.RowCount - 1; i++) {
                dictToAddTo.Add(this.suppressionTable.Rows[i].Cells[0].Value.ToString(), this.suppressionTable.Rows[i].Cells[1].Value.ToString());
            }
        }

        private void AddReplacementButtonClicked(object sender, EventArgs e) {
            try {
                if (rowItems.ContainsKey(suppressedComponentComboBox.SelectedItem.ToString())) {

                    this.suppressionTable.Rows.RemoveAt(rowItems[suppressedComponentComboBox.SelectedItem.ToString()]);
                    rowItems.Remove(suppressedComponentComboBox.SelectedItem.ToString());
                }


                this.suppressionTable.Rows.Add(new object[] {
                    suppressedComponentComboBox.SelectedItem.ToString(),
                    ReplacementComponentComboBox.SelectedItem.ToString()
                });
                suppressionTable.Rows[suppressionTable.RowCount - 2].HeaderCell.Value = $"{suppressionTable.RowCount - 2}";
                rowItems.Add(suppressedComponentComboBox.SelectedItem.ToString(), suppressionTable.RowCount - 2);
            } catch (NullReferenceException ex) {
                MessageBox.Show("Please only select an item in the drop-down menu! Do NOT write in your own values!",
                    "Problematic Operation: Write-In Value", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void DeleteReplacementRowButtonClicked(object sender, EventArgs e) {
            if (this.suppressionTable.SelectedRows.Count > 0 && this.suppressionTable.SelectedRows[0].Index != this.suppressionTable.Rows.Count - 1) {
                rowItems.Remove(this.suppressionTable.SelectedRows[0].Cells[0].Value.ToString());
                this.suppressionTable.Rows.RemoveAt(this.suppressionTable.SelectedRows[0].Index);
            }
        }

        private void CancelButtonClicked(object sender, EventArgs e) {
            UserQuit = true;
            this.Close();
        }

        private void RunButtonClicked(object sender, EventArgs e) {
            this.Close();
        }

        private void SuppressedComponentComboBoxSelectedIndexChanged(object sender, EventArgs e) {
            this.replacementComponentLabel.Text = $"Replace {suppressedComponentComboBox.SelectedText.ToUpper()} in {copiedConfig.Name} with: ";
        }
    }

    
}