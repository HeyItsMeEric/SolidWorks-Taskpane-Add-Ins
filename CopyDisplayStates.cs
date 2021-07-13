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
using System.Linq;
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
            this.copiedConfigComboBox.Size = new System.Drawing.Size(458, 23);
            this.copiedConfigComboBox.TabIndex = 1;
            // 
            // copyButton
            // 
            this.copyButton.Font = new System.Drawing.Font("Century Gothic", 7.8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.copyButton.Location = new System.Drawing.Point(244, 254);
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
            this.pastedConfigComboBox.Size = new System.Drawing.Size(458, 23);
            this.pastedConfigComboBox.TabIndex = 4;
            // 
            // configurationTextFromLabel
            // 
            this.configurationTextFromLabel.AutoSize = true;
            this.configurationTextFromLabel.Font = new System.Drawing.Font("Century Gothic", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.configurationTextFromLabel.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.configurationTextFromLabel.Location = new System.Drawing.Point(12, 74);
            this.configurationTextFromLabel.Name = "configurationTextFromLabel";
            this.configurationTextFromLabel.Size = new System.Drawing.Size(111, 19);
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
            this.configurationTextPastedLabel.Size = new System.Drawing.Size(111, 19);
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
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 15F);
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

            string[] listOfConfigNames = (string[]) copiedDoc.GetConfigurationNames(); //Get string array where each element is a configuration name
            copiedConfigArr = new Configuration[listOfConfigNames.Length]; //instantiate private class member copiedConfigArr to hold the configurations of
            int index = -1;
            foreach (string configName in listOfConfigNames) {//for each configuration in this document, get the names of the configurations and save them in the combobox
                copiedConfigArr[++index] = (Configuration) copiedDoc.GetConfigurationByName(configName); //add config to copied config arrays
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
                }
                else { //The other file is a .SLDASM
                    openFileDialog.Filter = "SolidWorks Assemblies (*.SLDASM)|*.SLDASM";
                }
            }
            else { //else nothing has been selected yet.
                openFileDialog.Filter = "SolidWorks Parts (SolidWorks Assemblies (*.SLDASM)|*.SLDASM|*.SLDPRT)|*.SLDPRT";
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

            string[] listOfConfigNames = (string[]) pastedDoc.GetConfigurationNames(); //String array where each element is a configuration name
            pastedConfigArr = new Configuration[listOfConfigNames.Length]; //instantiate private class member copiedConfigArr to hold the configurations of
            int index = -1;
            foreach (string configName in listOfConfigNames) { //for each configuration in this document, get the names of the configurations and save them in the combobox
                pastedConfigArr[++index] = (Configuration) pastedDoc.GetConfigurationByName(configName); //add config to copied config arrays
                pastedConfigComboBox.Items.Add(configName); //update the combo box (drop down menus)
            }
            pastedConfigComboBox.SelectedIndex = 0; //set selection to the first element
        }

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
                StreamWriter debugger =
                    new StreamWriter(
                        @"C:\Users\eric.gustafson\Documents\Code\SolidWorks\bin\Release\New output text.txt");
                debugger.AutoFlush = true;
                try {
                    copyDisplayStatesFromAssemblies(ref debugger);
                } catch (Exception ex) {
                    debugger.WriteLine($"{DateTime.Now}: {ex.ToString()}");
                } finally {
                    debugger.Close();
                }
            }
        }

        /// <summary>
        /// This method is called when the "Copy Display States" button is pressed and the two configurations in question are from two
        /// assembly documents.
        /// </summary>
        internal void copyDisplayStatesFromAssemblies(ref StreamWriter debugger) {
            //Obtain the configurations to copy and paste the display states to and from, respectively, by looking at the combo boxes in the WinForms GUI.
            //debugger.WriteLine("The length is: \t" + copiedDoc.GetConfigurationByName(copiedConfigComboBox.SelectedItem.ToString().ToString()));
            Configuration copiedConfig = (Configuration)copiedDoc.GetConfigurationByName(copiedConfigComboBox.SelectedItem.ToString());
            Configuration pastedConfig = (Configuration)pastedDoc.GetConfigurationByName(pastedConfigComboBox.SelectedItem.ToString());
            
            string[] copiedDisplayStateNames = (string[])copiedConfig.GetDisplayStates(); //obtain the names of the display states we're going to copy
            object oComponents = null; //this variable is simply a temporary one that is passed by reference to the'
                                       //Configuration.GetDisplayStateComponentVisibility to obtain the components of that configurations

            copiedDoc.ShowConfiguration2(copiedConfig.Name); //switch to the configuration that we're copying display states from.
            int[] copiedConfigDisplayStateVisibilities = (int[])copiedConfig.GetDisplayStateComponentVisibility(copiedDisplayStateNames[0], false, false, out oComponents); //get component visibility AND get components
            object[] copiedComponents = (object[])oComponents; //since oComponents is passed by reference in GetDisplayStateComponentVisibility, it can now be cast to an array containing our components
            int[] copiedConfigSuppressedComponents = new int[copiedComponents.Length]; //instantiate a new array to hold the suppression states for each component
            for (int i = 0; i < copiedComponents.Length; i++) { //for each component, get the suppression state
                copiedConfigSuppressedComponents[i] = ((Component2)copiedComponents[i]).GetSuppression(); //THE SOLIDWORKS 2018 ....sldworks.dll DOES NOT HAVE GetSuppression2(), BUT 2019 DOES
            }

            /*                                                      Get components of the pasted config
             *
             * First, switch to the pasted configuration. If we do NOT switch to the pasted configuration and stay on the
             * copied configuration, when we go to get the components in the pasted configuration, even though
             * we've listed the name of a pastedConfig display state, we will still be getting components and their
             * visibilities from the active configuration (which would be the copied one). In fact, if we don't
             * switch and the pastedConfig display state name we pass in is not a valid display state name for the
             * active configuration, then oComponents will be set to null.
             */
            pastedDoc.ShowConfiguration2(pastedConfig.Name); //Switch to pasted configuration

            //Now that the pastedConfig is open, we can get the components of the pasted config and put them into a dictionary by name
            pastedConfig.GetDisplayStateComponentVisibility(((string[])(pastedConfig.GetDisplayStates()))[0], false, false, out oComponents); //no need to save component visibility in a new display state,
                                                                                                                                              //The point of this line is to get the pastedConfig components
            /*
             * The key value pair of this dictionary is <Key, Value> = <string, object> = <"SolidWorks Component Name", "SolidWorks COM object as object class">
             *
             * The reason I'm using a Dictionary here as opposed to a regular array is because this method is supposed to work for different assembly files, and they
             * will most likely not have the same parts in within each other; we need to make sure that we only want to set the visibilities and suppression states of
             * components/parts that are common between the configurations, so, we must first check whether a component/part exists in the pasted configuration before
             * changing its visibility or suppression state. The easiest way to do this is with a dictionary (same data structure as a "HashMap/HashTable" in Java),
             * because a Dictionary has O(1) (fastest) lookup time. Also, it is convenient to store the Components (which at this point are COM Objects because we
             * have not yet casted them to a Component2 type yet) in a key-value pair with the key being the component's name as a string and the value being the
             * uncasted COM object.
             *
             * We can use the ToDictionary() method from the System.Linq library to convert the casted array from oComponents into a Dictionary.
             */
            object[] comps = (object[]) oComponents;
            Dictionary<string, object> pastedConfigDictionaryOfComponents = new Dictionary<string, object>(comps.Length);
            foreach (object o in comps) {
                debugger.WriteLine($"{DateTime.Now}: {((Component2) o).Name2}");
                try {
                    pastedConfigDictionaryOfComponents.Add(((Component2) o).Name2, o);
                }
                catch (Exception ex) {
                    debugger.WriteLine($"{DateTime.Now}: {ex.ToString()}");
                }
            }
            //Dictionary<string, object> pastedConfigDictionaryOfComponents = comps.ToDictionary(key => ((Component2)key).Name2, value => value);
            
            
            
            //Dictionary<string, object> pastedConfigDictionaryOfComponents = ((object[])oComponents).ToDictionary(key => ((Component2)key).Name2, value => value);

            //Here we iterate through the copied Components and apply their suppression states to those common components/parts in the pasted configuration
            //Suppressed components/parts are applied to a configuration, not specific display states; in other words, if you suppress a component in one
            //display state, all other display states will make that component suppressed as well (this is a SolidWorks feature, not something I coded).
            //Because of this, we only want to check and apply suppression states once, not for every display state.
           /* for (int i = 0; i < copiedComponents.Length; i++) {

                //Check to see if the copied component exists as a part in the pasted configuration; if so, then set the suppression state to what is was in the copied config.
                if (pastedConfigDictionaryOfComponents.ContainsKey(((Component2)copiedComponents[i]).Name2)) {
                    ((Component2)pastedConfigDictionaryOfComponents[((Component2)copiedComponents[i]).Name2]).SetSuppression2(copiedConfigSuppressedComponents[i]);
                } //else the copied component does not exist in the pasted assembly document; do nothing
            }*/

            //This big bad foreach loop is what creates the new display states in the pasted configurations and sets the visibilities of their components
            foreach (string displayStateName in copiedDisplayStateNames) {
                copiedDoc.ShowConfiguration2(copiedConfig.Name); //open copied configuration so we may get the display state visibilities of that configuration.
                copiedConfigDisplayStateVisibilities = (int[])copiedConfig.GetDisplayStateComponentVisibility(displayStateName, false, false, out oComponents); //no need to save oComponents this time

                pastedDoc.ShowConfiguration2(pastedConfig.Name); //show pasted configuration so we may create a new display state for it
                pastedConfig.CreateDisplayState($"{displayStateName} copy"); //create a new display state in the pasted config
                pastedConfig.ApplyDisplayState($"{displayStateName} copy"); //apply this new display state so we may apply the appropriate visibilities to its components

                #region please work on this
                pastedConfig.GetDisplayStateComponentVisibility(((string[])(pastedConfig.GetDisplayStates()))[0], false, false, out oComponents); //just here for components, not component visibilities
                //pastedConfigDictionaryOfComponents = ((object[])oComponents).ToDictionary(key => ((Component2)key).Name2, value => value); //create new Dictionary of all parts in the pasted configuration
                comps = (object[])oComponents;
                pastedConfigDictionaryOfComponents.Clear();
                foreach (object o in comps) {
                    //debugger.WriteLine($"{DateTime.Now}: {((Component2)o).Name2}");
                    try {
                        pastedConfigDictionaryOfComponents.Add(((Component2)o).Name2, o);
                    } catch (Exception ex) {
                        debugger.WriteLine($"{DateTime.Now}: {ex.ToString()}");
                    }
                }
                #endregion 
                for (int i = 0; i < copiedComponents.Length; i++) { //for each component
                    if (pastedConfigDictionaryOfComponents.ContainsKey(((Component2)copiedComponents[i]).Name2)) { //check if the component from the copied config exists in the pasted config
                        ((Component2)pastedConfigDictionaryOfComponents[((Component2)copiedComponents[i]).Name2]).Visible = copiedConfigDisplayStateVisibilities[i]; //apply visibilities
                    }
                }
                this.Close();
            }
        }
    }
}