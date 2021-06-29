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
using System.IO;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swpublished;

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
            this.taskpaneCopyDisplayStatesButton = new System.Windows.Forms.Button();
            this.CDSBArtist = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // taskpaneCopyDisplayStatesButton
            // 
            this.taskpaneCopyDisplayStatesButton.Location = new System.Drawing.Point(15, 10);
            this.taskpaneCopyDisplayStatesButton.Margin = new System.Windows.Forms.Padding(2);
            this.taskpaneCopyDisplayStatesButton.Name = "taskpaneCopyDisplayStatesButton";
            this.taskpaneCopyDisplayStatesButton.Size = new System.Drawing.Size(118, 46);
            this.taskpaneCopyDisplayStatesButton.TabIndex = 0;
            this.taskpaneCopyDisplayStatesButton.Text = "Copy Display States";
            this.taskpaneCopyDisplayStatesButton.UseVisualStyleBackColor = true;
            this.taskpaneCopyDisplayStatesButton.Click += new System.EventHandler(this.CopyDisplayStatesButtonClicked);
            // 
            // CDSBArtist
            // 
            this.CDSBArtist.AutoSize = true;
            this.CDSBArtist.Location = new System.Drawing.Point(3, 58);
            this.CDSBArtist.Name = "CDSBArtist";
            this.CDSBArtist.Size = new System.Drawing.Size(130, 13);
            this.CDSBArtist.TabIndex = 1;
            this.CDSBArtist.Text = "Created by Eric Gustafson";
            // 
            // TaskpaneAddInManager
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.CDSBArtist);
            this.Controls.Add(this.taskpaneCopyDisplayStatesButton);
            this.Margin = new System.Windows.Forms.Padding(2);
            this.Name = "TaskpaneAddInManager";
            this.Size = new System.Drawing.Size(333, 659);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        private System.Windows.Forms.Button taskpaneCopyDisplayStatesButton;
        private Label CDSBArtist;
        #endregion

        /// <summary>
        /// Called when the user clicks on the "Copy Display States" Button. Runs the CopyDisplayStateForm.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void CopyDisplayStatesButtonClicked(object sender, EventArgs e) {
            //Application.SetHighDpiMode(HighDpiMode.SystemAware);
            Application.EnableVisualStyles();
            //Application.SetCompatibleTextRenderingDefault(false);
            StreamWriter debugger = new StreamWriter(@"C:\Users\13016\Documents\COMP SCI - C#\SolidWorks Add-Ins\bin\Release\text3.txt");
            debugger.WriteLine($"{DateTime.Now}: {"Hello"}");
            debugger.WriteLine($"{DateTime.Now}: {"Hello"}");
            debugger.WriteLine($"{DateTime.Now}: {"Hello"}");
            debugger.WriteLine($"{DateTime.Now}: {"Hello"}");
            try {
                debugger.WriteLine($"{DateTime.Now}: {"Hello"}");
                Application.Run(new CopyDisplayStatesForm());
            } catch (Exception ex) {
                debugger.WriteLine($"{DateTime.Now}: {ex.ToString()}");
            } finally {
                debugger.Close();
            }


        }

    }


















    /**
     * I did not code this class, I merely copied it from this Youtube video: https://www.youtube.com/watch?v=7DlG6OQeJP0
     *
     * Everything else (that was not autogenerated) is my code and mine alone.
     *
     *
     * This class registers this add-in with the Windows Registry. It also contains initial code for what SolidWorks does with this add-in once it's registered.
     */
    public class TaskpaneExtensionsAddIn : ISwAddin {
        #region SolidWorks Members

        private TaskpaneView mySolidWorksTaskPane;
        public static SldWorks mySolidWorks { get; set; }

        private TaskpaneAddInManager manager;
        private readonly string PROGID = $@"{typeof(TaskpaneExtensionsAddIn).Namespace}";


        #endregion

        /// <summary>
        ///   Called when SolidWorks loads the Add-In
        /// </summary>
        /// <param name="thisSw"> Instance of the SolidWorks Application </param>
        /// <param name="cookie"></param>
        /// <returns></returns>
        public bool ConnectToSW(object thisSw, int cookie) {
            //For debugging
            /*
            StreamWriter debugger = new StreamWriter(@"C:\Users\eric.gustafson\Documents\Code\SolidWorks\bin\Release\New output text.txt");
            debugger.WriteLine($"{DateTime.Now}: 1");
            */
            mySolidWorks = (SldWorks)thisSw;
            mySolidWorks.SetAddinCallbackInfo2(0, this, cookie);

            Application.SetCompatibleTextRenderingDefault(false); //Must be set before the first IWin32Window object (WinForms object) is created, which is the next line.
            mySolidWorksTaskPane = mySolidWorks.CreateTaskpaneView2($@"{System.Reflection.Assembly.GetExecutingAssembly().Location.Replace("Gustafson.SolidWorks.TaskpaneAddIns.dll", "")}\\..\\..\\Images\\SW Custom Add-in Taskpane Icon.png", "Custom SolidWorks Add-Ins");
            manager = (TaskpaneAddInManager)mySolidWorksTaskPane.AddControl(PROGID, string.Empty);
           /* debugger.WriteLine($"{DateTime.Now}: 2");
            debugger.Close(); //for debugging*/
            return true; //success
        }

        
        public bool DisconnectFromSW() {
            mySolidWorksTaskPane.DeleteView();
            Marshal.ReleaseComObject(mySolidWorksTaskPane);

            //Not necessary, but good for precautionary reasons
            manager = null;
            mySolidWorksTaskPane = null;

            return true; //success
        }

        /*
         * The Windows operating System has a registry that stores much of the information and settings for software
         * programs, hardware devices, user preferences, and operating-system configurations (https://bit.ly/3j9NkH1)."
         * Windows provides a GUI application, Registry Editor (aka RegEdit), that displays the information the
         * registry has.
         *
         * When we "build" (compile) this project, it is turned into a .NET assembly in the form of a dynamic linked
         * library (.dll file). We need to register this .dll file with the Windows registry in order for SolidWorks 
         * to see this assembly. To do this, we use the RegAsm.exe (which stands for register assembly) executable
         * which is located for me at         C:\Windows\Microsoft.NET\Framework64\v4.0.30319
         * While running command prompt as an administrator, cd into this directory and run the command:
         * "RegAsm.exe /codebase "<absolute or relative path of .dll file>"
         *
         * When this RegAsm.exe executable is run, RegAsm.exe will take care of registering this .dll with the 
         * Windows Registry, and it will then look for a method in the .dll with the [ComRegisterFunction()]
         * attribute. This attribute "...specifies the method to call when you register an assembly for use from COM;
         * this enables the execution of user-written code during the registration process" and the
         * [ComUnregisterFunction()] attribute allows user-code to be specified to aid in deregistration the
         * .dll file.
         *
         * The use the following two methods, COMRegistration() and COMUnregister(), are the user-defined code to
         * to aid in registering the .dll file with RegEdit so that it can add a new subkey for SolidWorks
         *
         * A subkey contains application information needed to support COM functionality.
         *
         * -Eric Gustafson
         */
        #region COM registration
        /// <summary>
        ///     User-defined code for registering this .NET assembly with RegEdit
        /// </summary>
        /// <param name="t"></param>
        [ComRegisterFunction()]
        private static void COMRegistration(Type t) {
            //set the registryKey variable to a new subkey of where the SolidWorks add-ins are located
            using (var registryKey = Microsoft.Win32.Registry.LocalMachine.CreateSubKey($@"SOFTWARE\SolidWorks\AddIns\{t.GUID:b}")) {
                registryKey.SetValue(null, 1); //not sure what this does

                //This sets the name of the Add-In in SolidWorks to "Taskpane Custom Function Manager"
                registryKey.SetValue("Title", "Taskpane Custom Function Manager");

                //In, Tools >> Add-Ins, hovering over the add-in gives a description of "User-defined functions"
                registryKey.SetValue("Description", "User-defined functions");
            }
        }
        /// <summary>
        ///     NOTE: This method is only run if the .dll containing this method is explicitly unregistered via
        ///     C:\Windows\Microsoft.NET\Framework64\v4.0.30319>RegAsm.exe /unregister "<absolute or relative path of .dll file>"
        ///
        ///     If the .DLL is deleted before the it is unregistered, it will not work in SolidWorks, but since it wasn't
        ///     unregistered, RegEdit will still have a subkey of the add-in AND SolidWorks will still display the
        ///     Add-In in Tools >> Add-Ins. The fix for this is to manually delete the subkey in RegEdit. The subkey
        ///     should be located in this directory of RegEdit: Computer\HKEY_LOCAL_MACHINE\SOFTWARE\SolidWorks\AddIns
        /// </summary>
        /// <param name="t"></param>
        [ComUnregisterFunction()]
        private static void COMUnregister(Type t) {
            //Deletes subkey in RegEdit
            Microsoft.Win32.Registry.LocalMachine.DeleteSubKeyTree($@"SOFTWARE\SolidWorks\AddIns\{t.GUID:b}");
        }
        #endregion
    }
}