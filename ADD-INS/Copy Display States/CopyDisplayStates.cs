/*
 * 
 *
 * @author        Eric Gustafson
 * @date_created  June 20, 2021, 04:37:42 PM
 * @version       1.0
 */

using SolidWorks.Interop.swpublished;
using SolidWorks.Interop.sldworks;
using System;
using System.Windows.Forms;

using System.IO;
using System.Runtime.InteropServices;


namespace Gustafson.SolidWorks.TaskpaneIntegration {
    public class CopyStatesAddIn : ISwAddin {
        #region SolidWorks Members

        private int mySolidWorksCookie;
        private TaskpaneView mySolidWorksTaskPane;
        private SldWorks mySolidWorks;

        private TaskpaneAddInManager manager;
        private const string PROGID = "Gustafson.SolidWorks.TaskpaneIntegration";

        #endregion

        /// <summary>
        ///   
        /// </summary>
        /// <param name="thisSw"></param>
        /// <param name="cookie"></param>
        /// <returns></returns>
        /// <exception cref="NotImplementedException"></exception>
        public bool ConnectToSW(object thisSw, int cookie) {
            mySolidWorks = (SldWorks) thisSw;
            mySolidWorksCookie = cookie;
            mySolidWorks.SetAddinCallbackInfo2(0, this, cookie);

            mySolidWorksTaskPane =
                mySolidWorks.CreateTaskpaneView2(@"C:\Users\13016\Documents\COMP SCI - C#\Gustafson.SolidWorks.TaskpaneAddIns\SW Custom Add-in Taskpane Icon.png", "Custom SolidWorks Add-Ins");
                /*mySolidWorks.CreateTaskpaneView2($"{Directory.GetCurrentDirectory()}\\SW Macros\\Copy Display States between Configurations\\Copy Display States icon.png",
                                         "Copy the display states from one configuration to another");*/

            manager = (TaskpaneAddInManager)mySolidWorksTaskPane.AddControl(PROGID, string.Empty);

            return true;
        }


        public bool DisconnectFromSW() {
            manager = null;

            mySolidWorksTaskPane.DeleteView();

            Marshal.ReleaseComObject(mySolidWorksTaskPane);

            mySolidWorksTaskPane = null;

            return true;
        }

        #region COM registration

        [ComRegisterFunction()]
        private static void COMRegistration(Type t) {
            using (var rk = Microsoft.Win32.Registry.LocalMachine.CreateSubKey($@"SOFTWARE\SolidWorks\AddIns\{t.GUID:b}")) {
                rk.SetValue(null, 1);
                rk.SetValue("Title", "Taskpane Custom Function Manager");
                rk.SetValue("Description", "All your pixels are belong to us!");
            }
        }

        [ComUnregisterFunction()]
        private static void COMUnregister(Type t) {
            Microsoft.Win32.Registry.LocalMachine.DeleteSubKeyTree($@"SOFTWARE\SolidWorks\AddIns\{t.GUID:b}");
        }
        #endregion

        /// <summary>
        ///  The main entry point for the application.
        /// </summary>
        [STAThread]
        public static void Main(string[] args) {
            //Application.SetHighDpiMode(HighDpiMode.SystemAware);
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new FormApp());
        }
    }





    /**
     * For the dialog box that pops-up
     */
    internal class FormApp : Form {
        public FormApp() {
            InitializeComponent();
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(FormApp));
            this.fromConfigLabel = new System.Windows.Forms.Label();
            this.fromConfigComboBox = new System.Windows.Forms.ComboBox();
            this.button = new System.Windows.Forms.Button();
            this.toConfigLabel = new System.Windows.Forms.Label();
            this.toConfigComboBox = new System.Windows.Forms.ComboBox();
            this.SuspendLayout();
            // 
            // fromConfigLabel
            // 
            this.fromConfigLabel.Font = new System.Drawing.Font("Century Gothic", 7.8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.fromConfigLabel.Location = new System.Drawing.Point(19, 20);
            this.fromConfigLabel.Name = "fromConfigLabel";
            this.fromConfigLabel.Size = new System.Drawing.Size(413, 22);
            this.fromConfigLabel.TabIndex = 0;
            this.fromConfigLabel.Text = "Choose which configuration to copy display states FROM:";
            this.fromConfigLabel.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // fromConfigComboBox
            // 
            this.fromConfigComboBox.FormattingEnabled = true;
            this.fromConfigComboBox.Location = new System.Drawing.Point(52, 45);
            this.fromConfigComboBox.Name = "fromConfigComboBox";
            this.fromConfigComboBox.Size = new System.Drawing.Size(358, 23);
            this.fromConfigComboBox.TabIndex = 1;
            // 
            // button
            // 
            this.button.Font = new System.Drawing.Font("Century Gothic", 7.8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.button.Location = new System.Drawing.Point(133, 183);
            this.button.Name = "button";
            this.button.Size = new System.Drawing.Size(197, 35);
            this.button.TabIndex = 2;
            this.button.Text = "Copy Display States";
            this.button.UseVisualStyleBackColor = true;
            // 
            // toConfigLabel
            // 
            this.toConfigLabel.Font = new System.Drawing.Font("Century Gothic", 7.8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.toConfigLabel.Location = new System.Drawing.Point(52, 110);
            this.toConfigLabel.Name = "toConfigLabel";
            this.toConfigLabel.Size = new System.Drawing.Size(358, 22);
            this.toConfigLabel.TabIndex = 3;
            this.toConfigLabel.Text = "Choose which configuration to copy display state TO:";
            // 
            // toConfigComboBox
            // 
            this.toConfigComboBox.FormattingEnabled = true;
            this.toConfigComboBox.Location = new System.Drawing.Point(52, 135);
            this.toConfigComboBox.Name = "toConfigComboBox";
            this.toConfigComboBox.Size = new System.Drawing.Size(358, 23);
            this.toConfigComboBox.TabIndex = 4;
            // 
            // FormApp
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 15F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(464, 230);
            this.Controls.Add(this.toConfigComboBox);
            this.Controls.Add(this.toConfigLabel);
            this.Controls.Add(this.button);
            this.Controls.Add(this.fromConfigComboBox);
            this.Controls.Add(this.fromConfigLabel);
            this.Font = new System.Drawing.Font("Times New Roman", 7.8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            //this.Icon = ((System.Drawing.Icon) (resources.GetObject("$this.Icon")));
            this.Name = "FormApp";
            this.Text = "Copy Display States";
            this.ResumeLayout(false);
        }

        private Button button;
        private ComboBox fromConfigComboBox;
        private ComboBox toConfigComboBox;
        private Label fromConfigLabel;
        private Label toConfigLabel;

        #endregion

        private void ButtonClicked(object sender, EventArgs e) {
            throw new System.NotImplementedException();
        }
    }
}