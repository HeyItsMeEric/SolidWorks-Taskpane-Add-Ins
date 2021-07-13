/*
 * A WinForm that has info about me!
 *
 * @author        Eric Gustafson
 * @date_created  July 12, 2021, 07:11:00 PM
 * @version       1.0
 */

using System.ComponentModel;
using System.Windows.Forms;

namespace COM_Assembly_Registration_App {
    public class AboutTheAuthorForm : Form {

        
        public AboutTheAuthorForm() {
            InitializeComponent();

            #region Custom UI Component settings
                pictureBox.ImageLocation = @"https://avatars.githubusercontent.com/u/82784765?v=4";
                pictureBox.SizeMode = PictureBoxSizeMode.Zoom;
            #endregion
            
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(AboutTheAuthorForm));
            this.pictureBox = new System.Windows.Forms.PictureBox();
            this.githubLabel = new System.Windows.Forms.LinkLabel();
            this.aboutMeLabel = new System.Windows.Forms.Label();
            this.linkedInLink = new System.Windows.Forms.LinkLabel();
            this.contactInfo = new System.Windows.Forms.LinkLabel();
            ((System.ComponentModel.ISupportInitialize) (this.pictureBox)).BeginInit();
            this.SuspendLayout();
            // 
            // pictureBox1
            // 
            this.pictureBox.ImageLocation = @"https://avatars.githubusercontent.com/u/82784765?v=4";
            this.pictureBox.Location = new System.Drawing.Point(12, 12);
            this.pictureBox.Name = "pictureBox";
            this.pictureBox.Size = new System.Drawing.Size(200, 258);
            this.pictureBox.TabIndex = 0;
            this.pictureBox.TabStop = false;
            // 
            // githubLabel
            // 
            this.githubLabel.Font = new System.Drawing.Font("Microsoft Sans Serif", 20.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte) (0)));
            this.githubLabel.Location = new System.Drawing.Point(12, 378);
            this.githubLabel.Name = "githubLabel";
            this.githubLabel.Size = new System.Drawing.Size(275, 47);
            this.githubLabel.TabIndex = 2;
            this.githubLabel.TabStop = true;
            this.githubLabel.Text = "Visit my Github";
            this.githubLabel.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(GithubLabelLinkClicked);
            // 
            // aboutMeLabel
            // 
            this.aboutMeLabel.Font = new System.Drawing.Font("Microsoft Sans Serif", 14.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte) (0)));
            this.aboutMeLabel.Location = new System.Drawing.Point(226, 30);
            this.aboutMeLabel.Name = "aboutMeLabel";
            this.aboutMeLabel.Size = new System.Drawing.Size(490, 162);
            this.aboutMeLabel.TabIndex = 3;
            this.aboutMeLabel.Text = resources.GetString("aboutMeLabel.Text");
            // 
            // linkedInLink
            // 
            this.linkedInLink.Font = new System.Drawing.Font("Microsoft Sans Serif", 20.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte) (0)));
            this.linkedInLink.Location = new System.Drawing.Point(12, 331);
            this.linkedInLink.Name = "linkedInLink";
            this.linkedInLink.Size = new System.Drawing.Size(275, 47);
            this.linkedInLink.TabIndex = 4;
            this.linkedInLink.TabStop = true;
            this.linkedInLink.Text = "Visit my LinkedIn\r\n";
            this.linkedInLink.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(LinkedInLinkLinkClicked);
            // 
            // contactInfo
            // 
            this.contactInfo.Font = new System.Drawing.Font("Microsoft Sans Serif", 20.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte) (0)));
            this.contactInfo.Location = new System.Drawing.Point(12, 288);
            this.contactInfo.Name = "contactInfo";
            this.contactInfo.Size = new System.Drawing.Size(275, 47);
            this.contactInfo.TabIndex = 5;
            this.contactInfo.TabStop = true;
            this.contactInfo.Text = "Contact Me!";
            this.contactInfo.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.contactInfo_LinkClicked);
            // 
            // AboutTheAuthorForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(800, 450);
            this.Controls.Add(this.contactInfo);
            this.Controls.Add(this.linkedInLink);
            this.Controls.Add(this.aboutMeLabel);
            this.Controls.Add(this.githubLabel);
            this.Controls.Add(this.pictureBox);
            this.Name = "AboutTheAuthorForm";
            this.Text = "AboutTheAuthorForm";
            ((System.ComponentModel.ISupportInitialize) (this.pictureBox)).EndInit();
            this.ResumeLayout(false);
        }

        private System.Windows.Forms.LinkLabel linkedInLink;
        private System.Windows.Forms.LinkLabel contactInfo;
        private System.Windows.Forms.Label aboutMeLabel;
        private System.Windows.Forms.LinkLabel githubLabel;
        private System.Windows.Forms.PictureBox pictureBox;
        
        #endregion
        
        
        
        

        /// <summary>
        /// Opens my GitHub profile in the user's default browser
        /// 
        /// Method Contents from 
        /// https://docs.microsoft.com/en-us/dotnet/api/system.windows.forms.linklabel?view=net-5.0#examples
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void GithubLabelLinkClicked(object sender, LinkLabelLinkClickedEventArgs e) {
            // Specify that the link was visited.
            this.githubLabel.LinkVisited = true;

            // Navigate to a URL.
            System.Diagnostics.Process.Start(@"https://github.com/HeyItsMeEric");
        }

        /// <summary>
        /// Opens my LinkedIn profile in the user's default browser
        /// 
        /// Method Contents from 
        /// https://docs.microsoft.com/en-us/dotnet/api/system.windows.forms.linklabel?view=net-5.0#examples
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void LinkedInLinkLinkClicked(object sender, LinkLabelLinkClickedEventArgs e) {
            // Specify that the link was visited.
            this.githubLabel.LinkVisited = true;

            // Navigate to a URL.
            System.Diagnostics.Process.Start(@"https://www.linkedin.com/in/eric-gustafson-at-georgia-tech/");
        }

        /// <summary>
        /// Starts a new email for the user to email me
        /// 
        /// Method Contents from 
        /// https://docs.microsoft.com/en-us/dotnet/api/system.windows.forms.linklabel?view=net-5.0#examples
        /// and
        /// https://stackoverflow.com/questions/2030414/how-to-create-a-mailto-link-in-windows-forms-application/2030556
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void contactInfo_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e) {
            // Specify that the link was visited.
            this.githubLabel.LinkVisited = true;

            // Navigate to a URL.
            System.Diagnostics.Process.Start(@"mailto: Eric.Gustafson@Gatech.Edu");
        }
    }
}