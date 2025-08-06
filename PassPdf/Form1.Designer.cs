namespace PassPdf
{
    partial class Form1
    {
        /// <summary>
        ///  Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        ///  Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        ///  Required method for Designer support - do not modify
        ///  the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            openFileDialog1 = new OpenFileDialog();
            btnBrowse = new Button();
            txtFile = new TextBox();
            txtResult = new TextBox();
            txtExportFolder = new TextBox();
            folderBrowserDialog1 = new FolderBrowserDialog();
            btnBrowseExport = new Button();
            btnExport = new Button();
            btnSetPass = new Button();
            SuspendLayout();
            // 
            // openFileDialog1
            // 
            openFileDialog1.FileName = "openFileDialog1";
            // 
            // btnBrowse
            // 
            btnBrowse.Location = new Point(12, 22);
            btnBrowse.Name = "btnBrowse";
            btnBrowse.Size = new Size(75, 23);
            btnBrowse.TabIndex = 0;
            btnBrowse.Text = "Browse File";
            btnBrowse.UseVisualStyleBackColor = true;
            btnBrowse.Click += btnBrowse_Click;
            // 
            // txtFile
            // 
            txtFile.Location = new Point(106, 22);
            txtFile.Name = "txtFile";
            txtFile.Size = new Size(444, 23);
            txtFile.TabIndex = 1;
            // 
            // txtResult
            // 
            txtResult.Location = new Point(12, 118);
            txtResult.Multiline = true;
            txtResult.Name = "txtResult";
            txtResult.ScrollBars = ScrollBars.Both;
            txtResult.Size = new Size(538, 251);
            txtResult.TabIndex = 2;
            // 
            // txtExportFolder
            // 
            txtExportFolder.Location = new Point(106, 68);
            txtExportFolder.Name = "txtExportFolder";
            txtExportFolder.Size = new Size(444, 23);
            txtExportFolder.TabIndex = 1;
            // 
            // btnBrowseExport
            // 
            btnBrowseExport.Location = new Point(12, 68);
            btnBrowseExport.Name = "btnBrowseExport";
            btnBrowseExport.Size = new Size(88, 23);
            btnBrowseExport.TabIndex = 3;
            btnBrowseExport.Text = "Export Folder";
            btnBrowseExport.UseVisualStyleBackColor = true;
            btnBrowseExport.Click += btnBrowseExport_Click;
            // 
            // btnExport
            // 
            btnExport.Location = new Point(587, 21);
            btnExport.Name = "btnExport";
            btnExport.Size = new Size(75, 23);
            btnExport.TabIndex = 4;
            btnExport.Text = "Export PDF";
            btnExport.UseVisualStyleBackColor = true;
            btnExport.Click += btnExport_Click;
            // 
            // btnSetPass
            // 
            btnSetPass.Location = new Point(587, 67);
            btnSetPass.Name = "btnSetPass";
            btnSetPass.Size = new Size(107, 23);
            btnSetPass.TabIndex = 5;
            btnSetPass.Text = "Set Password";
            btnSetPass.UseVisualStyleBackColor = true;
            btnSetPass.Click += btnSetPass_Click;
            // 
            // Form1
            // 
            AutoScaleDimensions = new SizeF(7F, 15F);
            AutoScaleMode = AutoScaleMode.Font;
            ClientSize = new Size(800, 450);
            Controls.Add(btnSetPass);
            Controls.Add(btnExport);
            Controls.Add(btnBrowseExport);
            Controls.Add(txtResult);
            Controls.Add(txtExportFolder);
            Controls.Add(txtFile);
            Controls.Add(btnBrowse);
            Name = "Form1";
            Text = "Set Pass";
            ResumeLayout(false);
            PerformLayout();
        }

        #endregion

        private OpenFileDialog openFileDialog1;
        private Button btnBrowse;
        private TextBox txtFile;
        private TextBox txtResult;
        private TextBox txtExportFolder;
        private FolderBrowserDialog folderBrowserDialog1;
        private Button btnBrowseExport;
        private Button btnExport;
        private Button btnSetPass;
    }
}
