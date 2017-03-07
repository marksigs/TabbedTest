namespace PreMigrationConverter
{
    partial class frmMain
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
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
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
			this.components = new System.ComponentModel.Container();
			this.btnStartStop = new System.Windows.Forms.Button();
			this.folderBrowserDialog1 = new System.Windows.Forms.FolderBrowserDialog();
			this.btnSelectInputFolder = new System.Windows.Forms.Button();
			this.label1 = new System.Windows.Forms.Label();
			this.txtInputFolder = new System.Windows.Forms.TextBox();
			this.btnExit = new System.Windows.Forms.Button();
			this.label2 = new System.Windows.Forms.Label();
			this.txtLogFile = new System.Windows.Forms.TextBox();
			this.txtSelectLogFile = new System.Windows.Forms.Button();
			this.openFileDialog1 = new System.Windows.Forms.OpenFileDialog();
			this.statusStrip1 = new System.Windows.Forms.StatusStrip();
			this.toolStripStatusLabel1 = new System.Windows.Forms.ToolStripStatusLabel();
			this.btnViewLogFile = new System.Windows.Forms.Button();
			this.rbtnCommentCode = new System.Windows.Forms.RadioButton();
			this.rbtnRemoveCode = new System.Windows.Forms.RadioButton();
			this.txtDefectId = new System.Windows.Forms.TextBox();
			this.label3 = new System.Windows.Forms.Label();
			this.toolTip = new System.Windows.Forms.ToolTip(this.components);
			this.chkOverwrite = new System.Windows.Forms.CheckBox();
			this.chkVerboseLogging = new System.Windows.Forms.CheckBox();
			this.statusStrip1.SuspendLayout();
			this.SuspendLayout();
			// 
			// btnStartStop
			// 
			this.btnStartStop.Location = new System.Drawing.Point(432, 160);
			this.btnStartStop.Name = "btnStartStop";
			this.btnStartStop.Size = new System.Drawing.Size(121, 23);
			this.btnStartStop.TabIndex = 0;
			this.btnStartStop.Text = "&Start";
			this.btnStartStop.UseVisualStyleBackColor = true;
			this.btnStartStop.Click += new System.EventHandler(this.btnStart_Click);
			// 
			// btnSelectInputFolder
			// 
			this.btnSelectInputFolder.Location = new System.Drawing.Point(524, 8);
			this.btnSelectInputFolder.Name = "btnSelectInputFolder";
			this.btnSelectInputFolder.Size = new System.Drawing.Size(27, 23);
			this.btnSelectInputFolder.TabIndex = 2;
			this.btnSelectInputFolder.Text = "...";
			this.btnSelectInputFolder.UseVisualStyleBackColor = true;
			this.btnSelectInputFolder.Click += new System.EventHandler(this.btnSelectInputFolder_Click);
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.Location = new System.Drawing.Point(9, 13);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(57, 13);
			this.label1.TabIndex = 0;
			this.label1.Text = "Input&folder";
			// 
			// txtInputFolder
			// 
			this.txtInputFolder.Location = new System.Drawing.Point(78, 10);
			this.txtInputFolder.Name = "txtInputFolder";
			this.txtInputFolder.Size = new System.Drawing.Size(440, 20);
			this.txtInputFolder.TabIndex = 1;
			// 
			// btnExit
			// 
			this.btnExit.Location = new System.Drawing.Point(432, 220);
			this.btnExit.Name = "btnExit";
			this.btnExit.Size = new System.Drawing.Size(121, 23);
			this.btnExit.TabIndex = 15;
			this.btnExit.Text = "E&xit";
			this.btnExit.UseVisualStyleBackColor = true;
			this.btnExit.Click += new System.EventHandler(this.btnExit_Click);
			// 
			// label2
			// 
			this.label2.AutoSize = true;
			this.label2.Location = new System.Drawing.Point(9, 50);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(41, 13);
			this.label2.TabIndex = 3;
			this.label2.Text = "&Log file";
			// 
			// txtLogFile
			// 
			this.txtLogFile.Location = new System.Drawing.Point(78, 47);
			this.txtLogFile.Name = "txtLogFile";
			this.txtLogFile.Size = new System.Drawing.Size(440, 20);
			this.txtLogFile.TabIndex = 4;
			// 
			// txtSelectLogFile
			// 
			this.txtSelectLogFile.Location = new System.Drawing.Point(524, 45);
			this.txtSelectLogFile.Name = "txtSelectLogFile";
			this.txtSelectLogFile.Size = new System.Drawing.Size(27, 23);
			this.txtSelectLogFile.TabIndex = 5;
			this.txtSelectLogFile.Text = "...";
			this.txtSelectLogFile.UseVisualStyleBackColor = true;
			this.txtSelectLogFile.Click += new System.EventHandler(this.txtSelectLogFile_Click);
			// 
			// openFileDialog1
			// 
			this.openFileDialog1.FileName = "openFileDialog1";
			// 
			// statusStrip1
			// 
			this.statusStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.toolStripStatusLabel1});
			this.statusStrip1.Location = new System.Drawing.Point(0, 255);
			this.statusStrip1.Name = "statusStrip1";
			this.statusStrip1.Size = new System.Drawing.Size(565, 22);
			this.statusStrip1.TabIndex = 8;
			this.statusStrip1.Text = "statusStrip1";
			// 
			// toolStripStatusLabel1
			// 
			this.toolStripStatusLabel1.Name = "toolStripStatusLabel1";
			this.toolStripStatusLabel1.Size = new System.Drawing.Size(0, 17);
			// 
			// btnViewLogFile
			// 
			this.btnViewLogFile.Location = new System.Drawing.Point(432, 191);
			this.btnViewLogFile.Name = "btnViewLogFile";
			this.btnViewLogFile.Size = new System.Drawing.Size(121, 23);
			this.btnViewLogFile.TabIndex = 14;
			this.btnViewLogFile.Text = "&View Log File";
			this.btnViewLogFile.UseVisualStyleBackColor = true;
			this.btnViewLogFile.Click += new System.EventHandler(this.btnViewLogFile_Click);
			// 
			// rbtnCommentCode
			// 
			this.rbtnCommentCode.AutoSize = true;
			this.rbtnCommentCode.Checked = true;
			this.rbtnCommentCode.Location = new System.Drawing.Point(12, 84);
			this.rbtnCommentCode.Name = "rbtnCommentCode";
			this.rbtnCommentCode.Size = new System.Drawing.Size(149, 17);
			this.rbtnCommentCode.TabIndex = 6;
			this.rbtnCommentCode.TabStop = true;
			this.rbtnCommentCode.Text = "Comment Unwanted Code";
			this.rbtnCommentCode.UseVisualStyleBackColor = true;
			// 
			// rbtnRemoveCode
			// 
			this.rbtnRemoveCode.AutoSize = true;
			this.rbtnRemoveCode.Location = new System.Drawing.Point(12, 108);
			this.rbtnRemoveCode.Name = "rbtnRemoveCode";
			this.rbtnRemoveCode.Size = new System.Drawing.Size(145, 17);
			this.rbtnRemoveCode.TabIndex = 7;
			this.rbtnRemoveCode.Text = "Remove Unwanted Code";
			this.rbtnRemoveCode.UseVisualStyleBackColor = true;
			// 
			// txtDefectId
			// 
			this.txtDefectId.Location = new System.Drawing.Point(252, 84);
			this.txtDefectId.Name = "txtDefectId";
			this.txtDefectId.Size = new System.Drawing.Size(100, 20);
			this.txtDefectId.TabIndex = 9;
			// 
			// label3
			// 
			this.label3.AutoSize = true;
			this.label3.Location = new System.Drawing.Point(195, 87);
			this.label3.Name = "label3";
			this.label3.Size = new System.Drawing.Size(51, 13);
			this.label3.TabIndex = 8;
			this.label3.Text = "Defect Id";
			// 
			// toolTip
			// 
			this.toolTip.Tag = "";
			// 
			// chkOverwrite
			// 
			this.chkOverwrite.AutoSize = true;
			this.chkOverwrite.Location = new System.Drawing.Point(12, 141);
			this.chkOverwrite.Name = "chkOverwrite";
			this.chkOverwrite.Size = new System.Drawing.Size(92, 17);
			this.chkOverwrite.TabIndex = 10;
			this.chkOverwrite.Text = "Overwrite files";
			this.chkOverwrite.UseVisualStyleBackColor = true;
			// 
			// chkVerboseLogging
			// 
			this.chkVerboseLogging.AutoSize = true;
			this.chkVerboseLogging.Location = new System.Drawing.Point(12, 166);
			this.chkVerboseLogging.Name = "chkVerboseLogging";
			this.chkVerboseLogging.Size = new System.Drawing.Size(102, 17);
			this.chkVerboseLogging.TabIndex = 11;
			this.chkVerboseLogging.Text = "Verbose logging";
			this.chkVerboseLogging.UseVisualStyleBackColor = true;
			// 
			// frmMain
			// 
			this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
			this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
			this.ClientSize = new System.Drawing.Size(565, 277);
			this.Controls.Add(this.chkVerboseLogging);
			this.Controls.Add(this.chkOverwrite);
			this.Controls.Add(this.label3);
			this.Controls.Add(this.txtDefectId);
			this.Controls.Add(this.rbtnRemoveCode);
			this.Controls.Add(this.rbtnCommentCode);
			this.Controls.Add(this.btnViewLogFile);
			this.Controls.Add(this.statusStrip1);
			this.Controls.Add(this.txtSelectLogFile);
			this.Controls.Add(this.txtLogFile);
			this.Controls.Add(this.label2);
			this.Controls.Add(this.btnExit);
			this.Controls.Add(this.txtInputFolder);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.btnSelectInputFolder);
			this.Controls.Add(this.btnStartStop);
			this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
			this.MaximizeBox = false;
			this.Name = "frmMain";
			this.Text = "Pre-Migration Code Converter";
			this.Load += new System.EventHandler(this.Form1_Load);
			this.statusStrip1.ResumeLayout(false);
			this.statusStrip1.PerformLayout();
			this.ResumeLayout(false);
			this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button btnStartStop;
        private System.Windows.Forms.FolderBrowserDialog folderBrowserDialog1;
        private System.Windows.Forms.Button btnSelectInputFolder;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox txtInputFolder;
        private System.Windows.Forms.Button btnExit;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TextBox txtLogFile;
        private System.Windows.Forms.Button txtSelectLogFile;
        private System.Windows.Forms.OpenFileDialog openFileDialog1;
        private System.Windows.Forms.StatusStrip statusStrip1;
        private System.Windows.Forms.ToolStripStatusLabel toolStripStatusLabel1;
        private System.Windows.Forms.Button btnViewLogFile;
        private System.Windows.Forms.RadioButton rbtnCommentCode;
        private System.Windows.Forms.RadioButton rbtnRemoveCode;
        private System.Windows.Forms.TextBox txtDefectId;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.ToolTip toolTip;
        private System.Windows.Forms.CheckBox chkOverwrite;
		private System.Windows.Forms.CheckBox chkVerboseLogging;
    }
}

