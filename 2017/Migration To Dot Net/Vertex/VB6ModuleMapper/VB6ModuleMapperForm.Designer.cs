namespace VB6ModuleMapper
{
	partial class VB6ModuleMapperForm
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
			this.buttonExit = new System.Windows.Forms.Button();
			this.labelMappingFile = new System.Windows.Forms.Label();
			this.textBoxMappingFile = new System.Windows.Forms.TextBox();
			this.buttonBrowseMappingFile = new System.Windows.Forms.Button();
			this.openFileDialogMappingFile = new System.Windows.Forms.OpenFileDialog();
			this.labelRootDirectory = new System.Windows.Forms.Label();
			this.textBoxRootDirectory = new System.Windows.Forms.TextBox();
			this.buttonBrowseRootDirectory = new System.Windows.Forms.Button();
			this.folderBrowserDialogRootDirectory = new System.Windows.Forms.FolderBrowserDialog();
			this.buttonGenerate = new System.Windows.Forms.Button();
			this.SuspendLayout();
			// 
			// buttonExit
			// 
			this.buttonExit.Location = new System.Drawing.Point(359, 80);
			this.buttonExit.Name = "buttonExit";
			this.buttonExit.Size = new System.Drawing.Size(75, 23);
			this.buttonExit.TabIndex = 7;
			this.buttonExit.Text = "E&xit";
			this.buttonExit.UseVisualStyleBackColor = true;
			this.buttonExit.Click += new System.EventHandler(this.buttonExit_Click);
			// 
			// labelMappingFile
			// 
			this.labelMappingFile.AutoSize = true;
			this.labelMappingFile.Location = new System.Drawing.Point(12, 9);
			this.labelMappingFile.Name = "labelMappingFile";
			this.labelMappingFile.Size = new System.Drawing.Size(70, 13);
			this.labelMappingFile.TabIndex = 0;
			this.labelMappingFile.Text = "MappingFile &File:";
			// 
			// textBoxMappingFile
			// 
			this.textBoxMappingFile.Location = new System.Drawing.Point(108, 6);
			this.textBoxMappingFile.Name = "textBoxMappingFile";
			this.textBoxMappingFile.Size = new System.Drawing.Size(326, 20);
			this.textBoxMappingFile.TabIndex = 1;
			this.textBoxMappingFile.TextChanged += new System.EventHandler(this.textBoxMappingFile_TextChanged);
			// 
			// buttonBrowseMappingFile
			// 
			this.buttonBrowseMappingFile.Location = new System.Drawing.Point(440, 3);
			this.buttonBrowseMappingFile.Name = "buttonBrowseMappingFile";
			this.buttonBrowseMappingFile.Size = new System.Drawing.Size(75, 23);
			this.buttonBrowseMappingFile.TabIndex = 2;
			this.buttonBrowseMappingFile.Text = "&Browse...";
			this.buttonBrowseMappingFile.UseVisualStyleBackColor = true;
			this.buttonBrowseMappingFile.Click += new System.EventHandler(this.buttonBrowseInputFile_Click);
			// 
			// openFileDialogMappingFile
			// 
			this.openFileDialogMappingFile.Filter = "csv files|*.csv|All files|*.*";
			this.openFileDialogMappingFile.Title = "Select the MappingFile File";
			// 
			// labelRootDirectory
			// 
			this.labelRootDirectory.AutoSize = true;
			this.labelRootDirectory.Location = new System.Drawing.Point(12, 35);
			this.labelRootDirectory.Name = "labelRootDirectory";
			this.labelRootDirectory.Size = new System.Drawing.Size(78, 13);
			this.labelRootDirectory.TabIndex = 3;
			this.labelRootDirectory.Text = "Root &Directory:";
			// 
			// textBoxRootDirectory
			// 
			this.textBoxRootDirectory.Location = new System.Drawing.Point(108, 32);
			this.textBoxRootDirectory.Name = "textBoxRootDirectory";
			this.textBoxRootDirectory.Size = new System.Drawing.Size(326, 20);
			this.textBoxRootDirectory.TabIndex = 4;
			this.textBoxRootDirectory.TextChanged += new System.EventHandler(this.textBoxRootDirectory_TextChanged);
			// 
			// buttonBrowseRootDirectory
			// 
			this.buttonBrowseRootDirectory.Location = new System.Drawing.Point(440, 30);
			this.buttonBrowseRootDirectory.Name = "buttonBrowseRootDirectory";
			this.buttonBrowseRootDirectory.Size = new System.Drawing.Size(75, 23);
			this.buttonBrowseRootDirectory.TabIndex = 5;
			this.buttonBrowseRootDirectory.Text = "B&rowse...";
			this.buttonBrowseRootDirectory.UseVisualStyleBackColor = true;
			this.buttonBrowseRootDirectory.Click += new System.EventHandler(this.buttonBrowseRootDirectory_Click);
			// 
			// folderBrowserDialogRootDirectory
			// 
			this.folderBrowserDialogRootDirectory.Description = "Browse for the root directory for the Visual Basic 6 source code.";
			this.folderBrowserDialogRootDirectory.ShowNewFolderButton = false;
			// 
			// buttonGenerate
			// 
			this.buttonGenerate.Location = new System.Drawing.Point(108, 80);
			this.buttonGenerate.Name = "buttonGenerate";
			this.buttonGenerate.Size = new System.Drawing.Size(75, 23);
			this.buttonGenerate.TabIndex = 6;
			this.buttonGenerate.Text = "&Generate";
			this.buttonGenerate.UseVisualStyleBackColor = true;
			this.buttonGenerate.Click += new System.EventHandler(this.buttonGenerate_Click);
			// 
			// VB6ModuleMapperForm
			// 
			this.AcceptButton = this.buttonExit;
			this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
			this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
			this.ClientSize = new System.Drawing.Size(523, 115);
			this.Controls.Add(this.buttonGenerate);
			this.Controls.Add(this.buttonBrowseRootDirectory);
			this.Controls.Add(this.textBoxRootDirectory);
			this.Controls.Add(this.labelRootDirectory);
			this.Controls.Add(this.buttonBrowseMappingFile);
			this.Controls.Add(this.textBoxMappingFile);
			this.Controls.Add(this.labelMappingFile);
			this.Controls.Add(this.buttonExit);
			this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
			this.MaximizeBox = false;
			this.Name = "VB6ModuleMapperForm";
			this.Text = "VB6 Module Mapper";
			this.FormClosed += new System.Windows.Forms.FormClosedEventHandler(this.VB6ModuleMapperForm_FormClosed);
			this.ResumeLayout(false);
			this.PerformLayout();

		}

		#endregion

		private System.Windows.Forms.Button buttonExit;
		private System.Windows.Forms.Label labelMappingFile;
		private System.Windows.Forms.TextBox textBoxMappingFile;
		private System.Windows.Forms.Button buttonBrowseMappingFile;
		private System.Windows.Forms.OpenFileDialog openFileDialogMappingFile;
		private System.Windows.Forms.Label labelRootDirectory;
		private System.Windows.Forms.TextBox textBoxRootDirectory;
		private System.Windows.Forms.Button buttonBrowseRootDirectory;
		private System.Windows.Forms.FolderBrowserDialog folderBrowserDialogRootDirectory;
		private System.Windows.Forms.Button buttonGenerate;
	}
}

