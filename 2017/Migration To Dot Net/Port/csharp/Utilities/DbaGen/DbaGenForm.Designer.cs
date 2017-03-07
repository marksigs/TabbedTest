namespace DbaGen
{
	partial class DbaGenForm
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
			System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(DbaGenForm));
			this.labelDescription = new System.Windows.Forms.Label();
			this.radioButtonFormatCSharp = new System.Windows.Forms.RadioButton();
			this.radioButtonFormatVB = new System.Windows.Forms.RadioButton();
			this.labelRoot = new System.Windows.Forms.Label();
			this.textBoxRoot = new System.Windows.Forms.TextBox();
			this.labelResult = new System.Windows.Forms.Label();
			this.textBoxResult = new System.Windows.Forms.TextBox();
			this.buttonGenerate = new System.Windows.Forms.Button();
			this.buttonExit = new System.Windows.Forms.Button();
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.checkBoxCopyToClipboard = new System.Windows.Forms.CheckBox();
			this.groupBox1.SuspendLayout();
			this.SuspendLayout();
			// 
			// labelDescription
			// 
			this.labelDescription.Location = new System.Drawing.Point(12, 9);
			this.labelDescription.Name = "labelDescription";
			this.labelDescription.Size = new System.Drawing.Size(268, 70);
			this.labelDescription.TabIndex = 0;
			this.labelDescription.Text = resources.GetString("labelDescription.Text");
			// 
			// radioButtonFormatCSharp
			// 
			this.radioButtonFormatCSharp.AutoSize = true;
			this.radioButtonFormatCSharp.Checked = true;
			this.radioButtonFormatCSharp.Location = new System.Drawing.Point(48, 19);
			this.radioButtonFormatCSharp.Name = "radioButtonFormatCSharp";
			this.radioButtonFormatCSharp.Size = new System.Drawing.Size(63, 17);
			this.radioButtonFormatCSharp.TabIndex = 0;
			this.radioButtonFormatCSharp.TabStop = true;
			this.radioButtonFormatCSharp.Text = "&C#/C++";
			this.radioButtonFormatCSharp.UseVisualStyleBackColor = true;
			// 
			// radioButtonFormatVB
			// 
			this.radioButtonFormatVB.AutoSize = true;
			this.radioButtonFormatVB.Location = new System.Drawing.Point(117, 19);
			this.radioButtonFormatVB.Name = "radioButtonFormatVB";
			this.radioButtonFormatVB.Size = new System.Drawing.Size(82, 17);
			this.radioButtonFormatVB.TabIndex = 1;
			this.radioButtonFormatVB.TabStop = true;
			this.radioButtonFormatVB.Text = "&Visual Basic";
			this.radioButtonFormatVB.UseVisualStyleBackColor = true;
			// 
			// labelRoot
			// 
			this.labelRoot.AutoSize = true;
			this.labelRoot.Location = new System.Drawing.Point(12, 153);
			this.labelRoot.Name = "labelRoot";
			this.labelRoot.Size = new System.Drawing.Size(33, 13);
			this.labelRoot.TabIndex = 2;
			this.labelRoot.Text = "&Root:";
			// 
			// textBoxRoot
			// 
			this.textBoxRoot.Location = new System.Drawing.Point(60, 150);
			this.textBoxRoot.Name = "textBoxRoot";
			this.textBoxRoot.Size = new System.Drawing.Size(210, 20);
			this.textBoxRoot.TabIndex = 3;
			// 
			// labelResult
			// 
			this.labelResult.AutoSize = true;
			this.labelResult.Location = new System.Drawing.Point(12, 182);
			this.labelResult.Name = "labelResult";
			this.labelResult.Size = new System.Drawing.Size(40, 13);
			this.labelResult.TabIndex = 4;
			this.labelResult.Text = "Re&sult:";
			// 
			// textBoxResult
			// 
			this.textBoxResult.Location = new System.Drawing.Point(60, 179);
			this.textBoxResult.Name = "textBoxResult";
			this.textBoxResult.ReadOnly = true;
			this.textBoxResult.Size = new System.Drawing.Size(210, 20);
			this.textBoxResult.TabIndex = 5;
			// 
			// buttonGenerate
			// 
			this.buttonGenerate.Location = new System.Drawing.Point(60, 239);
			this.buttonGenerate.Name = "buttonGenerate";
			this.buttonGenerate.Size = new System.Drawing.Size(75, 23);
			this.buttonGenerate.TabIndex = 7;
			this.buttonGenerate.Text = "&Generate";
			this.buttonGenerate.UseVisualStyleBackColor = true;
			this.buttonGenerate.Click += new System.EventHandler(this.buttonGenerate_Click);
			// 
			// buttonExit
			// 
			this.buttonExit.DialogResult = System.Windows.Forms.DialogResult.Cancel;
			this.buttonExit.Location = new System.Drawing.Point(195, 239);
			this.buttonExit.Name = "buttonExit";
			this.buttonExit.Size = new System.Drawing.Size(75, 23);
			this.buttonExit.TabIndex = 8;
			this.buttonExit.Text = "E&xit";
			this.buttonExit.UseVisualStyleBackColor = true;
			this.buttonExit.Click += new System.EventHandler(this.buttonExit_Click);
			// 
			// groupBox1
			// 
			this.groupBox1.Controls.Add(this.radioButtonFormatCSharp);
			this.groupBox1.Controls.Add(this.radioButtonFormatVB);
			this.groupBox1.Location = new System.Drawing.Point(12, 82);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(258, 53);
			this.groupBox1.TabIndex = 1;
			this.groupBox1.TabStop = false;
			this.groupBox1.Text = "Format";
			// 
			// checkBoxCopyToClipboard
			// 
			this.checkBoxCopyToClipboard.AutoSize = true;
			this.checkBoxCopyToClipboard.Checked = true;
			this.checkBoxCopyToClipboard.CheckState = System.Windows.Forms.CheckState.Checked;
			this.checkBoxCopyToClipboard.Location = new System.Drawing.Point(60, 205);
			this.checkBoxCopyToClipboard.Name = "checkBoxCopyToClipboard";
			this.checkBoxCopyToClipboard.Size = new System.Drawing.Size(175, 17);
			this.checkBoxCopyToClipboard.TabIndex = 6;
			this.checkBoxCopyToClipboard.Text = "Co&py the result to the clipboard.";
			this.checkBoxCopyToClipboard.UseVisualStyleBackColor = true;
			// 
			// DbaGenForm
			// 
			this.AcceptButton = this.buttonGenerate;
			this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
			this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
			this.CancelButton = this.buttonExit;
			this.ClientSize = new System.Drawing.Size(293, 277);
			this.Controls.Add(this.checkBoxCopyToClipboard);
			this.Controls.Add(this.buttonExit);
			this.Controls.Add(this.buttonGenerate);
			this.Controls.Add(this.textBoxResult);
			this.Controls.Add(this.labelResult);
			this.Controls.Add(this.textBoxRoot);
			this.Controls.Add(this.labelRoot);
			this.Controls.Add(this.labelDescription);
			this.Controls.Add(this.groupBox1);
			this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
			this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
			this.MaximizeBox = false;
			this.MinimizeBox = false;
			this.Name = "DbaGenForm";
			this.Text = "DLL Base Address Generator";
			this.Shown += new System.EventHandler(this.DbaGenForm_Shown);
			this.groupBox1.ResumeLayout(false);
			this.groupBox1.PerformLayout();
			this.ResumeLayout(false);
			this.PerformLayout();

		}

		#endregion

		private System.Windows.Forms.Label labelDescription;
		private System.Windows.Forms.RadioButton radioButtonFormatCSharp;
		private System.Windows.Forms.RadioButton radioButtonFormatVB;
		private System.Windows.Forms.Label labelRoot;
		private System.Windows.Forms.TextBox textBoxRoot;
		private System.Windows.Forms.Label labelResult;
		private System.Windows.Forms.TextBox textBoxResult;
		private System.Windows.Forms.Button buttonGenerate;
		private System.Windows.Forms.Button buttonExit;
		private System.Windows.Forms.GroupBox groupBox1;
		private System.Windows.Forms.CheckBox checkBoxCopyToClipboard;
	}
}

