namespace Vertex.Fsd.Omiga.AxWord
{
	partial class PrintDialog
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
			System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(PrintDialog));
			this.labelName = new System.Windows.Forms.Label();
			this.comboBoxName = new System.Windows.Forms.ComboBox();
			this.labelCopies = new System.Windows.Forms.Label();
			this.numericUpDownCopies = new System.Windows.Forms.NumericUpDown();
			this.labelFirstPageBin = new System.Windows.Forms.Label();
			this.comboBoxFirstPageBin = new System.Windows.Forms.ComboBox();
			this.checkBoxUseDifferentBin = new System.Windows.Forms.CheckBox();
			this.labelOtherPagesBin = new System.Windows.Forms.Label();
			this.comboBoxOtherPagesBin = new System.Windows.Forms.ComboBox();
			this.buttonOK = new System.Windows.Forms.Button();
			this.buttonCancel = new System.Windows.Forms.Button();
			((System.ComponentModel.ISupportInitialize)(this.numericUpDownCopies)).BeginInit();
			this.SuspendLayout();
			// 
			// labelName
			// 
			this.labelName.AutoSize = true;
			this.labelName.Location = new System.Drawing.Point(12, 9);
			this.labelName.Name = "labelName";
			this.labelName.Size = new System.Drawing.Size(38, 13);
			this.labelName.TabIndex = 0;
			this.labelName.Text = "&Name:";
			// 
			// comboBoxName
			// 
			this.comboBoxName.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
			this.comboBoxName.FormattingEnabled = true;
			this.comboBoxName.Location = new System.Drawing.Point(118, 6);
			this.comboBoxName.Name = "comboBoxName";
			this.comboBoxName.Size = new System.Drawing.Size(315, 21);
			this.comboBoxName.TabIndex = 1;
			this.comboBoxName.SelectedIndexChanged += new System.EventHandler(this.comboBoxName_SelectedIndexChanged);
			// 
			// labelCopies
			// 
			this.labelCopies.AutoSize = true;
			this.labelCopies.Location = new System.Drawing.Point(12, 35);
			this.labelCopies.Name = "labelCopies";
			this.labelCopies.Size = new System.Drawing.Size(42, 13);
			this.labelCopies.TabIndex = 2;
			this.labelCopies.Text = "&Copies:";
			// 
			// numericUpDownCopies
			// 
			this.numericUpDownCopies.Location = new System.Drawing.Point(118, 33);
			this.numericUpDownCopies.Minimum = new decimal(new int[] {
            1,
            0,
            0,
            0});
			this.numericUpDownCopies.Name = "numericUpDownCopies";
			this.numericUpDownCopies.Size = new System.Drawing.Size(75, 20);
			this.numericUpDownCopies.TabIndex = 3;
			this.numericUpDownCopies.TextAlign = System.Windows.Forms.HorizontalAlignment.Right;
			this.numericUpDownCopies.Value = new decimal(new int[] {
            1,
            0,
            0,
            0});
			// 
			// labelFirstPageBin
			// 
			this.labelFirstPageBin.AutoSize = true;
			this.labelFirstPageBin.Location = new System.Drawing.Point(12, 63);
			this.labelFirstPageBin.Name = "labelFirstPageBin";
			this.labelFirstPageBin.Size = new System.Drawing.Size(76, 13);
			this.labelFirstPageBin.TabIndex = 4;
			this.labelFirstPageBin.Text = "&First page tray:";
			// 
			// comboBoxFirstPageBin
			// 
			this.comboBoxFirstPageBin.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
			this.comboBoxFirstPageBin.FormattingEnabled = true;
			this.comboBoxFirstPageBin.Location = new System.Drawing.Point(118, 60);
			this.comboBoxFirstPageBin.Name = "comboBoxFirstPageBin";
			this.comboBoxFirstPageBin.Size = new System.Drawing.Size(141, 21);
			this.comboBoxFirstPageBin.TabIndex = 5;
			// 
			// checkBoxUseDifferentBin
			// 
			this.checkBoxUseDifferentBin.AutoSize = true;
			this.checkBoxUseDifferentBin.Location = new System.Drawing.Point(266, 63);
			this.checkBoxUseDifferentBin.Name = "checkBoxUseDifferentBin";
			this.checkBoxUseDifferentBin.Size = new System.Drawing.Size(180, 17);
			this.checkBoxUseDifferentBin.TabIndex = 6;
			this.checkBoxUseDifferentBin.Text = "&Use different tray for other pages";
			this.checkBoxUseDifferentBin.UseVisualStyleBackColor = true;
			this.checkBoxUseDifferentBin.CheckedChanged += new System.EventHandler(this.checkBoxUseDifferentBin_CheckedChanged);
			// 
			// labelOtherPagesBin
			// 
			this.labelOtherPagesBin.AutoSize = true;
			this.labelOtherPagesBin.Location = new System.Drawing.Point(12, 91);
			this.labelOtherPagesBin.Name = "labelOtherPagesBin";
			this.labelOtherPagesBin.Size = new System.Drawing.Size(88, 13);
			this.labelOtherPagesBin.TabIndex = 7;
			this.labelOtherPagesBin.Text = "&Other pages tray:";
			// 
			// comboBoxOtherPagesBin
			// 
			this.comboBoxOtherPagesBin.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
			this.comboBoxOtherPagesBin.FormattingEnabled = true;
			this.comboBoxOtherPagesBin.Location = new System.Drawing.Point(118, 88);
			this.comboBoxOtherPagesBin.Name = "comboBoxOtherPagesBin";
			this.comboBoxOtherPagesBin.Size = new System.Drawing.Size(141, 21);
			this.comboBoxOtherPagesBin.TabIndex = 8;
			// 
			// buttonOK
			// 
			this.buttonOK.DialogResult = System.Windows.Forms.DialogResult.OK;
			this.buttonOK.Location = new System.Drawing.Point(118, 124);
			this.buttonOK.Name = "buttonOK";
			this.buttonOK.Size = new System.Drawing.Size(75, 23);
			this.buttonOK.TabIndex = 9;
			this.buttonOK.Text = "&OK";
			this.buttonOK.UseVisualStyleBackColor = true;
			this.buttonOK.Click += new System.EventHandler(this.buttonOK_Click);
			// 
			// buttonCancel
			// 
			this.buttonCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel;
			this.buttonCancel.Location = new System.Drawing.Point(199, 124);
			this.buttonCancel.Name = "buttonCancel";
			this.buttonCancel.Size = new System.Drawing.Size(75, 23);
			this.buttonCancel.TabIndex = 10;
			this.buttonCancel.Text = "&Cancel";
			this.buttonCancel.UseVisualStyleBackColor = true;
			// 
			// PrintDialog
			// 
			this.AcceptButton = this.buttonOK;
			this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
			this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
			this.CancelButton = this.buttonCancel;
			this.ClientSize = new System.Drawing.Size(455, 159);
			this.Controls.Add(this.buttonCancel);
			this.Controls.Add(this.buttonOK);
			this.Controls.Add(this.comboBoxOtherPagesBin);
			this.Controls.Add(this.labelOtherPagesBin);
			this.Controls.Add(this.checkBoxUseDifferentBin);
			this.Controls.Add(this.comboBoxFirstPageBin);
			this.Controls.Add(this.labelFirstPageBin);
			this.Controls.Add(this.numericUpDownCopies);
			this.Controls.Add(this.labelCopies);
			this.Controls.Add(this.comboBoxName);
			this.Controls.Add(this.labelName);
			this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
			this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
			this.MaximizeBox = false;
			this.MinimizeBox = false;
			this.Name = "PrintDialog";
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
			this.Text = "Print";
			this.TopMost = true;
			this.Shown += new System.EventHandler(this.PrintDialog_Shown);
			((System.ComponentModel.ISupportInitialize)(this.numericUpDownCopies)).EndInit();
			this.ResumeLayout(false);
			this.PerformLayout();

		}

		#endregion

		private System.Windows.Forms.Label labelName;
		private System.Windows.Forms.ComboBox comboBoxName;
		private System.Windows.Forms.Label labelCopies;
		private System.Windows.Forms.NumericUpDown numericUpDownCopies;
		private System.Windows.Forms.Label labelFirstPageBin;
		private System.Windows.Forms.ComboBox comboBoxFirstPageBin;
		private System.Windows.Forms.CheckBox checkBoxUseDifferentBin;
		private System.Windows.Forms.Label labelOtherPagesBin;
		private System.Windows.Forms.ComboBox comboBoxOtherPagesBin;
		private System.Windows.Forms.Button buttonOK;
		private System.Windows.Forms.Button buttonCancel;
	}
}