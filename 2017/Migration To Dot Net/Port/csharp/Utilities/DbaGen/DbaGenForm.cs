using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace DbaGen
{
	public partial class DbaGenForm : Form
	{
		public DbaGenForm()
			: this(null)
		{
		}

		public DbaGenForm(string root)
		{
			InitializeComponent();
			if (!string.IsNullOrEmpty(root))
			{
				textBoxRoot.Text = root;
				buttonGenerate_Click(null, null);
			}
		}

		private void DbaGenForm_Shown(object sender, EventArgs e)
		{
			textBoxRoot.Focus();
		}

		private void buttonGenerate_Click(object sender, EventArgs e)
		{
			textBoxResult.Text = DllBaseAddress.NewDllBaseAddress(textBoxRoot.Text).ToString(radioButtonFormatCSharp.Checked ? "x" : "H");
			if (checkBoxCopyToClipboard.Checked)
			{
				Clipboard.SetData(DataFormats.Text, (object)textBoxResult.Text);
			}
		}

		private void buttonExit_Click(object sender, EventArgs e)
		{
			Close();
		}
	}
}
