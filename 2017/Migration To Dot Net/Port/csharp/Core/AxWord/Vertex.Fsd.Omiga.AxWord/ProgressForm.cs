using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Threading;
using System.Windows.Forms;

namespace Vertex.Fsd.Omiga.AxWord
{
	public interface IProgressWorker
	{
		void DoWork(BackgroundWorker backgroundWorker);
	}

	public partial class ProgressForm : Form
	{
		public string Message
		{
			get { return labelMessage.Text; }
			set	{ labelMessage.Text = value; }
		}

		public BackgroundWorker BackgroundWorker
		{
			get { return backgroundWorker; } 
		}

		private IProgressWorker _progressFormWorker;

		public ProgressForm(IProgressWorker progressFormWorker)
			: this(progressFormWorker, null, null)
		{
		}

		public ProgressForm(IProgressWorker progressFormWorker, string text, string message)
		{
			InitializeComponent();

			if (string.IsNullOrEmpty(text))
			{
				text = "Please wait...";
			}
			Text = text;
			
			if (string.IsNullOrEmpty(message))
			{
				message = "Please wait...";
			}
			labelMessage.Text = message;

			_progressFormWorker = progressFormWorker;
		}

		private void ProgressForm_Load(object sender, EventArgs e)
		{
			Cursor = Cursors.WaitCursor;
			Activate();
			backgroundWorker.RunWorkerAsync(_progressFormWorker);
		}

		private void ProgressForm_FormClosing(object sender, FormClosingEventArgs e)
		{
			Cursor = Cursors.Default;
		}

		private void backgroundWorker_DoWork(object sender, DoWorkEventArgs e)
		{
			((IProgressWorker)e.Argument).DoWork(backgroundWorker);
		}

		private void backgroundWorker_ProgressChanged(object sender, ProgressChangedEventArgs e)
		{
			progressBar.Value = e.ProgressPercentage;
		}

		private void backgroundWorker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
		{
			Close();
		}
	}
}