using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using PreMigrationConverter.Properties;
using System.IO;
using System.Threading;
using System.Diagnostics;

namespace PreMigrationConverter
{
    public partial class frmMain : Form
    {
        /// <summary>
        /// Principal form 
        /// </summary>
        public frmMain()
        {
            InitializeComponent();
        }

        enum Status
        {
            /// <summary>
            /// Conversion in progress
            /// </summary>
            Started,
            /// <summary>
            /// Conversion not in progress
            /// </summary>
            Stopped
        }

        private Status _status;
        private Converter _converter;
        private int _filesAmended;
        private int _filesUnchanged;

        private void btnSelectInputFolder_Click(object sender, EventArgs e)
        {
            folderBrowserDialog1.SelectedPath = txtInputFolder.Text;
            if (folderBrowserDialog1.ShowDialog() == DialogResult.OK)
            {
                txtInputFolder.Text = folderBrowserDialog1.SelectedPath;
            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            txtDefectId.Text = Settings.Default.DefectId;
            rbtnRemoveCode.Checked = Settings.Default.DeleteUnwantedCode;
            txtInputFolder.Text = Settings.Default.InputPath;
            txtLogFile.Text = Settings.Default.LogFileName;
            _status = Status.Stopped;
            _converter = new Converter();
            _converter.FinishedConvertingFile += 
                new Converter.ConverterEventHandler(_converter_FinishedConvertingFile);
			toolTip.SetToolTip(btnStartStop, "Start processing files.");
			toolTip.SetToolTip(txtDefectId, "This defect id will appear in a comment inserted in the code.");
        }

        void _converter_FinishedConvertingFile(object sender, ConverterEventArgs e)
        {
            UpdateStatusStripText(e.FileName);
            if (e.FileAmended)
            {
                _filesAmended++;
            }
            else
            {
                _filesUnchanged++;
            }
        }

        delegate void UpdateStatusStripCallback(string text);

        /// <summary>
        /// Update the status strip
        /// </summary>
        /// <param name="text"></param>
        private void UpdateStatusStripText(string text)
        {
            if (statusStrip1.InvokeRequired)
            {
                UpdateStatusStripCallback del = 
                    new UpdateStatusStripCallback(UpdateStatusStripText);
                if (this != null)
                {
                    this.Invoke(del, new object[] { text });
                }
            }
            else
            {
                toolStripStatusLabel1.Text = text;
            }
        }

        delegate void UpdateStartStopButtonCallback(string text);

        private void UpdateStartStopButtonText(string text)
        {
            if (btnStartStop.InvokeRequired)
            {
                UpdateStartStopButtonCallback del =
                    new UpdateStartStopButtonCallback(UpdateStartStopButtonText);
                if (this != null)
                {
                    this.Invoke(del, new object[] { text });
                }
            }
            else
            {
                btnStartStop.Text = text;
            }
        }

        private void btnExit_Click(object sender, EventArgs e)
        {
            UpdateSettings();
            this.Close();
        }

        private void UpdateSettings()
        {
            Settings.Default.InputPath = txtInputFolder.Text;
            Settings.Default.LogFileName = txtLogFile.Text;
            Settings.Default.DeleteUnwantedCode = rbtnRemoveCode.Checked;
            Settings.Default.DefectId = txtDefectId.Text;
            Settings.Default.OverwriteFiles = chkOverwrite.Checked; //GHun
            Settings.Default.LogVerbose = chkVerboseLogging.Checked; // RF
            Settings.Default.Save();
        }

        public delegate void ConverterFinishedCallback(bool success);

        public void ReportConverterFinished(bool success)
        {
            //if (success == null)
            //{
            //    throw new ArgumentNullException();
            //}

            if (success)
            {
                toolStripStatusLabel1.Text = "Finished";
                MessageBox.Show("Finished. Files amended = " + _filesAmended.ToString() +
                    "; files unchanged = " + _filesUnchanged.ToString() +
                    "; total = " + (_filesAmended + _filesUnchanged).ToString());
                UpdateStartStopButtonText("Start");
                _status = Status.Stopped;
            }
            else
            {
                MessageBox.Show("See log file for error details", "Error");
            }
        }

        /// <summary>
        /// Start conversion
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnStart_Click(object sender, EventArgs e)
        {
			if (_status == Status.Stopped)
			{
				try
				{
					_status = Status.Started;
					UpdateStartStopButtonText("Stop");
					toolStripStatusLabel1.Text = "Started";
					_filesAmended = 0;
					_filesUnchanged = 0;
					// update app settings as they will be checked in FileConverter
					UpdateSettings();

					ConverterFinishedCallback cb = new ConverterFinishedCallback(
						ReportConverterFinished);
					_converter.Initialise(txtInputFolder.Text, txtLogFile.Text, cb);

					ThreadStart start = new ThreadStart(_converter.Start);
					Thread workerThread = new Thread(start);
					workerThread.Start();
				}
				catch (Exception exception)
				{
					MessageBox.Show(exception.ToString(), "Error", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					UpdateStartStopButtonText("Start");
					_status = Status.Stopped;
				}
			}
			else
			{
				UpdateStartStopButtonText("Start");
				_status = Status.Stopped;
			}
        }

        private void txtSelectLogFile_Click(object sender, EventArgs e)
        {
			try
			{
				openFileDialog1.FileName = txtLogFile.Text;
				if (openFileDialog1.ShowDialog() == DialogResult.OK)
				{
					txtLogFile.Text = openFileDialog1.FileName; 
				}            
			}
			catch (Exception exception)
			{
				MessageBox.Show(exception.ToString(), "Error", MessageBoxButtons.OK, MessageBoxIcon.Stop);
			}
        }

        private void btnViewLogFile_Click(object sender, EventArgs e)
        {
			try
			{
				// AS 26/11/2007 Use Windows file associations to launch the log file.
				// Thus, temp.csv will be opened in Excel.
				Process newProc = Process.Start(txtLogFile.Text);
				// AS 26/11/2007 newProc will be null when using Windows file associations.
				if (newProc != null)
				{
					// AS 26/11/2007 Do not wait for process to exit.
					// newProc.WaitForExit();
				}
			}
			catch (Exception exception)
			{
				MessageBox.Show(exception.ToString(), "Error", MessageBoxButtons.OK, MessageBoxIcon.Stop);
			}
        }

		private void btnRemoveComments_Click(object sender, EventArgs e)
		{

		}
    }
}