using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Text;
using System.Windows.Forms;

namespace Vertex.Fsd.Omiga.AxWord
{
	public partial class PrintDialog : Form
	{
		private PrintJob _printJob;
		public PrintJob PrintJob
		{
			get { return _printJob; }
			set { _printJob = value; }
		}

		private bool _showProgess;
		private PrinterCollection _printers;

		public PrintDialog(PrintJob printJob, bool showProgress)
		{
			if (printJob == null)
			{
				throw new ArgumentNullException("printJob");
			}

			_printJob = printJob;
			_showProgess = showProgress;

			InitializeComponent();
		}

		private void PrintDialog_Shown(object sender, EventArgs e)
		{
			Enabled = false;
			_printers = new PrinterCollection(new RunWorkerCompletedEventHandler(backgroundWorker_RunWorkerCompleted), _showProgess);
		}

		private void backgroundWorker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
		{
			Cursor = Cursors.WaitCursor;

			try
			{
				comboBoxName.Items.Clear();
				string defaultPrinterName = null;
				foreach (Printer printer in _printers)
				{
					comboBoxName.Items.Add(printer);
					if (printer.IsDefault)
					{
						defaultPrinterName = printer.Name;
					}
				}

				if (_printJob.FirstPagePrinterBin.Number == -1)
				{
					// Not set in caller, so use default bin.
					_printJob.FirstPagePrinterBin.Number = MicrosoftWordConstants.wdPrinterDefaultBin;
				}

				if (_printJob.OtherPagesPrinterBin.Number == -1)
				{
					// Not set in caller, so use default bin.
					_printJob.OtherPagesPrinterBin.Number = MicrosoftWordConstants.wdPrinterDefaultBin;
				}

				if (!string.IsNullOrEmpty(_printJob.Printer.Name))
				{
					comboBoxName.SelectedIndex = comboBoxName.FindString(_printJob.Printer.Name);
				}
				else if (!string.IsNullOrEmpty(defaultPrinterName))
				{
					comboBoxName.SelectedIndex = comboBoxName.FindString(defaultPrinterName);
				}

				numericUpDownCopies.Value = _printJob.Copies;

				checkBoxUseDifferentBin.Checked = _printJob.UseDifferentBinForOtherPages;

				Enabled = true;

				Activate();
			}
			finally
			{
				Cursor = Cursors.Default;
			}
		}

		private void comboBoxName_SelectedIndexChanged(object sender, EventArgs e)
		{
			if (comboBoxName.SelectedIndex != -1)
			{
				Cursor = Cursors.WaitCursor;

				try
				{
					Printer printer = (Printer)comboBoxName.Items[comboBoxName.SelectedIndex];
					if (printer != null)
					{
						comboBoxFirstPageBin.Items.Clear();
						comboBoxOtherPagesBin.Items.Clear();

						int selectedFirstPageBinIndex = -1;
						int selectedOtherPagesBinIndex = -1;
						foreach (PrinterBin printerBin in printer.Bins)
						{
							int firstPageBinIndex = comboBoxFirstPageBin.Items.Add(printerBin);
							if ((_printJob.FirstPagePrinterBin.Number != -1 && printerBin.Number == _printJob.FirstPagePrinterBin.Number) || printerBin.Number == printer.DefaultBin.Number)
							{
								selectedFirstPageBinIndex = firstPageBinIndex;
							}
							int otherPagesBinIndex = comboBoxOtherPagesBin.Items.Add(printerBin);
							if ((_printJob.OtherPagesPrinterBin.Number != -1 && printerBin.Number == _printJob.OtherPagesPrinterBin.Number) || printerBin.Number == printer.DefaultBin.Number)
							{
								selectedOtherPagesBinIndex = otherPagesBinIndex;
							}
						}

						if (selectedFirstPageBinIndex != -1 && selectedFirstPageBinIndex < comboBoxFirstPageBin.Items.Count)
						{
							comboBoxFirstPageBin.SelectedIndex = selectedFirstPageBinIndex;
							_printJob.FirstPagePrinterBin = (PrinterBin)comboBoxFirstPageBin.Items[selectedFirstPageBinIndex];
						}
						else
						{
							if (comboBoxFirstPageBin.Items.Count > 0)
							{
								// First item is always Default tray.
								comboBoxFirstPageBin.SelectedIndex = 0;
							}
							_printJob.FirstPagePrinterBin = ((Printer)comboBoxName.SelectedItem).DefaultBin;
						}

						if (selectedOtherPagesBinIndex != -1 && selectedOtherPagesBinIndex < comboBoxOtherPagesBin.Items.Count)
						{
							comboBoxOtherPagesBin.SelectedIndex = selectedOtherPagesBinIndex;
							_printJob.OtherPagesPrinterBin = (PrinterBin)comboBoxOtherPagesBin.Items[selectedOtherPagesBinIndex];
						}
						else
						{
							if (comboBoxOtherPagesBin.Items.Count > 0)
							{
								// First item is always Default tray.
								comboBoxOtherPagesBin.SelectedIndex = 0;
							}
							_printJob.OtherPagesPrinterBin = ((Printer)comboBoxName.SelectedItem).DefaultBin;
						}
					}
				}
				finally
				{
					Cursor = Cursors.Default;
				}
			}

			labelCopies.Enabled = numericUpDownCopies.Enabled = comboBoxName.SelectedIndex != -1;
			labelFirstPageBin.Enabled = comboBoxFirstPageBin.Enabled = comboBoxName.SelectedIndex != -1;
			checkBoxUseDifferentBin.Enabled = comboBoxName.SelectedIndex != -1;
			checkBoxUseDifferentBin_CheckedChanged(null, null);
			buttonOK.Enabled = comboBoxName.SelectedIndex != -1;
		}

		private void checkBoxUseDifferentBin_CheckedChanged(object sender, EventArgs e)
		{
			labelOtherPagesBin.Enabled = comboBoxOtherPagesBin.Enabled = comboBoxName.SelectedIndex != -1 && checkBoxUseDifferentBin.Checked;
		}

		private void buttonOK_Click(object sender, EventArgs e)
		{
			_printJob.Printer.Name = ((Printer)comboBoxName.SelectedItem).Name;
			_printJob.Copies = (int)numericUpDownCopies.Value;
			_printJob.FirstPagePrinterBin = (PrinterBin)comboBoxFirstPageBin.SelectedItem;
			_printJob.OtherPagesPrinterBin = (PrinterBin)comboBoxOtherPagesBin.SelectedItem;
			_printJob.UseDifferentBinForOtherPages = checkBoxUseDifferentBin.Checked;
		}

	}
}