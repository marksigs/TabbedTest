using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing.Printing;
using System.IO;
using System.Text;
using System.Threading;
using System.Xml;

namespace Vertex.Fsd.Omiga.AxWord
{
	public class PrinterCollection : List<Printer>
	{
		private AutoResetEvent _initializeEvent = new AutoResetEvent(false);

		public PrinterCollection() 
		{
			Initialize(new RunWorkerCompletedEventHandler(backgroundWorker_RunWorkerCompleted), false);
			_initializeEvent.WaitOne();
		}

		public PrinterCollection(RunWorkerCompletedEventHandler runWorkerCompletedEventHandler, bool showProgress)
		{
			Initialize(runWorkerCompletedEventHandler, showProgress);
		}

		private void Initialize(RunWorkerCompletedEventHandler runWorkerCompletedEventHandler, bool showProgress)
		{
			if (runWorkerCompletedEventHandler == null)
			{
				throw new ArgumentNullException("runWorkerCompletedEventHandler");
			}

			InitializePrintersWorker initializePrintersWorker = new InitializePrintersWorker(this);
			if (showProgress)
			{
				ProgressForm progressForm = new ProgressForm(initializePrintersWorker, "Please wait...", "Getting printer information.");
				progressForm.BackgroundWorker.RunWorkerCompleted += runWorkerCompletedEventHandler;
				progressForm.Show();
			}
			else
			{
				initializePrintersWorker.DoWork(null);
				runWorkerCompletedEventHandler(null, null);
			}
		}

		private void backgroundWorker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
		{
			_initializeEvent.Set();
		}

		public string ToXmlString()
		{
			string xmlString = null;

			using (StringWriter stringWriter = new StringWriter())
			{
				XmlWriterSettings xmlWriterSettings = new XmlWriterSettings();
				xmlWriterSettings.OmitXmlDeclaration = true;
				using (XmlWriter xmlWriter = XmlWriter.Create(stringWriter, xmlWriterSettings))
				{
					xmlWriter.WriteStartDocument();
					xmlWriter.WriteStartElement("RESPONSE");
					xmlWriter.WriteStartElement("PRINTERS");

					foreach (Printer printer in this)
					{
						xmlWriter.WriteStartElement("PRINTER");
						xmlWriter.WriteAttributeString("DEVICENAME", printer.Name);
						xmlWriter.WriteAttributeString("DEFAULT", printer.IsDefault ? "TRUE" : "FALSE");
						xmlWriter.WriteAttributeString("BINS", printer.Bins.Count.ToString());
						xmlWriter.WriteAttributeString("DEFAULTBIN", printer.DefaultBin.Number.ToString());

						foreach (PrinterBin printerBin in printer.Bins)
						{
							xmlWriter.WriteStartElement("BIN");
							xmlWriter.WriteAttributeString("BINNUMBER", printerBin.Number.ToString());
							xmlWriter.WriteAttributeString("BINNAME", printerBin.Name);
							xmlWriter.WriteEndElement();
						}

						xmlWriter.WriteEndElement();
					}

					xmlWriter.WriteEndElement();
					xmlWriter.WriteEndElement();

					xmlWriter.WriteEndDocument();
					xmlWriter.Flush();
				}
				xmlString = stringWriter.ToString();
			}
			return xmlString;
		}

	}

	internal class InitializePrintersWorker : IProgressWorker
	{
		private List<Printer> _printers;

		public InitializePrintersWorker(List<Printer> printers)
		{
			_printers = printers;
		}

		#region IProgressWorker Members

		public void DoWork(System.ComponentModel.BackgroundWorker backgroundWorker)
		{
			PrinterSettings.StringCollection installedPrinters = PrinterSettings.InstalledPrinters;

			int currentProgress = 0;
			int progressIncrement = 100 / installedPrinters.Count;

			foreach (string installedPrinter in installedPrinters)
			{
				if (backgroundWorker == null || !backgroundWorker.CancellationPending)
				{
					Printer printer = new Printer(installedPrinter);
					PrinterSettings printerSettings = new PrinterSettings();
					printerSettings.PrinterName = printer.Name;
					printer.IsDefault = printerSettings.IsDefaultPrinter;
					_printers.Add(printer);

					currentProgress = currentProgress + progressIncrement < 100 ? currentProgress + progressIncrement : 100;
					if (backgroundWorker != null)
					{
						backgroundWorker.ReportProgress(currentProgress);
					}
#if DEBUG
					Thread.Sleep(500);
#endif
				}
			}
		}
		#endregion IProgressWorker Members

	}
}
