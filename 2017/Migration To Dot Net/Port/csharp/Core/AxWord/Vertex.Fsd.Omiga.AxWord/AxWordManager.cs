/*
-------------------------------------------------------------------------------------------
Workfile:			AxWord.cs
Copyright:			Copyright © 2007 Marlborough Stirling
Description:		Main AxWord type.
--------------------------------------------------------------------------------------------
History:

Prog	Date		Description
AS		15/08/2007	First version. Ported from axwordclass.cls.
--------------------------------------------------------------------------------------------
*/
using System;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.Diagnostics;
using System.IO;
using System.Reflection;
using System.Threading;
using System.Windows.Forms;
using System.Xml;

namespace Vertex.Fsd.Omiga.AxWord
{
	public class AxWordManager
	{
		#region Properties
		private static Logger _logger = Logger.GetLogger();

		private bool _persistState;
		public bool PersistState
		{
			get { return _persistState; }
			set { _persistState = value; }
		}

		private bool _modeless;
		public bool Modeless
		{
			get { return _modeless; }
			set { _modeless = value; }
		}

		private bool _readOnly;
		public bool ReadOnly
		{
			get { return _readOnly; }
			set { _readOnly = value; }
		}

		private bool _documentEdited;
		public bool DocumentEdited
		{
			get { return _documentEdited; }
		}

		public bool LogEnabled
		{
			get { return _logger.Enabled; }
			set { _logger.Enabled = value; }
		}

		public string LogFileName
		{
			get { return _logger.FileName; }
			set { _logger.FileName = value; }
		}

		private bool _showPrintDialog;
		public bool ShowPrintDialog
		{
			get { return _showPrintDialog; }
			set { _showPrintDialog = value; }
		}

		private bool _showProgressBar;
		public bool ShowProgressBar
		{
			get { return _showProgressBar; }
			set { _showProgressBar = value; }
		}

		private bool _disablePrintOut;
		public bool DisablePrintOut
		{
			get { return _disablePrintOut; }
			set { _disablePrintOut = value; }
		}

		private bool _documentPrinted;
		public bool DocumentPrinted
		{
			get { return _documentPrinted; }
		}
		#endregion Properties

		#region DisplayDocument
		public string DisplayDocument(AxWordRequest axWordRequest)
		{
			string fileContentsBinBase64 = "";

			if (axWordRequest == null)
			{
				throw new ArgumentNullException("axWordRequest");
			}

			_logger.WriteLine("->DisplayDocument()");

			try
			{
				string fileName = axWordRequest.Document.ToFile();

				if (_modeless)
				{
					// If mode less then we do not wait for the user to unlock the file, and
					// therefore we can not pick up any edits to the file, thus make the file
					// read only to prevent the user editing it.
					_readOnly = true;
				}

				if (axWordRequest.Document.DeliveryType != DocumentDeliveryType.Rtf && axWordRequest.Document.DeliveryType != DocumentDeliveryType.Word)
				{
					// Only RTF and Word documents can be edited.
					_readOnly = true;
				}

				DateTime oldFileDateTime = File.GetLastWriteTime(fileName);

				using (Process process = ShellExecute(fileName))
				{
					if (process != null)
					{
						if (!_modeless)
						{
							// If not modeless and not read only then we must wait to pick up 
							// any edits to the file, once the file has been closed.
							bool success = WaitForFileToClose(axWordRequest, fileName, process);

							if (success && !_readOnly)
							{
								// Not read only so check if file has been edited.
								DateTime newFileDateTime = File.GetLastWriteTime(fileName);

								_documentEdited = newFileDateTime != oldFileDateTime;
								if (_documentEdited)
								{
									// The file date has changed so assume it has been edited, 
									// so read it back in.
									axWordRequest.Document.FromFile();
								}
							}
						}

						fileContentsBinBase64 = axWordRequest.Document.FileContentsBinBase64;
					}
					else
					{
						throw new InvalidOperationException("Unable to execute application associated with " + fileName);
					}
				}
			}
			catch (Exception exception)
			{
				ReportException(exception);
				throw;
			}

			_logger.WriteLine("<-DisplayDocument()");

			return fileContentsBinBase64;
		}

		// Do not use Win API ShellExecuteEx as it returns a process handle, not the process id.
		// The process id is required to activate the process window in AppActivate.
		internal static Process ShellExecute(string fileName)
		{
			Process process = null;

			/*
			// The following is not necessary as Process.Start(fileName) 
			// seems to work for all file types.
			string executableFileName = null;

			// Find the executable for this document type.
			if (string.Compare(Path.GetExtension(fileName), ".afp", true) == 0)
			{
				// AFPs (IBM print file format) should be viewed in Internet Explorer using IBM's browser plug-in.
				string htmlFileName = Path.ChangeExtension(fileName, ".html");
				try
				{
					// File must exist for FindExecutable to work, so create a 
					// temporary HTML file.
					File.Copy(fileName, htmlFileName, true);
					executableFileName = NativeMethods.FindExecutable(htmlFileName);
				}
				finally
				{
					try
					{
						File.SetAttributes(htmlFileName, FileAttributes.Normal);
						File.Delete(htmlFileName);
					}
					catch (Exception exception)
					{
						AxWordManager.ReportException(exception);
					}
				}
			}
			else
			{
				executableFileName = NativeMethods.FindExecutable(fileName);
			}

			if (!string.IsNullOrEmpty(executableFileName))
			{
				// Start the shelled application.
				process = Process.Start(executableFileName, "\"" + fileName + "\"");
			}
			*/

			process = Process.Start(fileName);

			// Find the actual process that has the document open.
			// If Word is already running when the document is opened, then 
			// Process.Start will use the existing WinWord process, but the Process 
			// object returned by Process.Start will NOT be for this existing 
			// WinWord process, but for a temporary process object that will be 
			// immediately exited (e.g., HasExited == true).

			Process[] processes = Process.GetProcesses();
			string name = Path.GetFileName(fileName);
			foreach (Process p in processes)
			{
				if (p.MainWindowTitle.Contains(name))
				{
					process = p;
					break;
				}
			}

			return process;
		}

		private bool WaitForFileToClose(AxWordRequest axWordRequest, string fileName, Process process)
		{
			bool success = false;

			if (!process.HasExited && process.Responding)
			{
				process.WaitForInputIdle();

				int iteration = 0;
				bool loop = true;
				bool waitingForProcessToOpenFile = true;
				while (loop)
				{
					iteration++;

					bool isFileLockedByProcess = false;

					if (
							axWordRequest.Document.DeliveryType == DocumentDeliveryType.Rtf || 
							axWordRequest.Document.DeliveryType == DocumentDeliveryType.Word || 
							axWordRequest.Document.DeliveryType == DocumentDeliveryType.Pdf)
					{
						// Word (.doc and .rtf) and Acrobat (.pdf) obey file locking, 
						// so we can tell when the file is no longer being viewed (in Word or Acrobat) 
						// by using file locks.
						try
						{
							// Try to open the file for writing. This will try to take out a write lock, 
							// which will throw an IOException if the file is already locked by the external process.
							using (FileStream fileStream = File.OpenWrite(fileName))
							{
							}
						}
						catch (IOException)
						{
							isFileLockedByProcess = true;
						}
					}
					else
					{
						// For other file types, e.g., Internet Explorer (.afp) 
						// assume that if the process is running then the file is locked.
						isFileLockedByProcess = !process.HasExited;
					}

					if (isFileLockedByProcess)
					{
						// File is already open so no longer waiting for process to open the file.
						waitingForProcessToOpenFile = false;
					}
					else if (waitingForProcessToOpenFile)
					{
						// Waiting for the process to open the file.
						if (iteration > 120)
						{
							// Run out of time waiting for the process to open the file.
							loop = false;
						}
					}
					else
					{
						// The file is not locked and not waiting for the process to 
						// open it, so assume the process has finished with the file.
						loop = false;
						success = true;
					}

					if (loop)
					{
						Thread.Sleep(500);
					}
				}
			}
			else
			{
				throw new InvalidOperationException("Invalid process for executing '" + fileName + "'");
			}

			return success;
		}
		#endregion DisplayDocument

		#region PrintDocument
		public bool PrintDocument(AxWordRequest axWordRequest)
		{
			bool printed = false;

			if (axWordRequest == null)
			{
				throw new ArgumentNullException("axWordRequest");
			}

			_logger.WriteLine("->PrintDocument()");

			try
			{
				axWordRequest.Document.ToFile();

				bool success = true;
				if (_showPrintDialog)
				{
					success = false;
					PrintDialog printDialog = new PrintDialog(axWordRequest.PrintJob, _showProgressBar);
					if (printDialog.ShowDialog() == DialogResult.OK)
					{
						axWordRequest.PrintJob = printDialog.PrintJob;
						if (_persistState)
						{
							axWordRequest.PrintJob.SaveState();
						}
						success = true;
					}
				}

				if (success)
				{
					_documentPrinted = printed = axWordRequest.Document.Print(null, _disablePrintOut, axWordRequest.PrintJob);
				}
			}
			catch (Exception exception)
			{
				ReportException(exception);
				throw;
			}

			_logger.WriteLine("<-PrintDocument() = " + printed.ToString());

			return printed;
		}
		#endregion PrintDocument

		#region Conversion
		public string ConvertFileToBase64(string fileName, string compressionMethod, bool deleteFile)
		{
			string binBase64 = "";

			if (string.IsNullOrEmpty(fileName))
			{
				throw new ArgumentNullException("fileName");
			}

			_logger.WriteLine("->ConvertFileToBase64('" + fileName + "', " + compressionMethod + ", " + deleteFile.ToString() + ")");

			try
			{
				Document document = Document.CreateDocument(fileName, Document.ToDocumentCompressionMethod(compressionMethod));
				binBase64 = document.FromFile();
				if (deleteFile)
				{
					File.Delete(fileName);
				}
			}
			catch (Exception exception)
			{
				ReportException(exception);
				throw;
			}

			_logger.WriteLine("<-ConvertFileToBase64()");

			return binBase64;
		}

		public bool ConvertBase64ToFile(string binBase64, string fileName, string compressionMethod)
		{
			bool success = false;

			if (string.IsNullOrEmpty(binBase64))
			{
				throw new ArgumentNullException("binBase64");
			}
			if (string.IsNullOrEmpty(fileName))
			{
				throw new ArgumentNullException("fileName");
			}

			_logger.WriteLine("->ConvertBase64ToFile('" + fileName + "', " + compressionMethod);

			try
			{
				Document document = Document.CreateDocument(fileName, Document.ToDocumentCompressionMethod(compressionMethod), binBase64);
				document.DeleteFilesOnDisposal = false;
				document.ToFile();
				success = true;
			}
			catch (Exception exception)
			{
				ReportException(exception);
				throw;
			}

			_logger.WriteLine("<-ConvertBase64ToFile()");

			return success;
		}
		#endregion Conversion

		#region Error handling
		internal static void ReportException(Exception exception)
		{
			WriteToEventLog(exception);
			if (_logger.Enabled)
			{
				_logger.WriteLine(exception.ToString());
			}
		}

		internal static bool WriteToEventLog(Exception exception)
		{
			return WriteToEventLog(exception, null);
		}

		internal static bool WriteToEventLog(Exception exception, string message)
		{
			bool success = true;

			try
			{
				string source = "AxWord";
				if (!EventLog.SourceExists(source))
				{
					EventLog.CreateEventSource(source, "Application");
				}

				if (!string.IsNullOrEmpty(message))
				{
					message = message.Trim();
					if (!message.EndsWith("."))
					{
						message += ".";
					}
					message += " ";
				}

				EventLog eventLog = new EventLog();
				eventLog.Source = source;

				string threadId = Thread.CurrentThread.ManagedThreadId.ToString();
				string entry =
					"Error in " + Assembly.GetExecutingAssembly().FullName + ", Thread=" + threadId + "\n" +
					"Error Type: " + exception.GetType().ToString() + "\n" +
					"Error Description: " + (message != null ? message : "") + exception.Message + "\n" +
					"Error Source: " + exception.StackTrace;

				eventLog.WriteEntry(entry, EventLogEntryType.Error);
			}
			catch (Exception)
			{
				// No recovery possible.
				success = false;
			}

			return success;
		}
		#endregion Error handling

	}
}
