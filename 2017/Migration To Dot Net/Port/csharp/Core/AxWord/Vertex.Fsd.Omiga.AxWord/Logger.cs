/*
-------------------------------------------------------------------------------------------
Workfile:			Logger.cs
Copyright:			Copyright © 2007 Marlborough Stirling
Description:		Logger to file.
--------------------------------------------------------------------------------------------
History:

Prog	Date		Description
AS		12/09/2007	First version. 
--------------------------------------------------------------------------------------------
*/
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Reflection;
using System.Text;
using System.Threading;
using System.Xml;

namespace Vertex.Fsd.Omiga.AxWord
{
    public class Logger : IDisposable
    {
		private static Logger _logger = new Logger();
		public static Logger GetLogger()
		{
			return _logger;
		}

		private string _fileName = "";
		public string FileName
		{
			get { return _fileName; }
			set { _fileName = value; }
		}

		private bool _enabled;
		public bool Enabled
		{
			get { return _enabled; }
			set { _enabled = value; }
		}

		private StreamWriter _streamWriter;
		private int _xmlFileIndex;
		private bool _disposed;

		private Logger()
        {
        }

        ~Logger() 
        {
            Dispose(false);
        }

		private bool Open()
		{
			return Open(null);
		}

		private bool Open(string fileName) 
        {
            bool opened = false;

			try 
            {
                if (_streamWriter == null && _enabled)
                {
                    if (!string.IsNullOrEmpty(fileName))
                    {
                        _fileName = fileName;
                    }

                    if (string.IsNullOrEmpty(_fileName))
                    {
                        _fileName = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location) + Path.DirectorySeparatorChar + Path.GetFileNameWithoutExtension(System.Reflection.Assembly.GetExecutingAssembly().Location) + ".log";
                    }

					bool fileExists = File.Exists(_fileName);

                    _streamWriter = new StreamWriter(_fileName, true);
                    if (!fileExists)
                    {
                        WriteAbout();
                    }

                    opened = true;
                }
                else
                {
                    // Logger is already open; do not open again.
                    opened = true;
                }
            }
            catch (Exception exception)
            {
                AxWordManager.WriteToEventLog(exception, "Opening file " + _fileName);
            }

            return opened;
        }

		private bool Close() 
        {
            bool closed = false;

			try 
            {
                if (_enabled)
                {
                    if (_streamWriter != null)
                    {
                        _streamWriter.Close();
                    }
               }

               closed = true;
            }
            catch (Exception exception)
            {
                AxWordManager.WriteToEventLog(exception, "Closing file " + _fileName);
            }

			return closed;
        }

		private void WriteAbout()
		{
			Assembly assembly = Assembly.GetCallingAssembly();
			FileVersionInfo fileVersionInfo = FileVersionInfo.GetVersionInfo(assembly.Location);
			Version version = assembly.GetName().Version;
			WriteLine("Description: " + assembly.FullName);
			WriteLine("Location: " + assembly.Location);
		}

		#region IDisposable Members

		public void Dispose()
		{
			Dispose(true);
			GC.SuppressFinalize(this);
		}

		protected virtual void Dispose(bool disposing)
		{
			if (!_disposed)
			{
				if (disposing)
				{
					// Dispose managed resources.
				}

				Close();
			}

			_disposed = true;
		}

		#endregion

        public bool WriteLine(string line) 
        {
			bool success = false;

			try
			{
				if (_enabled && Open())
				{
					_streamWriter.WriteLine(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff") + " 0x" + System.Threading.Thread.CurrentThread.ManagedThreadId.ToString("x") + ": " + line);
					_streamWriter.Flush();
				}

				success = true;
			}
			catch(Exception exception)
			{
				AxWordManager.WriteToEventLog(exception, "Writing to file " + _fileName);
			}

			return success;
		}

		public bool WriteLineXml(string line, string xml)
		{
			bool success = false;

		    try
			{
				if (_enabled && Open())
				{
					WriteLine(line + "(" + WriteXml(xml) + ")");
				}
			    
				success = true;
			}
			catch(Exception exception)
			{
				AxWordManager.WriteToEventLog(exception, "Writing to file " + _fileName);
			}

			return success;
		}

		public string WriteXml(string xml)
		{
			string xmlFileName = "";

			try
			{
				XmlDocument xmlDoc = null;
			    
				if (_enabled)
				{
					xmlDoc = new XmlDocument();
					xmlDoc.LoadXml(xml);
			        
					xmlFileName = Path.GetFileNameWithoutExtension(_fileName) + "." + _xmlFileIndex.ToString() + ".xml";			        
					_xmlFileIndex++;			        
					xmlDoc.Save(xmlFileName);
				}
			}
			catch(Exception exception)
			{
				AxWordManager.WriteToEventLog(exception, "Writing to file " + xmlFileName);
			}

			return xmlFileName;
		}			                        			   
	}
}
