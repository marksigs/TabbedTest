/*
--------------------------------------------------------------------------------------------
Workfile:			Logging.cs
Copyright:			Copyright ©2006 Vertex Financial Services

Description:		
--------------------------------------------------------------------------------------------
History:

Prog	Date		Description
PE		27/11/2006	EP2_24 - Created
--------------------------------------------------------------------------------------------
*/
using System;
using System.Reflection;

namespace Vertex.Fsd.Omiga.omVex
{
	public class LoggingClass
	{

		private  omLogging.Logger _logger = null;

		public LoggingClass(string name)
		{
			_logger = omLogging.LogManager.GetLogger(name);
		}

		public void Debug(string Message)
		{
			if(_logger.IsDebugEnabled)
			{
				_logger.Debug(Message);
			}
		}

		public void LogStart(string Source)
		{
			if(_logger.IsDebugEnabled)
			{
				_logger.Debug("Starting:" + Source + "()");
			}
		}

		public void LogStart(string Source, string Parameter1)
		{
			if(_logger.IsDebugEnabled)
			{
				_logger.Debug("Starting:" + Source + "(" + Parameter1 + ")");
			}
		}

		public void LogStart(string Source, string Parameter1, string Parameter2)
		{
			if(_logger.IsDebugEnabled)
			{
				_logger.Debug("Starting:" + Source + "(" + Parameter1 + ", " + Parameter2 + ")");
			}
		}

		public void LogStart(string Source, string Parameter1, string Parameter2, string Parameter3)
		{
			if(_logger.IsDebugEnabled)
			{
				_logger.Debug("Starting:" + Source + "(" + Parameter1 + ", " + Parameter2 + ", " + Parameter3 + ")");
			}
		}

		public void LogFinsh(string Source)
		{
			if(_logger.IsDebugEnabled)
			{
				_logger.Debug("Finished:" + Source);
			}
		}

		public void LogError(string Source, String Message)
		{
			if(_logger.IsErrorEnabled)
			{
				_logger.Error(Source + " : " + Message);
			}
		}
		
		public bool ErrorCheck(string Source, OmigaMessageClass Response)
		{

			bool ProcessOK = true;

			if(Response.Result != "SUCCESS")
			{
				ProcessOK = false;			
				if(_logger.IsErrorEnabled)
				{					
					_logger.Error(Source + " : " + Response.Error_Description  + "(" + Response.Error_Source + ")");
				}				
			}

			return(ProcessOK);
		}

		public bool ErrorCheck(string Source, string Message, OmigaMessageClass Response)
		{

			bool ProcessOK = true;

			if(Response.Result != "SUCCESS")
			{
				ProcessOK = false;
				if(_logger.IsErrorEnabled)
				{
					_logger.Error(Source + " : " + Message + " : " + Response.Error_Description  + "(" + Response.Error_Source + ")");
				}				
			}

			return(ProcessOK);
		}
	}
}
