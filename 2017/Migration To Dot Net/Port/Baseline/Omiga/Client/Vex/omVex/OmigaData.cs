// --------------------------------------------------------------------------------------------
// Workfile:			OmigaData.cs
// Copyright:			Copyright ©2006 Vertex Financial Services
// 
// Description:		
// --------------------------------------------------------------------------------------------
// History:
// 
// Prog		Date		Description
// PE		08/11/2006	EP2_24 - Created
// PE		06/12/2006	EP2_280 - Unit test
// PB		03/01/2007	EP2_528 - WP8 - Xit2 Interface - Error codes as spec'd to be added into OmVEx
// PE		04/01/2007	EP2_449	- WP8 - Xit2 Interface - new SP to purge VexInboundMessage table
// --------------------------------------------------------------------------------------------
using System;
using System.IO;
using System.Xml;
using System.Xml.Xsl;
using System.Data;
using System.Data.SqlClient;
using System.Reflection;

using Vertex.Fsd.Omiga.Core;

namespace Vertex.Fsd.Omiga.omVex
{

	public class OmigaDataClass
	{
		LoggingClass _Logger = new LoggingClass(MethodBase.GetCurrentMethod().DeclaringType.ToString());

		public bool GetVexInstruction(OmigaMessageClass Request, ref XmlDocument VexInstruction)
		{
			string MethodName = MethodInfo.GetCurrentMethod().Name.ToString();
			_Logger.LogStart(MethodName);
			const string VexStoredProcedure = "usp_GetVexInterfaceRequest";
			const string VexXSLTFile = "VexInstruction.xslt";

			bool ProcessOK = false;
			SqlConnection OmigaConnection;
			XmlReader AppDataXMLReader;
			XmlDocument AppDataXMLDOM;
			SqlCommand GetAppDataCommand;

			try
			{
				OmigaConnection= new SqlConnection(Global.DatabaseConnectionString);
				GetAppDataCommand = new SqlCommand(VexStoredProcedure, OmigaConnection);
				GetAppDataCommand.CommandType = CommandType.StoredProcedure;

				SqlParameter AppNoParameter = GetAppDataCommand.Parameters.Add("@applicationnumber", SqlDbType.NVarChar, 12);
				AppNoParameter.Value = Request.Application_ApplicationNumber;

				SqlParameter FactFindNoParameter = GetAppDataCommand.Parameters.Add("@applicationfactfindnumber", SqlDbType.Int);
				FactFindNoParameter.Value = Request.Application_ApplicationFactFindNumber;

				OmigaConnection.Open();
				AppDataXMLReader = GetAppDataCommand.ExecuteXmlReader();
							
				ProcessOK = true;
			}
			catch(Exception NewError)
			{
				_Logger.LogError(MethodName, NewError.Message);
				throw new OmigaException(8801, NewError.Message);					
			}

			if(ProcessOK)
			{
				AppDataXMLDOM = new XmlDocument();					
				AppDataXMLDOM.Load(AppDataXMLReader);					

				StringWriter AppDataWriter = new StringWriter();

				XslTransform Transformer = new XslTransform();
				Transformer.Load(XmlHelper.GetXML(VexXSLTFile));
				Transformer.Transform(AppDataXMLDOM.FirstChild, null, AppDataWriter, null);

				VexInstruction = new XmlDocument();
				VexInstruction.LoadXml(AppDataWriter.ToString());

				_Logger.Debug("VexInstruction = " + VexInstruction.InnerXml);
			}
	
			_Logger.LogFinsh(MethodName);
			return ProcessOK;
		}

		public bool GetVexStatus(OmigaMessageClass Request, ref XmlDocument VexStatus)
		{
			string MethodName = MethodInfo.GetCurrentMethod().Name.ToString();
			_Logger.LogStart(MethodName);
			const string VexStoredProcedure = "usp_GetVexStatusRequest";
			const string VexXSLTFile = "VexStatus.xslt";

			bool ProcessOK = false;
			SqlConnection OmigaConnection;
			XmlReader AppDataXMLReader;
			XmlDocument AppDataXMLDOM;
			SqlCommand GetAppDataCommand;

			try
			{
				OmigaConnection= new SqlConnection(Global.DatabaseConnectionString);
				GetAppDataCommand = new SqlCommand(VexStoredProcedure, OmigaConnection);
				GetAppDataCommand.CommandType = CommandType.StoredProcedure;

				SqlParameter AppNoParameter = GetAppDataCommand.Parameters.Add("@applicationnumber", SqlDbType.NVarChar, 12);
				AppNoParameter.Value = Request.Application_ApplicationNumber;

				SqlParameter FactFindNoParameter = GetAppDataCommand.Parameters.Add("@applicationfactfindnumber", SqlDbType.Int);
				FactFindNoParameter.Value = Request.Application_ApplicationFactFindNumber;

				OmigaConnection.Open();
				AppDataXMLReader = GetAppDataCommand.ExecuteXmlReader();
							
				ProcessOK = true;
			}
			catch(Exception NewError)
			{
				_Logger.LogError(MethodName, NewError.Message);
				throw new OmigaException(8802, NewError.Message); //PB 03/01/2007 EP2_528
			}

			if(ProcessOK)
			{
				try
				{
					AppDataXMLDOM = new XmlDocument();					
					AppDataXMLDOM.Load(AppDataXMLReader);					

					StringWriter AppDataWriter = new StringWriter();

					XslTransform Transformer = new XslTransform();
					Transformer.Load(XmlHelper.GetXML(VexXSLTFile));
					Transformer.Transform(AppDataXMLDOM.FirstChild, null, AppDataWriter, null);

					VexStatus = new XmlDocument();
					VexStatus.LoadXml(AppDataWriter.ToString());

					_Logger.Debug("VexInstruction = " + VexStatus.InnerXml);
				}
				catch(Exception NewError)
				{
					_Logger.LogError(MethodName, NewError.Message);
				}
			}
	
			_Logger.LogFinsh(MethodName);
			return ProcessOK;
		}

		public bool StoreVexInboundMessage(string Message, string ResponseCode)
		{
			string MethodName = MethodInfo.GetCurrentMethod().Name.ToString();
			_Logger.LogStart(MethodName);
			const string VexStoredProcedure = "usp_InsertVexInboundMessage";
			string ApplicationNumber = "";
			bool ProcessOK = false;
			XmlDocument MessageDocument = new XmlDocument();
			SqlConnection OmigaConnection;
			SqlCommand StoreVexMessageCommand;

			try
			{
				MessageDocument.LoadXml(Message);
				ApplicationNumber = MessageDocument.SelectSingleNode("/Instruction/InstructionRef").InnerText;
				ApplicationNumber = ApplicationNumber.Substring(0, ApplicationNumber.Length-2);
				ProcessOK = true;
			}
			catch
			{
				if(ApplicationNumber != "")
				{
					ProcessOK = true;
				}
			}

			if(ProcessOK)
			{
				try
				{
					OmigaConnection= new SqlConnection(Global.DatabaseConnectionString);
					StoreVexMessageCommand = new SqlCommand(VexStoredProcedure, OmigaConnection);
					StoreVexMessageCommand.CommandType = CommandType.StoredProcedure;

					SqlParameter AppNoParameter = StoreVexMessageCommand.Parameters.Add("@applicationnumber", SqlDbType.NVarChar, 12);
					AppNoParameter.Value = ApplicationNumber;

					SqlParameter MessageParameter = StoreVexMessageCommand.Parameters.Add("@vexmessagexml", SqlDbType.NVarChar, 4000);
					MessageParameter.Value = Message;

					SqlParameter ResponseCodeParamater = StoreVexMessageCommand.Parameters.Add("@responsecodetovex", SqlDbType.Int);
					ResponseCodeParamater.Value = int.Parse(ResponseCode);

					OmigaConnection.Open();
					StoreVexMessageCommand.ExecuteNonQuery();							
				}
				catch(Exception NewError)
				{
					_Logger.LogError(MethodName, NewError.Message);
					throw new OmigaException(8801, NewError.Message);					
				}
			}

			_Logger.LogFinsh(MethodName + "(" + ProcessOK.ToString() + ")");
			return(ProcessOK);
		}

		public bool GetComboValue(string GroupName, string ValidationType, out string ValueName, out string ValueID)
		{
			string MethodName = MethodInfo.GetCurrentMethod().Name.ToString();
			_Logger.LogStart(MethodName);
			const string VexStoredProcedure = "usp_GetComboValueForValidationType";
			
			bool ProcessOK = false;
			SqlConnection OmigaConnection;
			SqlDataReader Reader;
			SqlCommand GetAppDataCommand;

			ValueName = "";
			ValueID = "";

			try
			{
				OmigaConnection= new SqlConnection(Global.DatabaseConnectionString);
				GetAppDataCommand = new SqlCommand(VexStoredProcedure, OmigaConnection);
				GetAppDataCommand.CommandType = CommandType.StoredProcedure;

				SqlParameter GroupNameParameter = GetAppDataCommand.Parameters.Add("@p_sGroupName", SqlDbType.NVarChar, 30);
				GroupNameParameter.Value = GroupName;

				SqlParameter ValidationTypeParameter = GetAppDataCommand.Parameters.Add("@p_sValidationType", SqlDbType.NChar, 20);
				ValidationTypeParameter.Value = ValidationType;
				
				OmigaConnection.Open();
				Reader = GetAppDataCommand.ExecuteReader();
				if(Reader.Read())
				{
					ValueName = Reader.GetValue(0).ToString();
					ValueID = Reader.GetValue(1).ToString();
				}
				ProcessOK = true;
			}
			catch(Exception NewError)
			{
				_Logger.LogError(MethodName, NewError.Message);
				throw new OmigaException(8801, NewError.Message);					
			}		

			_Logger.LogFinsh(MethodName);
			return(ProcessOK);
		}
		
		//USP_GETCOMBOS
		public bool GetComboIDByValue(string GroupName, string ValueName, out string ValueID)
		{
			string MethodName = MethodInfo.GetCurrentMethod().Name.ToString();
			_Logger.LogStart(MethodName);
			const string VexStoredProcedure = "usp_GetComboValueForValueName";
			
			bool ProcessOK = false;
			SqlConnection OmigaConnection;
			SqlDataReader Reader;
			SqlCommand GetAppDataCommand;

			ValueID = "";

			try
			{
				OmigaConnection= new SqlConnection(Global.DatabaseConnectionString);
				GetAppDataCommand = new SqlCommand(VexStoredProcedure, OmigaConnection);
				GetAppDataCommand.CommandType = CommandType.StoredProcedure;

				SqlParameter GroupNameParameter = GetAppDataCommand.Parameters.Add("@p_sGroupName", SqlDbType.NVarChar, 30);
				GroupNameParameter.Value = GroupName;

				SqlParameter ValueNameParameter = GetAppDataCommand.Parameters.Add("@p_sValueName", SqlDbType.NChar, 50);
				ValueNameParameter.Value = ValueName;
				
				OmigaConnection.Open();
				Reader = GetAppDataCommand.ExecuteReader();
				if(Reader.Read())
				{
					ValueID = Reader.GetValue(0).ToString();
				}
				ProcessOK = true;
			}
			catch(Exception NewError)
			{
				_Logger.LogError(MethodName, NewError.Message);
				throw new OmigaException(8801, NewError.Message);					
			}		

			_Logger.LogFinsh(MethodName);
			return(ProcessOK);
		}

		// EP2_449 - Peter Edney - 04/01/2007
		public bool PurgeVexMessage(OmigaMessageClass Request)
		{
			string MethodName = MethodInfo.GetCurrentMethod().Name.ToString();
			_Logger.LogStart(MethodName);
			const string VexStoredProcedure = "Usp_PurgeVEXMessage";
			
			bool ProcessOK = false;
			SqlConnection OmigaConnection;
			SqlCommand GetAppDataCommand;

			try
			{
				OmigaConnection= new SqlConnection(Global.DatabaseConnectionString);
				GetAppDataCommand = new SqlCommand(VexStoredProcedure, OmigaConnection);
				GetAppDataCommand.CommandType = CommandType.StoredProcedure;

				SqlParameter AppNoParameter = GetAppDataCommand.Parameters.Add("@ApplicationNumber", SqlDbType.NVarChar, 12);
				AppNoParameter.Value = Request.Application_ApplicationNumber;

				SqlParameter ResponseCodeParameter = GetAppDataCommand.Parameters.Add("@ResponseCode", SqlDbType.Int);
				ResponseCodeParameter.Value = Request.Vex_ResponseCode;

				OmigaConnection.Open();
				GetAppDataCommand.ExecuteNonQuery();
				ProcessOK = true;
			}
			catch(Exception NewError)
			{
				_Logger.LogError(MethodName, NewError.Message);
				throw new OmigaException(8801, NewError.Message);					
			}		

			_Logger.LogFinsh(MethodName);
			return(ProcessOK);

		}

	} //class OmigaDataClass

}