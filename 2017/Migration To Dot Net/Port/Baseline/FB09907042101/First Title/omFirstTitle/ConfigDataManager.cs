/*
--------------------------------------------------------------------------------------------
Workfile:			ConfigDataManager.cs
Copyright:			Copyright © Vertex Financial Services Ltd 2005

Description:		
--------------------------------------------------------------------------------------------
History:

Prog	Date		Description
MV		17/10/2005	MAR182	Created
GHun	09/11/2005	MAR182	Changed to use omCore to get db connection string and for exceptions
GHun	22/11/2005	MAR609	Removed GlobalParameter function as this is available in omCode.GlobalParameter
--------------------------------------------------------------------------------------------
*/
using System;
using System.Data;
using System.Data.SqlClient;
using System.Xml;
using System.Text;
using Vertex.Fsd.Omiga.Core;

using XMLManager = Vertex.Fsd.Omiga.omFirstTitle.XML.XMLManager;

namespace Vertex.Fsd.Omiga.omFirstTitle.ConfigData
{
	public class ConfigDataManager
	{
		XMLManager _xmlMgr = new XMLManager();
		
		public string GetComboValueIDForValidationType(string GroupName, string ValidationType)
		{
			XmlDocument xdComboGroup = new XmlDocument();
			
			xdComboGroup.LoadXml(GetComboGroup(GroupName));

			string Pattern = "/RESPONSE/COMBOGROUP/COMBOVALUE[COMBOVALIDATION[@VALIDATIONTYPE='" + ValidationType + "']]";
			
			XmlNode xnTempNode = xdComboGroup.SelectSingleNode(Pattern);
			
			if (xnTempNode == null)
				return "";
			else
				return _xmlMgr.xmlGetAttributeText(xnTempNode,"VALUEID");
		}

		public string GetComboValidationTypeForValueID(string vstrGroupName, int ValueID) 
		{
			string valType;

			try 
			{
				string _ConnectionString = Global.DatabaseConnectionString;

				using (SqlConnection oConn = new SqlConnection(_ConnectionString))
				{
					SqlCommand oComm = new SqlCommand("USP_GETCOMBOVALIDATIONTYPEFORVALUEID", oConn);
					oComm.CommandType = CommandType.StoredProcedure;
					
					SqlParameter oParam = oComm.Parameters.Add("@p_GroupName", SqlDbType.NVarChar, vstrGroupName.Length);
					oParam.Value = vstrGroupName;

					oParam = oComm.Parameters.Add("@p_ValueID", SqlDbType.Int);
					oParam.Value = ValueID;

					StringBuilder response = new StringBuilder("<RESPONSE TYPE='SUCCESS'>");

					oConn.Open();
				
					SqlDataReader oReader = oComm.ExecuteReader();
				
					do
					{
					while (oReader.Read())
					{
						response.Append(oReader.GetString(0));
					}
					} while (oReader.NextResult());

					response.Append("</RESPONSE>");

					XmlDocument xdResponse = new XmlDocument();
					xdResponse.LoadXml(response.ToString());
					XmlNode xnComboValidation = xdResponse.SelectSingleNode("/RESPONSE/COMBOVALIDATION");
					valType = xnComboValidation.Attributes.GetNamedItem("VALIDATIONTYPE").Value;
				}
			}
			catch (Exception ex)
			{
				throw new OmigaException(ex);
			}
			return valType;
		}
	
		public string GetComboGroup(string vstrGroupName, int ValueID) 
		{
			StringBuilder response = new StringBuilder("<RESPONSE TYPE='SUCCESS'>");

			try 
			{
				string _ConnectionString = Global.DatabaseConnectionString;

				using (SqlConnection oConn = new SqlConnection(_ConnectionString))
				{
					SqlCommand oComm = new SqlCommand("USP_GETCOMBOGROUP", oConn);
					oComm.CommandType = CommandType.StoredProcedure;

					SqlParameter oParam = oComm.Parameters.Add("@p_GroupName", SqlDbType.NVarChar, vstrGroupName.Length);
					oParam.Value = vstrGroupName;
				
					oParam = oComm.Parameters.Add("@p_ValueID", SqlDbType.Int);
					oParam.Value = ValueID;

					oConn.Open();
				
					SqlDataReader oReader = oComm.ExecuteReader();
					do
					{
						while (oReader.Read())
						{
							response.Append(oReader.GetString(0));
						}
					} while (oReader.NextResult());

					response.Append("</RESPONSE>");
				}
			}
			catch (Exception ex)
			{
				throw new OmigaException(ex);
			}
		
			return response.ToString();
		}

		public bool IsValidationTypeForValueID (string GroupName,int ValueID, string ValidationType)
		{
			bool IsValidationType  = false;
			XmlDocument xdComboGroup = new XmlDocument();
			xdComboGroup.LoadXml(GetComboGroup(GroupName,ValueID));
			string Pattern = "/RESPONSE/COMBOGROUP[@GROUPNAME='" + GroupName + "']/COMBOVALUE[@VALUEID=" + 
				            ValueID + "]/COMBOVALIDATION[@VALIDATIONTYPE='" + ValidationType + "']";

			if  (xdComboGroup.SelectSingleNode(Pattern) == null)
				IsValidationType = false;
			else
				IsValidationType = true;
			
			return IsValidationType;
		}

		public string GetComboGroup(string vstrGroupName) 
		{
			return GetComboGroup(vstrGroupName, 0);
		}
		
		public string GetComboValueNameForValueID(string GroupName, int ValueID)
		{
		
			XmlDocument xdComboGroup = new XmlDocument();
			
			xdComboGroup.LoadXml(GetComboGroup(GroupName,ValueID));

			string Pattern = "/RESPONSE/COMBOGROUP[@GROUPNAME='" + GroupName + "']/COMBOVALUE[@VALUEID=" + ValueID + "]";
			
			XmlNode xnTempNode = xdComboGroup.SelectSingleNode(Pattern);
			
			if  (xnTempNode  == null)
				return "";
			else
				return _xmlMgr.xmlGetAttributeText(xnTempNode,"VALUENAME");
		}

		public string GetComboValueIDForValueName (string GroupName,string ValueName)
		{
		
			XmlDocument xdComboGroup = new XmlDocument();
			
			xdComboGroup.LoadXml(GetComboGroup(GroupName));

			string Pattern = "/RESPONSE/COMBOGROUP[@GROUPNAME='" + GroupName + "']/COMBOVALUE[@VALUENAME='" + ValueName + "']";
			
			XmlNode xnTempNode = xdComboGroup.SelectSingleNode(Pattern);
			
			if  (xnTempNode  == null)
				return "";
			else
				return _xmlMgr.xmlGetAttributeText(xnTempNode, "VALUEID");
		}
	}
}
