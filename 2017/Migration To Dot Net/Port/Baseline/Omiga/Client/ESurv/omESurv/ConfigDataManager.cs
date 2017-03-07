/*
--------------------------------------------------------------------------------------------
Workfile:			ConfigDataManager.cs
Copyright:			Copyright © Vertex Financial Services Ltd 2005

Description:		
--------------------------------------------------------------------------------------------
History:

Prog	Date		Description
JD		28/10/2005	MAR342 Continued work
JD		09/11/2005	MAR488 error in getGlobalParameter
GHun	21/11/2005	MAR538 Removed GlobalParameter function as this is available in omCode.GlobalParameter
--------------------------------------------------------------------------------------------
*/
using System;
using System.Data;
using System.Data.SqlClient;
using System.Xml;
using System.Text;

using Vertex.Fsd.Omiga.Core;
using XMLManager = Vertex.Fsd.Omiga.omESurv.XML.XMLManager;

namespace Vertex.Fsd.Omiga.omESurv.ConfigData
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
			
			if (xnTempNode  == null)
				return "";
			else
				return _xmlMgr.xmlGetAttributeText(xnTempNode,"VALUEID");
		}

		public string GetComboValidationTypeForValueID(string vstrGroupName, int ValueID) 
		{
			SqlCommand oComm;
			SqlParameter oParam;
			StringBuilder Response = new StringBuilder("<RESPONSE TYPE='SUCCESS'>");
			string valType;

			try 
			{
				string _ConnectionString = Global.DatabaseConnectionString;

				using (SqlConnection oConn = new SqlConnection(_ConnectionString))
				{
					oComm = new SqlCommand("USP_GETCOMBOVALIDATIONTYPEFORVALUEID", oConn);
					oComm.CommandType = CommandType.StoredProcedure;
					
					oParam = oComm.Parameters.Add("@p_GroupName", SqlDbType.NVarChar, vstrGroupName.Length);
					oParam.Value = vstrGroupName;

					oParam = oComm.Parameters.Add("@p_ValueID", SqlDbType.Int);
					oParam.Value = ValueID;

					oConn.Open();
				
					SqlDataReader oReader = oComm.ExecuteReader();
				
					do 
					{
						if (oReader.Read())
						{
							Response.Append(oReader.GetString(0));
						}
					} while (oReader.NextResult());

					Response.Append("</RESPONSE>");

					XmlDocument xdResponse = new XmlDocument();
					xdResponse.LoadXml(Response.ToString());
					XmlNode xnComboValidation = xdResponse.SelectSingleNode("/RESPONSE/COMBOVALIDATION");
					valType = xnComboValidation.Attributes.GetNamedItem("VALIDATIONTYPE").Value;
				}
			}
			catch (Exception ex)
			{
				throw new OmigaException(ex);
			}
			finally
			{
				oParam = null;
				oComm = null;
			}

			return valType;
		}
	
		public string GetComboGroup(string vstrGroupName, int ValueID) 
		{
			SqlCommand oComm;
			SqlParameter oParam;
			StringBuilder Response = new StringBuilder("<RESPONSE TYPE='SUCCESS'>");

			try 
			{
				string _ConnectionString = Global.DatabaseConnectionString;

				using (SqlConnection oConn = new SqlConnection(_ConnectionString))
				{
					oComm = new SqlCommand("USP_GETCOMBOGROUP", oConn);
					oComm.CommandType = CommandType.StoredProcedure;

					oParam = oComm.Parameters.Add("@p_GroupName", SqlDbType.NVarChar, vstrGroupName.Length);
					oParam.Value = vstrGroupName;
				
					oParam = oComm.Parameters.Add("@p_ValueID", SqlDbType.Int );
					if (ValueID.ToString() == "")
						oParam.Value = null;
					else
						oParam.Value = ValueID;

					oConn.Open();
				
					SqlDataReader oReader = oComm.ExecuteReader();
					do 
					{
						if (oReader.Read())
						{
							Response.Append(oReader.GetString(0));
						}
					} while (oReader.NextResult());

					Response.Append("</RESPONSE>");
				}
			}
			catch (Exception ex)
			{
				throw new OmigaException(ex);
			}
			finally
			{
				oComm = null;
				oParam = null;
			}
		
			return Response.ToString();
		}

		public Boolean IsValidationTypeForValueID (string GroupName,int ValueID, string ValidationType)
		{
			Boolean IsValidationType  = false;
			XmlDocument xdComboGroup = new XmlDocument();
			xdComboGroup.LoadXml(GetComboGroup(GroupName, ValueID));
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
			SqlCommand oComm;
			SqlParameter oParam;
			StringBuilder Response = new StringBuilder("<RESPONSE TYPE='SUCCESS'>");

			try 
			{
				string _ConnectionString = Global.DatabaseConnectionString;

				using (SqlConnection oConn = new SqlConnection(_ConnectionString))
				{
					oComm = new SqlCommand("USP_GETCOMBOGROUP", oConn);
					oComm.CommandType = CommandType.StoredProcedure;

					oParam = oComm.Parameters.Add("@p_GroupName", SqlDbType.NVarChar, vstrGroupName.Length);
					oParam.Value = vstrGroupName;
				
					oParam = oComm.Parameters.Add("@p_ValueID", SqlDbType.Int );
					oParam.Value = 0;
				
					oConn.Open();
				
					SqlDataReader oReader = oComm.ExecuteReader();
					do 
					{
						if (oReader.Read())
						{
							Response.Append(oReader.GetString(0));
						}
					} while (oReader.NextResult());

					Response.Append("</RESPONSE>");
				}
			}
			catch(Exception e)
			{
				throw new OmigaException("Error in Database Call", e);
			}
			finally
			{
				oComm = null;
				oParam = null;
			}
		
			return Response.ToString();
		}

		public string GetComboValueNameForValueID (string GroupName,int ValueID)
		{
		
			XmlDocument xdComboGroup = new XmlDocument();
			
			xdComboGroup.LoadXml(GetComboGroup(GroupName, ValueID));

			string Pattern = "/RESPONSE/COMBOGROUP[@GROUPNAME='" + GroupName + "']/COMBOVALUE[@VALUEID=" + ValueID + "]";
			
			XmlNode xnTempNode = xdComboGroup.SelectSingleNode(Pattern);
			
			if (xnTempNode  == null)
				return "";
			else
				return _xmlMgr.xmlGetAttributeText(xnTempNode,"VALUENAME");
		}

		public string GetComboValueIDForValueName (string GroupName, string ValueName)
		{
		
			XmlDocument xdComboGroup = new XmlDocument();
			
			xdComboGroup.LoadXml(GetComboGroup(GroupName));

			string Pattern = "/RESPONSE/COMBOGROUP[@GROUPNAME='" + GroupName + "']/COMBOVALUE[@VALUENAME='" + ValueName + "']";
			
			XmlNode xnTempNode = xdComboGroup.SelectSingleNode(Pattern);
			
			if (xnTempNode  == null)
				return "";
			else
				return _xmlMgr.xmlGetAttributeText(xnTempNode,"VALUEID");
		}
	}
}
