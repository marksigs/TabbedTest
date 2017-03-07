/*
--------------------------------------------------------------------------------------------
Workfile:			DirectGatewayBO.cs
Copyright:			Copyright © Vertex Financial Services Ltd 2005

Description:		Direct Gateway Helper object, providing functions for setting common propeties
					required by the DirectGateway.
--------------------------------------------------------------------------------------------
History:

Prog	Date		Description
GHun	10/10/2005	MAR138 Created
JD		26/10/2995  MAR342 Added ESurvOutboundMSN
GHun	17/11/2005	MAR611 Changed GetCommunicationChannel response to upper case
GHun	12/12/2005	MAR852 Added GetOperator
PSC		29/03/2006	MAR1554 Amend GetNextSequenceNumber so it doesn't enlist in any
					encompassing transaction as it commits independently
GHun	18/04/2006	MAR1630 Changed GetNextSequenceNumber to fix null value check
--------------------------------------------------------------------------------------------
*/
using System;
using System.Data;
using System.Data.SqlClient;
using System.Net; // for proxy object
using System.IO;
using System.Xml;
using System.Xml.Serialization;
using System.Text;
using System.Runtime.InteropServices;
using Vertex.Fsd.Omiga.Core;	

namespace Vertex.Fsd.Omiga.omDg
{
	/// <summary>
	/// Helper functions for setting properties required for DirectGateway
	///

	/// </summary>
	[ProgId("omDG.DirectGatewayBO")]
	[Guid("679EFAF0-61E5-4046-9168-64B8DFB3F809")]
	[ComVisible(true)]
	public class DirectGatewayBO
	{
		private string _defaultUserId;
		
		public DirectGatewayBO()
		{
			_defaultUserId = String.Empty;
		}

		public string GetNextMessageSequenceNumber(string interfaceName)
		{
			//Validate interface name
			if (interfaceName == "Seq_ExpHunter" || 
				interfaceName == "Seq_ExpCreditCheck" ||
				interfaceName == "Seq_ExpKYC" || 
				interfaceName == "FirstTitleInboundMSN" ||
				interfaceName == "FirstTitleOutboundMSN" ||
				interfaceName == "ESurvOutboundMSN" )
			{
				return GetNextSequenceNumber(interfaceName);
			}
			else
				throw new OmigaException("Invalid interfaceName");
		}

		//Retrieve a sequence number from the SEQNEXTNUMBER table and update it
		private string GetNextSequenceNumber(string sequenceName)
		{
			string connString = Global.DatabaseConnectionString + "Enlist=false;";	// PSC 29/03/2006 MAR1554
			string nextNumber;

			using (SqlConnection conn = new SqlConnection(connString))
			{
				SqlCommand cmd = new SqlCommand("USP_GetNextSequenceNumber", conn);
				cmd.CommandType = CommandType.StoredProcedure;
				
				SqlParameter paramSName = cmd.Parameters.Add("@p_SequenceName", SqlDbType.NVarChar, sequenceName.Length);
				paramSName.Value = sequenceName;
				SqlParameter paramNextNumber = cmd.Parameters.Add("@p_NextNumber", SqlDbType.NVarChar, 12);
				paramNextNumber.Direction = ParameterDirection.Output;

				conn.Open();
				cmd.ExecuteNonQuery();

				if (paramNextNumber.Value != DBNull.Value)	//MAR1630 GHun
					nextNumber = (string) paramNextNumber.Value;
				else
					nextNumber = string.Empty;
			}

			return nextNumber;
		}

		private string GetOmigaPassword(string userId)
		{
			string password = String.Empty;
			string connString = Global.DatabaseConnectionString;
			SqlConnection conn = new SqlConnection(connString);
			SqlCommand cmd = null;
			SqlParameter spname = null;
			SqlDataReader sdr = null;
			try
			{
				cmd = new SqlCommand("SELECT TOP 1 PasswordValue FROM Password WHERE UserId = @UserId ORDER BY PASSWORDCREATIONDATE DESC", conn);
				cmd.CommandType = CommandType.Text;

				spname = new SqlParameter("@UserId", SqlDbType.NVarChar, userId.Length);
				spname.Value = userId;
				cmd.Parameters.Add(spname);

				conn.Open();
				
				sdr = cmd.ExecuteReader(CommandBehavior.SingleResult);
				if (sdr.HasRows)
				{
					sdr.Read();
					password = sdr.GetString(0);
				}
			}
			catch (Exception ex)
			{
				throw new OmigaException("Unable to retrieve password for user: " + userId, ex);
			}
			finally
			{
				conn.Close();
				conn.Dispose();
				conn = null;
				cmd.Dispose();
				cmd = null;
				spname = null;
				sdr = null;
			}
			
			return password;
		}

		public string GetClientDevice()
		{
			return "Omiga";
		}

		public string GetServiceName()
		{
			return "DirectGatewaySoapService";
		}

		public string GetServiceUrl()
		{
			GlobalParameter gp = new GlobalParameter("DirectGatewaySoapEndpoint");
			return gp.String;
		}

		public string GetProductType()
		{
			GlobalParameter gp = new GlobalParameter("ProfileMortgageProductCode");
			return gp.String;
		}
		
		public WebProxy GetProxy()
		{
			bool UseProxyServer = new GlobalParameter("UseProxyServerForDG").Boolean;
			if (UseProxyServer)
			{
				string proxyServer = new GlobalParameter("ProxyServer").String;
				WebProxy proxy = new WebProxy(proxyServer, true);

				//DefaultCredentials only work if the user running the code has proxy server access
				//proxy.Credentials = CredentialCache.DefaultCredentials;
				
				string user = new GlobalParameter("ProxyServerUser").String;
				string domain = new GlobalParameter("ProxyServerUserDomain").String;
				string password = new GlobalParameter("ProxyServerPassword").String;

				CredentialCache cc = new CredentialCache();
				cc.Add(proxy.Address, "Negotiate", new NetworkCredential(user, password, domain));
				cc.Add(proxy.Address, "Basic", new NetworkCredential(user, password, domain));
				proxy.Credentials = cc;

				return proxy;
			}
			else
				return null;
		}

		//Decrypt Omiga4 passwords
		private string Decrypt(string cypherText)
		{
			int inputLength = cypherText.Length;
			int charPosition = 0;
			string plainText = string.Empty;
			byte increment;
			byte char1;
			byte char2;
			bool isAdd = false;
			const byte asciiTilde = 126;

			byte[] random = new byte[30] {23, 12, 33, 40, 17, 31, 23, 25, 12, 43, 44, 43, 48, 43, 13, 35, 17, 23, 17, 37, 21, 34, 10, 38, 11, 12, 41, 21, 43, 23};
			byte randomIndex = 0;

			while (charPosition < inputLength)
			{
				char1 = (byte) cypherText.ToCharArray(charPosition, 1)[0];

				if (char1 == asciiTilde)
				{
					charPosition++;
					isAdd = false;
				}
				else
					isAdd = true;

				increment = random[randomIndex++];
				char1 = (byte) cypherText.ToCharArray(charPosition++, 1)[0];

				if (isAdd)
					char2 = (byte) (char1 - increment);
				else
					char2 = (byte) (char1 + increment);

				plainText += (char) char2;
			}
		
			return plainText;
		}

		public string GetProxyId()
		{
			_defaultUserId = new GlobalParameter("DefaultUserId").String;
			return _defaultUserId;	//"9940"
		}

		public string GetProxyPwd()
		{
			if (_defaultUserId.Length == 0)
				_defaultUserId = new GlobalParameter("DefaultUserId").String;
			return Decrypt(GetOmigaPassword(_defaultUserId));	//"PEUX";
		}

		public string GetCommunicationChannel(bool isUserLoggedOn)
		{
			if (isUserLoggedOn)
				return "PHONE";		//MAR611 GHun must now be in uppercase
			else
				return "INTERNAL";	//MAR611 GHun must now be in uppercase
		}

		public string GetXmlFromGatewayObject(Type objectType, object targetObject)
		{
			string serXML = null;
			MemoryStream ms = new MemoryStream();
			XmlTextWriter tw = new XmlTextWriter(ms, Encoding.UTF8);
			XmlSerializer serializedObj = new XmlSerializer(objectType);
			Byte[] characters;
			try
			{
				tw = new XmlTextWriter(ms, Encoding.UTF8);
				serializedObj.Serialize(tw, targetObject);
				ms = (MemoryStream)tw.BaseStream;
				characters = ms.ToArray();
				UTF8Encoding encoding = new UTF8Encoding();
				serXML = encoding.GetString(characters, 3, characters.Length-3);
				//serXML = UTF8ByteArrayToString(ms.ToArray());
			}
			catch (Exception ex)
			{
				throw new OmigaException("Unable to serialise gateway object to XML.", ex);
			}
			finally
			{
				tw.Flush();
				ms.Flush();
				tw.Close();
				ms.Close();
				serializedObj = null;
			}
			return serXML;
		}

		//MAR852 GHun
		public string GetOperator()
		{
			GlobalParameter gp = new GlobalParameter("DirectGatewayOperator");
			return gp.String;
		}
		//MAR852 End
	}
}
