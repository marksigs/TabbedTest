using System;
using System.Xml ;
using Microsoft.Win32;
namespace Vertex.Fsd.Omiga.omESurv.Data
{
	public class DataManager
	{
		public string genumDbProvider;
		public string gstrDbConnectionString;
		public int  gintDbRetries;

		public string adoBuildDbConnectionString() 
		{
			string strConnectionString = "";
			string strProvider;
			string strRetries;
			string strUserId;
			string strPassword;
			string strSource;
			string strDatabaseName;

			RegistryKey RegKey;
			RegKey = Registry.LocalMachine.OpenSubKey("SOFTWARE\\Omiga4\\Database Connection");
				
			strDatabaseName = (String)RegKey.GetValue("Database Name");
			strUserId = (String)RegKey.GetValue("User ID");
			strPassword = (String)RegKey.GetValue("Password");
			strProvider = (String)RegKey.GetValue("Provider");
			strSource = (String)RegKey.GetValue("Server");
			strRetries = (String)RegKey.GetValue("Retries");

			if (((strProvider == "MSDAORA") || (strProvider == "ORAOLEDB.ORACLE"))) 
			{
				switch (strProvider) 
				{
					case "MSDAORA":
						//genumDbProvider = DBPROVIDER.omiga4DBPROVIDERMSOracle;
						break;
					case "ORAOLEDB.ORACLE":
						//genumDbProvider = DBPROVIDER.omiga4DBPROVIDEROracle;
						break;
				}
				strConnectionString = "Provider=" + strProvider + ";Data Source=" + strSource + ";User ID=" 
					+ strUserId + ";Password=" + strPassword + ";" ;
			}
			else if ((strProvider == "SQLOLEDB")) 
			{
				strConnectionString = "data Source=" + strSource + ";initial catalog =" + strDatabaseName + ";";
				
				if ((strUserId.Trim().Length > 0)) 
				{
					strConnectionString = strConnectionString + "user id=" + strUserId + ";password=" + strPassword +";" ;
				}
				else 
				{
					strConnectionString = strConnectionString + "Integrated Security=SSPI;Persist Security Info=False";
				}
			}
			
			gstrDbConnectionString = strConnectionString;
			
			if ((strRetries.Length > 0)) 
			{
				gintDbRetries = short.Parse(strRetries);
			}
			
			return gstrDbConnectionString;
		}
		
		public string adoGetDbConnectString() 
		{
			if (gstrDbConnectionString == null ||  gstrDbConnectionString == "" ) 
			{
				return adoBuildDbConnectionString();
			}
			else
			{
				return gstrDbConnectionString;
			}
        }

    }

}