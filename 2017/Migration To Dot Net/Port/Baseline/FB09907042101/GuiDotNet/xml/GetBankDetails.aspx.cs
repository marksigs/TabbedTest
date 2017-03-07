/*
--------------------------------------------------------------------------------------------
Workfile:			GetBankDetails.aspx.cs
Copyright:			Copyright ©2005 Marlborough Stirling

Description:		Call BankWizardBO to validate bank details
--------------------------------------------------------------------------------------------
History:

Prog	Date		Description
GHun	08/10/2005	MAR137 Created
--------------------------------------------------------------------------------------------
*/
using System;
using System.Web;
using System.Web.UI;
using System.IO;
using System.Text;
using Vertex.Fsd.Omiga.omBankWizard;

namespace Vertex.Fsd.Omiga.Web
{
	/// <summary>
	/// Summary description for GetBankDetails
	/// </summary>
	public class GetBankDetails : System.Web.UI.Page
	{
		private void Page_Load(object sender, System.EventArgs e)
		{
			StreamReader sr = new StreamReader(Request.InputStream);
			string requestString = sr.ReadToEnd();

			BankWizardBO bwbo = new BankWizardBO();
			string responseString = bwbo.GetBankDetails(requestString);

			Response.ContentType = "text/xml";
			Response.Write(responseString);
		}
	}
}
