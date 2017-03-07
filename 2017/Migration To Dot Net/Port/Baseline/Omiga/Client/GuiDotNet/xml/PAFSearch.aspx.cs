/*
--------------------------------------------------------------------------------------------
Workfile:			PAFSearch.aspx.cs
Copyright:			Copyright ©2005 Marlborough Stirling

Description:		Call UniservBO to validate address details
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
using Vertex.Fsd.Omiga.omUniserv;

namespace Vertex.Fsd.Omiga.Web
{
	/// <summary>
	/// Summary description for PAFSearch
	/// </summary>
	public class PafSearch : System.Web.UI.Page
	{
		private void Page_Load(object sender, System.EventArgs e)
		{
			StreamReader sr = new StreamReader(Request.InputStream);
			string requestString = sr.ReadToEnd();

			UniservBO ubo = new UniservBO();
			string responseString = ubo.FindPAFAddress(requestString);

			Response.ContentType = "text/xml";
			Response.Write(responseString);
		}
	}
}
