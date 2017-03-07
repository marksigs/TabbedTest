/*
--------------------------------------------------------------------------------------------
Workfile:			omHomeTrack.aspx.cs
Copyright:			Copyright ©2005 Marlborough Stirling

Description:		Call HomeTrackBO to get latest present valuation
--------------------------------------------------------------------------------------------
History:

Prog	Date		Description
JD	29/11/2005	MAR709 Created
--------------------------------------------------------------------------------------------
*/
using System;
using System.Web;
using System.Web.UI;
using System.IO;
using System.Text;
using Vertex.Fsd.Omiga.omHomeTrack;

namespace Vertex.Fsd.Omiga.Web
{
	/// <summary>
	/// Summary description for GetPresentValuation
	/// </summary>
	public class GetPresentValuation : System.Web.UI.Page
	{
		private void Page_Load(object sender, System.EventArgs e)
		{
			StreamReader sr = new StreamReader(Request.InputStream);
			string requestString = sr.ReadToEnd();

			HomeTrackBO BO = new HomeTrackBO();
			string responseString = BO.GetPresentValuation(requestString);

			Response.ContentType = "text/xml";
			Response.Write(responseString);
		}
	}
}
