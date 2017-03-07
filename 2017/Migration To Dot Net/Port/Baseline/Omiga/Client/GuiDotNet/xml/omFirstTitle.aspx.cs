/*
--------------------------------------------------------------------------------------------
Workfile:			omFirstTitle.aspx.cs
Copyright:			Copyright ©2005 Marlborough Stirling

Description:		Call omFirstTitle Execute Method
--------------------------------------------------------------------------------------------
History:

Prog	Date		Description
MV		25/10/2005	MAR236 Created
--------------------------------------------------------------------------------------------
*/
using System;
using System.Collections;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Web;
using System.Web.SessionState;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Web.UI.HtmlControls;
using System.IO;
using System.Text;
using Vertex.Fsd.Omiga.omFirstTitle;

namespace Vertex.Fsd.Omiga.Web
{
	public class omFirstTitle : System.Web.UI.Page
	{
		private void Page_Load(object sender, System.EventArgs e)
		{
			
			StreamReader sr = new StreamReader(Request.InputStream);
			string requestString = sr.ReadToEnd();

			FirstTitleOutboundBO FirstTitleOutboundBO = new FirstTitleOutboundBO();
			string responseString = FirstTitleOutboundBO.Execute(requestString);

			Response.ContentType = "text/xml";
			Response.Write(responseString);
		}

		#region Web Form Designer generated code
		override protected void OnInit(EventArgs e)
		{
			//
			// CODEGEN: This call is required by the ASP.NET Web Form Designer.
			//
			InitializeComponent();
			base.OnInit(e);
		}
		
		/// <summary>
		/// Required method for Designer support - do not modify
		/// the contents of this method with the code editor.
		/// </summary>
		private void InitializeComponent()
		{    
			this.Load += new System.EventHandler(this.Page_Load);
		}
		#endregion
	}
}
