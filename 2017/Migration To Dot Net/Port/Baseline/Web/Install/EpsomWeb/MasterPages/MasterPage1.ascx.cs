namespace Epsom.Web.MasterPages
{
  using System;
  using System.Configuration;
  using System.Data;
  using System.Data.SqlClient;
  using System.Drawing;
  using System.Web;
  using System.Web.UI.WebControls;
  using System.Web.UI.HtmlControls;

	/// <summary>
	///		Summary description for MasterPage1.
	/// </summary>
	public class MasterPage1 : System.Web.UI.UserControl
	{
    protected Epsom.Web.WebUserControls.Links ctlUsefulLinks;
    protected Epsom.Web.WebUserControls.Links ctlCalculators;

		private void Page_Load(object sender, System.EventArgs e)
		{
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
		///		Required method for Designer support - do not modify
		///		the contents of this method with the code editor.
		/// </summary>
		private void InitializeComponent()
		{
			this.Load += new System.EventHandler(this.Page_Load);
		}
		#endregion
	}
}
