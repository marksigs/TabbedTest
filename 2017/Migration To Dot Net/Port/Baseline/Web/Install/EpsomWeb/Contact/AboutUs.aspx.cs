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

namespace Epsom.Web
{
	/// <summary>
	/// Summary description for AboutUs.
	/// </summary>
	public class AboutUs : System.Web.UI.Page
	{

    protected System.Web.UI.WebControls.Label lblLetterRange;
    protected System.Web.UI.WebControls.Label lblNoPackagersFoundMessage;
    protected System.Web.UI.WebControls.Repeater rptApprovedPackagers;

    private void Page_Load(object sender, System.EventArgs e)
		{
      if (!IsPostBack)
      {
        Scatter("A", "D");
      }
    }

    private void Scatter(string startLetter, string endLetter)
    {
      lblLetterRange.Text = startLetter + " - " + endLetter;
      
      Epsom.Web.Proxy.ApprovedPackager[] approvedPackagers = Epsom.Web.Helpers.WebServiceHelper.Instance.ApprovedPackagerWebService(startLetter, endLetter);
      rptApprovedPackagers.DataSource = approvedPackagers;
      rptApprovedPackagers.DataBind();

      SetVisibility();
    }

    private void SetVisibility()
    {
      lblNoPackagersFoundMessage.Visible = (rptApprovedPackagers.Items.Count == 0);
    }

    public void DisplayLetterRange(object sender, CommandEventArgs e)
    {
      string[] letterRange = ((string)e.CommandArgument).Split(',');
      Scatter(letterRange[0], letterRange[1]);
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
