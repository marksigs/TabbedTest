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
using Epsom.Web.Helpers ;

namespace Epsom.Web
{
	/// <summary>
	/// MyProfile Preferences form.
	/// </summary>
	public class MyProfilePreferences : System.Web.UI.Page
	{
    protected System.Web.UI.WebControls.Panel pnlEdit;
    protected System.Web.UI.WebControls.Panel pnlPassword;
    protected System.Web.UI.WebControls.Button cmdCancel;
    protected System.Web.UI.WebControls.Button cmdSubmit;
    protected System.Web.UI.WebControls.TextBox txtPassword;
    protected System.Web.UI.WebControls.Label lblMessage;
    protected System.Web.UI.WebControls.Button cmdSubmitPassword;
    protected System.Web.UI.WebControls.Button cmdCancelPassword;
    //protected Epsom.Web.WebUserControls.UserPreferences ctlUserPreferences;
    protected System.Web.UI.WebControls.CustomValidator CustomValidatorPassword;
    private void Page_Load(object sender, System.EventArgs e)
		{
      if (! IsPostBack )
      {
        pnlPassword.Visible = false; 
        pnlEdit.Visible = true; 
      }
    }
    private void cmdSubmit_Click(object sender, System.EventArgs e)
    {
      if (!Page.IsValid) return;
      //if (!ctlUserPreferences.IsValid()) return;
      pnlPassword.Visible = true; 
      pnlEdit.Visible = false; 
      Epsom.Web.Helpers.FormHelper.SetFocus(txtPassword);
    }
    private void cmdCancel_Click(object sender, System.EventArgs e)
    {
      Response.Redirect("MyProfile.aspx");
    }
    protected void CustomValidatorPassword_ServerValidate(object source, ServerValidateEventArgs args)
    {
      if (args == null) return;
      if (! pnlPassword.Visible) return;
      if (WebServiceHelper.Instance.AuthenticateUser(HttpSession.CurrentUser.EmailAddress, txtPassword.Text) == null)
      {
        args.IsValid = false;
        return;
      }
    }
    private void cmdSubmitPassword_Click(object sender, System.EventArgs e)
    {
      if (!Page.IsValid) return;
      if (HttpSession.CurrentUser == null) Response.Redirect("MyProfile.aspx");
      if (WebServiceHelper.Instance.AuthenticateUser(HttpSession.CurrentUser.EmailAddress, txtPassword.Text) == null)
      {
        lblMessage.Text = "Invalid password";
        return;
      }
      //ctlUserPreferences.Gather();
      Response.Redirect("MyProfile.aspx");
    }
    private void cmdCancelPassword_Click(object sender, System.EventArgs e)
    {
      Response.Redirect("MyProfile.aspx");
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
      this.cmdCancel.Click += new System.EventHandler(this.cmdCancel_Click);
      this.cmdSubmit.Click += new System.EventHandler(this.cmdSubmit_Click);
      this.cmdCancelPassword.Click += new System.EventHandler(this.cmdCancelPassword_Click);
      this.cmdSubmitPassword.Click += new System.EventHandler(this.cmdSubmitPassword_Click);
    }
		#endregion
	}
}
