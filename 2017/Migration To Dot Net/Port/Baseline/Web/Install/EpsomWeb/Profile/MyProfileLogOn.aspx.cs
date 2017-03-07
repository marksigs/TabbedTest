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
	/// MyProfile LogOn form.
	/// </summary>
	public class MyProfileLogOn : System.Web.UI.Page
	{
    protected System.Web.UI.WebControls.TextBox txtCurrentEmailAddress;
    protected System.Web.UI.WebControls.TextBox txtEmailAddress;
    protected System.Web.UI.WebControls.TextBox txtConfirmEmailAddress;
    protected System.Web.UI.WebControls.TextBox txtCurrentPassword;
    protected System.Web.UI.WebControls.TextBox txtPasswordNew;
    protected System.Web.UI.WebControls.TextBox txtConfirmPassword;
    protected System.Web.UI.WebControls.Button cmdCancel;
    protected System.Web.UI.WebControls.Button cmdSubmit;
    protected System.Web.UI.WebControls.Label lblMessage;

    private void Page_Load(object sender, System.EventArgs e)
		{
  	}
    private void Page_PreRender(object sender, System.EventArgs e)
    {
      Scatter();
      SetVisibility();
    }
    private void Scatter()
    {
      txtCurrentEmailAddress.Text = HttpSession.CurrentUser.EmailAddress;
      txtCurrentPassword.Attributes["value"] = txtCurrentPassword.Text;
      txtPasswordNew.Attributes["value"] = txtPasswordNew.Text;
      txtConfirmPassword.Attributes["value"] = txtConfirmPassword.Text;
    }
    private void SetVisibility()
    {
    }
    private void cmdSubmit_Click(object sender, System.EventArgs e)
    {
      if (!Page.IsValid) return;
      Epsom.Web.Proxy.UserResponse response;
      response = WebServiceHelper.Instance.AuthenticateUser(HttpSession.CurrentUser.EmailAddress, txtCurrentPassword.Text);
      // no logic to display messages yet
      if (response.User.Firms.Length == 0) return;

      if (txtEmailAddress.Text.Trim().Length != 0)
      {
        HttpSession.CurrentUser.EmailAddress = txtEmailAddress.Text;
      }
      if (txtPasswordNew.Text.Trim().Length != 0)
      {
        HttpSession.CurrentUser.Password = txtPasswordNew.Text;
      }
      // update user with web service here:
      response = WebServiceHelper.Instance.UpdateUser(HttpSession.CurrentUser);
      HttpSession.CurrentUser = response.User;
      Response.Redirect("MyProfile.aspx");
    }
    private void cmdCancel_Click(object sender, System.EventArgs e)
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
      this.PreRender += new System.EventHandler(this.Page_PreRender);
      this.cmdCancel.Click += new System.EventHandler(this.cmdCancel_Click);
      this.cmdSubmit.Click += new System.EventHandler(this.cmdSubmit_Click);
    }
		#endregion
	}
}
