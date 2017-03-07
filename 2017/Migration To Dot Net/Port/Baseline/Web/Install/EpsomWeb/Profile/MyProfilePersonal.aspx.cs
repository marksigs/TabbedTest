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
using Epsom.Web.Helpers;

namespace Epsom.Web.Profile
{
	/// <summary>
	/// Summary description for MyProfilePersonal.
	/// </summary>
	public class MyProfilePersonal : System.Web.UI.Page
	{
    protected System.Web.UI.WebControls.Panel pnlEdit;
    protected System.Web.UI.WebControls.Panel pnlPassword;
    protected System.Web.UI.WebControls.Button cmdCancel;
    protected System.Web.UI.WebControls.Button cmdSubmit;
    protected System.Web.UI.WebControls.DropDownList cmbTitle;
    protected System.Web.UI.WebControls.TextBox txtBirth;
    protected System.Web.UI.WebControls.TextBox txtFirstName;
    protected System.Web.UI.WebControls.TextBox txtLastName;
    protected Epsom.Web.WebUserControls.UserConfirmPassword ctlUserConfirmPassword;
    protected sstchur.web.SmartNav.SmartScroller ctlSmartScroller;

    private void Page_Load(object sender, System.EventArgs e)
		{
      if (! IsPostBack)
      {
        Scatter();
        Web.Helpers.FormHelper.SetFocus(txtFirstName);
      }
    }

    private void Page_PreRender(object sender, System.EventArgs e)
    {
      // password will be non-null if it has been validated and user update is required
      string password = ctlUserConfirmPassword.Password;
      if (password == null) return;
      pnlPassword.Visible = false;
      pnlEdit.Visible = true;
      Gather();
      // update user with web service here:
      Epsom.Web.Proxy.UserResponse response = WebServiceHelper.Instance.UpdateUser(HttpSession.CurrentUser);
      HttpSession.CurrentUser = response.User;
      Response.Redirect("MyProfile.aspx");
    }
    private void Scatter()
    {
      PopulateCombos();
      txtFirstName.Text = HttpSession.CurrentUser.FirstName;
      txtLastName.Text = HttpSession.CurrentUser.LastName;
      WebFormHelper.SelectComboItem(cmbTitle, HttpSession.CurrentUser.Title);
      txtBirth.Text = HttpSession.CurrentUser.BirthDate.ToShortDateString();
    }
    private void Gather()
    {
      HttpSession.CurrentUser.FirstName = txtFirstName.Text;
      HttpSession.CurrentUser.LastName = txtLastName.Text;
      HttpSession.CurrentUser.Title = cmbTitle.SelectedItem.Value;
    }
    private void PopulateCombos()
    {
      WebFormHelper.PopulateListControl(cmbTitle, Web.Helpers.ComboValues.ComboGroup(Web.Helpers.ComboValues.ComboGroupType.Title ), true);
    }

    /// <summary>
    /// Show password and set focus on it.
    /// </summary>
    /// <param name="sender"></param>
    /// <param name="e"></param>
    private void cmdSubmit_Click(object sender, System.EventArgs e)
    {
      pnlPassword.Visible = true; 
      pnlEdit.Visible = false;
      ctlUserConfirmPassword.SetVisibility();
    }

    /// <summary>
    /// Redirect to MyProfile form.
    /// </summary>
    /// <param name="sender"></param>
    /// <param name="e"></param>
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
      this.cmdSubmit.Click += new System.EventHandler(this.cmdSubmit_Click);
      this.cmdCancel.Click += new System.EventHandler(this.cmdCancel_Click);
    }
		#endregion
	}
}
