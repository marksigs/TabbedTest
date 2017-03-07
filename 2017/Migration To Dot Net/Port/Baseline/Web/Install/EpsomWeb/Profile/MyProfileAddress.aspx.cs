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

namespace Epsom.Web
{
	/// <summary>
	/// MyProfile Email Address Form
	/// </summary>
	public class MyProfileAddress : System.Web.UI.Page
	{
    protected System.Web.UI.WebControls.Panel pnlEdit;
    protected System.Web.UI.WebControls.Panel pnlPassword;
    protected System.Web.UI.WebControls.Button cmdCancel;
    protected System.Web.UI.WebControls.Button cmdSubmit;
    protected System.Web.UI.WebControls.TextBox txtDeskTelephoneAreaCode;
    protected System.Web.UI.WebControls.TextBox txtDeskTelephoneNumber;
    protected System.Web.UI.WebControls.TextBox txtDeskTelephoneExtension;
    protected System.Web.UI.WebControls.TextBox txtMobileTelephone;
    protected System.Web.UI.WebControls.TextBox txtOtherTelephoneAreaCode;
    protected System.Web.UI.WebControls.TextBox txtOtherTelephoneNumber;
    protected System.Web.UI.WebControls.TextBox txtOtherTelephoneExtension;
    protected System.Web.UI.WebControls.DropDownList cmbContactTelephoneUsage;
    protected System.Web.UI.WebControls.Label lblMessage;
    protected Epsom.Web.WebUserControls.Address ctlAddress;
    protected Epsom.Web.WebUserControls.UserConfirmPassword ctlUserConfirmPassword;
    protected sstchur.web.SmartNav.SmartScroller ctlSmartScroller;

    private void Page_Load(object sender, System.EventArgs e)
		{
      if (IsPostBack)
      {
//        Gather();        
        SetSmartScroller();
      }
      else
      {
        Scatter();
      } 
    }

    private void Page_PreRender(object sender, System.EventArgs e)
    {
      // password will be non-null if it has been validated and user update is required
      string password = ctlUserConfirmPassword.Password;
      if (password == null)
      {
        SetVisibility();
        return;
      }
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

      txtDeskTelephoneAreaCode.Text = HttpSession.CurrentUser.PhoneNumbers[0].AreaCode;
      txtDeskTelephoneNumber.Text = HttpSession.CurrentUser.PhoneNumbers[0].Number;
      txtDeskTelephoneExtension.Text = HttpSession.CurrentUser.PhoneNumbers[0].Extension;

      txtMobileTelephone.Text = HttpSession.CurrentUser.PhoneNumbers[1].Number;

      txtOtherTelephoneAreaCode.Text = HttpSession.CurrentUser.PhoneNumbers[2].AreaCode;
      txtOtherTelephoneNumber.Text = HttpSession.CurrentUser.PhoneNumbers[2].Number;
      txtOtherTelephoneExtension.Text = HttpSession.CurrentUser.PhoneNumbers[2].Extension;
      WebFormHelper.SelectComboItem(cmbContactTelephoneUsage, HttpSession.CurrentUser.PhoneNumbers[2].Usage);

      ctlAddress.Scatter(HttpSession.CurrentUser.Address);
    }

    private void Gather()
    {
      HttpSession.CurrentUser.Address = ctlAddress.Gather();

      HttpSession.CurrentUser.PhoneNumbers[0].AreaCode = txtDeskTelephoneAreaCode.Text ;
      HttpSession.CurrentUser.PhoneNumbers[0].Number = txtDeskTelephoneNumber.Text ;
      HttpSession.CurrentUser.PhoneNumbers[0].Extension = txtDeskTelephoneExtension.Text ;

      HttpSession.CurrentUser.PhoneNumbers[1].Number = txtMobileTelephone.Text;

      HttpSession.CurrentUser.PhoneNumbers[2].AreaCode = txtOtherTelephoneAreaCode.Text ;
      HttpSession.CurrentUser.PhoneNumbers[2].Number = txtOtherTelephoneNumber.Text;
      HttpSession.CurrentUser.PhoneNumbers[2].Extension = txtOtherTelephoneExtension.Text ;
      HttpSession.CurrentUser.PhoneNumbers[2].Usage  = cmbContactTelephoneUsage.SelectedItem.Value ;
    }

    private void SetVisibility()
    {
//      pnlPassword.Visible = false; 
      ctlAddress.SetVisibility(); 
//      pnlEdit.Visible = true; 
    }

    private void SetSmartScroller()
    {
      System.Web.UI.Control PostBackControl;
      PostBackControl  = WebFormHelper.GetPostBackControl(Page);
      ctlAddress.SetSmartScroller(ctlSmartScroller,PostBackControl); 
    }

    private void PopulateCombos()
    {
      WebFormHelper.PopulateListControl(cmbContactTelephoneUsage, Web.Helpers.ComboValues.ComboGroup(Web.Helpers.ComboValues.ComboGroupType.ContactTelephoneUsage), true);
    }


    /// <summary>
    /// Show password and set focus on it.
    /// </summary>
    /// <param name="sender"></param>
    /// <param name="e"></param>
    private void cmdSubmit_Click(object sender, System.EventArgs e)
    {
      if ( ! Page.IsValid) {return;} 
                            
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
