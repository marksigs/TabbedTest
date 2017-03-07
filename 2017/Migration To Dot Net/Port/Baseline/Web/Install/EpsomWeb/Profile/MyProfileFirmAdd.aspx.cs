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
  /// Add firm to user profile.
  /// </summary>
  public class MyProfileFirmAdd : System.Web.UI.Page
  {
    protected System.Web.UI.WebControls.Label lblFormErrorMessage;
    protected System.Web.UI.WebControls.RadioButton rbAr;
    protected System.Web.UI.WebControls.RadioButton rbDa;
    protected System.Web.UI.WebControls.TextBox txtFsaNumber;
    protected System.Web.UI.WebControls.Label lblFsaCompanyName;
    protected System.Web.UI.WebControls.Label lblFindFirmMessage;
    protected System.Web.UI.WebControls.Panel pnlEdit;
    protected System.Web.UI.WebControls.Panel pnlFindFirmMessage;
    protected System.Web.UI.WebControls.Panel pnlFirmDetails;
    protected System.Web.UI.WebControls.Panel pnlPassword;
    protected System.Web.UI.WebControls.Button cmdValidate;
    protected System.Web.UI.WebControls.Button cmdReset;
    protected System.Web.UI.WebControls.Button cmdCancel;
    protected System.Web.UI.WebControls.Button cmdSubmit;
    protected Epsom.Web.WebUserControls.UserConfirmPassword ctlUserConfirmPassword;

//    private Epsom.Web.Proxy.ResponseMessage _FindFirmMessage;

    private void Page_Load(object sender, System.EventArgs e)
    {
      if (! IsPostBack)
      {
        Web.Helpers.FormHelper.SetFocus(txtFsaNumber);
      }

      Scatter();
    }

    private void Page_PreRender(object sender, System.EventArgs e)
    {
      SetVisibility();

      Epsom.Web.Proxy.UserResponse userResponse;
      //Epsom.Web.Proxy.Firm firm;
      // password will be non-null if it has been validated and user update is required
      string password = ctlUserConfirmPassword.Password;
      if (password == null) return;
      pnlPassword.Visible = false;
      pnlEdit.Visible = true;
      // update user with web service here:
      //firm = (Epsom.Web.Proxy.Firm)ViewState["Firm"];
      userResponse = WebServiceHelper.Instance.UpdateUser((Epsom.Web.Proxy.User)ViewState["UpdatingUser"]);
      HttpSession.CurrentUser = userResponse.User;
      Response.Redirect("MyProfile.aspx");

      
    }

    private void Scatter()
    {
      rbAr.Checked = HttpSession.CurrentUser.IsAr;
      rbDa.Checked = ! HttpSession.CurrentUser.IsAr;
    }

    private void SetVisibility()
    {
      rbAr.Enabled = false;
      rbDa.Enabled = false;
    }

    private void cmdValidate_Click(object sender, System.EventArgs e)
    {
      Epsom.Web.Proxy.UserResponse userResponse =
        WebServiceHelper.Instance.AddFirmToUser(HttpSession.CurrentUser, txtFsaNumber.Text);
      if (userResponse.ResponseMessages.Length > 0)
      {
        lblFindFirmMessage.Text = userResponse.ResponseMessages[0].Text;
        pnlFindFirmMessage.Visible = true;
        pnlFirmDetails.Visible = false;
        return;
      }
      pnlFirmDetails.Visible = true;
      pnlFindFirmMessage.Visible = false;
      txtFsaNumber.Enabled = false;

      ViewState["UpdatingUser"] = userResponse.User;
      lblFsaCompanyName.Text = userResponse.User.Firms[userResponse.User.Firms.Length - 1].CompanyName;

      // HttpSession.CurrentUser = responseAddFirmToUser.User;
      // Scatter();
      //      SetVisibility();
    }

    private void cmdSubmit_Click(object sender, System.EventArgs e)
    {
      if (!Page.IsValid) return;
      // test if user has clicked "validate" button first
      if (txtFsaNumber.Enabled == true)
      {
        lblFindFirmMessage.Text = "Please click the Validate button to validate the FSA number";
        pnlFindFirmMessage.Visible = true;
        pnlFirmDetails.Visible = false;
        return;
      }
      pnlPassword.Visible = true; 
      pnlEdit.Visible = false; 
      ctlUserConfirmPassword.SetVisibility();
    }

    private void cmdReset_Click(object sender, System.EventArgs e)
    {
      pnlFirmDetails.Visible = false;
      pnlFindFirmMessage.Visible = false;
      txtFsaNumber.Enabled = true;
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
      this.cmdValidate.Click += new System.EventHandler(this.cmdValidate_Click);
      this.cmdReset.Click += new System.EventHandler(this.cmdReset_Click);
      this.cmdSubmit.Click += new System.EventHandler(this.cmdSubmit_Click);
      this.cmdCancel.Click += new System.EventHandler(this.cmdCancel_Click);
    }
    #endregion
  }
}
