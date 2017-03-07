using System;
using System.Collections;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Reflection;
using System.Web;
using System.Web.Security;
using System.Web.SessionState;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Web.UI.HtmlControls;
using Epsom.Web.Helpers ;
using Vertex.Fsd.Omiga.omLogging;

namespace Epsom.Web
{
	/// <summary>
	/// Form to collect broker registration details.
	/// </summary>
	public class Register : System.Web.UI.Page
	{

    protected System.Web.UI.WebControls.Label lblFormErrorMessage;
    protected System.Web.UI.WebControls.RadioButton rbAr;
    protected System.Web.UI.WebControls.RadioButton rbDa;
    protected System.Web.UI.WebControls.TextBox txtFsaNumber;
    protected System.Web.UI.WebControls.Label lblFsaCompanyName;
    protected System.Web.UI.WebControls.Label lblFindFirmMessage;
    protected System.Web.UI.WebControls.TextBox txtTradingAs;
    protected System.Web.UI.WebControls.DropDownList cmbTitle;
    protected System.Web.UI.WebControls.TextBox txtFirstName;
    protected System.Web.UI.WebControls.TextBox txtLastName;
    protected System.Web.UI.WebControls.TextBox txtBirth;
    protected System.Web.UI.WebControls.TextBox txtNINumber;
    protected System.Web.UI.WebControls.TextBox txtDeskTelephoneAreaCode;
    protected System.Web.UI.WebControls.TextBox txtDeskTelephoneNumber;
    protected System.Web.UI.WebControls.TextBox txtDeskTelephoneExtension;
    protected System.Web.UI.WebControls.TextBox txtMobileTelephone;
    protected System.Web.UI.WebControls.TextBox txtOtherTelephoneAreaCode;
    protected System.Web.UI.WebControls.TextBox txtOtherTelephoneNumber;
    protected System.Web.UI.WebControls.TextBox txtOtherTelephoneExtension;
    protected System.Web.UI.WebControls.DropDownList cmbContactTelephoneUsage;

    protected System.Web.UI.WebControls.Panel pnlDetails;
    protected System.Web.UI.WebControls.Panel pnlFirmDetails;
    protected System.Web.UI.WebControls.Panel pnlFindFirmMessage;
    protected System.Web.UI.WebControls.TextBox txtEmailAddress;
    protected System.Web.UI.WebControls.TextBox txtConfirmEmailAddress;
    protected System.Web.UI.WebControls.TextBox txtPassword;
    protected System.Web.UI.WebControls.TextBox txtConfirmPassword;
    protected Epsom.Web.WebUserControls.Address ctlAddress;
    protected System.Web.UI.WebControls.Button cmdFindFirm;
    protected System.Web.UI.WebControls.Button cmdResetFirm;
    protected System.Web.UI.WebControls.Button cmdRegister;
    protected System.Web.UI.WebControls.Button cmdCancel;
    protected System.Web.UI.WebControls.CustomValidator validatorFsaNumber;
    protected System.Web.UI.WebControls.CustomValidator validatorPassword;
    protected System.Web.UI.WebControls.CustomValidator validatorConfirmPassword;
    protected sstchur.web.SmartNav.SmartScroller ctlSmartScroller;

    private Epsom.Web.Proxy.ResponseMessage _FindFirmMessage;
    private string homepage = "~/Default.aspx";


    
    private Epsom.Web.Proxy.User RegisteringUser
    {
      get {return (Epsom.Web.Proxy.User)ViewState["RegisteringUser"];}
      set { ViewState["RegisteringUser"]=value;}
    }

    private void Page_Load(object sender, System.EventArgs e)
		{
      if (IsPostBack)
      {

        Gather(); 
        SetSmartScroller();


        // Repopulate Password on every postback
        txtPassword.Attributes["value"] = txtPassword.Text;
        txtConfirmPassword.Attributes["value"] = txtConfirmPassword.Text;
      }
      else
      {
        RegisteringUser = Epsom.Web.Helpers.WebServiceHelper.Instance.NewUser();
        RegisteringUser.IsBroker = true;
        Scatter();
      }
    }

    private void Page_PreRender(object sender, System.EventArgs e)
    {
      SetVisibility();
    }

    private void Gather()
    {
      if (RegisteringUser != null)  // assume that there's nothing to gather yet
      {
        RegisteringUser.EmailAddress = txtEmailAddress.Text;
        RegisteringUser.Title = cmbTitle.SelectedItem.Value;
        RegisteringUser.FirstName = txtFirstName.Text;
        RegisteringUser.LastName = txtLastName.Text;
        //RegisteringUser.CompanyName = txtTradingAs.Text;
        //RegisteringUser.NINumber = txtNINumber.Text.ToUpper(System.Globalization.CultureInfo.CurrentCulture);
        RegisteringUser.BirthDate = Web.Helpers.DateTimeHelper.FromString(txtBirth.Text);
        //RegisteringUser.TelephoneDesk = txtDeskTelephoneNumber.Text;
        //RegisteringUser.TelephoneMobile = txtMobileTelephone.Text;
        RegisteringUser.Password = txtPassword.Text;
        RegisteringUser.Address = ctlAddress.Gather();

        RegisteringUser.IsAr = rbAr.Checked;

        RegisteringUser.PhoneNumbers[0].AreaCode = txtDeskTelephoneAreaCode.Text ;
        RegisteringUser.PhoneNumbers[0].Number = txtDeskTelephoneNumber.Text ;
        RegisteringUser.PhoneNumbers[0].Extension = txtDeskTelephoneExtension.Text ;

        RegisteringUser.PhoneNumbers[1].Number = txtMobileTelephone.Text;

        RegisteringUser.PhoneNumbers[2].AreaCode = txtOtherTelephoneAreaCode.Text ;
        RegisteringUser.PhoneNumbers[2].Number = txtOtherTelephoneNumber.Text;
        RegisteringUser.PhoneNumbers[2].Extension = txtOtherTelephoneExtension.Text ;
        RegisteringUser.PhoneNumbers[2].Usage  = cmbContactTelephoneUsage.SelectedItem.Value ;
      }
    }

    private void Scatter()
    {
      PopulateCombos();

      if (RegisteringUser == null) return;
      if (!RegisteringUser.IsBroker ) return;

      txtPassword.Attributes["value"] = txtPassword.Text;
      txtConfirmPassword.Attributes["value"] = txtConfirmPassword.Text;

      if (RegisteringUser.Firms.Length > 0) 
      {
        txtFsaNumber.Text = RegisteringUser.Firms[0].FsaNumber ;
        lblFsaCompanyName.Text = RegisteringUser.Firms[0].CompanyName ;
      }

      if ( _FindFirmMessage != null) {lblFindFirmMessage.Text = _FindFirmMessage.Text;}
      ctlAddress.Scatter(RegisteringUser.Address,false);

    }

    private void PopulateCombos()
    {
      WebFormHelper.PopulateListControl(cmbContactTelephoneUsage, Web.Helpers.ComboValues.ComboGroup(Web.Helpers.ComboValues.ComboGroupType.ContactTelephoneUsage), true);
      WebFormHelper.PopulateListControl(cmbTitle, Web.Helpers.ComboValues.ComboGroup(Web.Helpers.ComboValues.ComboGroupType.Title ), true);
    }

    private void SetVisibility()
    {
      if (RegisteringUser != null)
      {
        // Show either Firm Finder or Firm Details
        if (RegisteringUser.Firms.Length > 0) 
        {
          this.pnlFirmDetails.Visible =true;
          this.txtFsaNumber.Enabled = false;
          this.cmdFindFirm.Enabled = false;
          this.rbAr.Enabled = false;
          this.rbDa.Enabled = false;
          this.cmdResetFirm.Enabled = true;
        }
        else
        {
          this.pnlFirmDetails.Visible =false;
          this.txtFsaNumber.Enabled = true;
          this.cmdFindFirm.Enabled = true;
          this.cmdResetFirm.Enabled = false;
        }

        ctlAddress.SetVisibility();
      }  
      pnlFindFirmMessage.Visible = (_FindFirmMessage != null);
    }

    private void SetSmartScroller()
    {
      System.Web.UI.Control PostBackControl;
      PostBackControl = WebFormHelper.GetPostBackControl(Page);

      if (PostBackControl.ID == "cmdFindFirm") { ctlSmartScroller.Scroll = true; }
      if (PostBackControl.ID == "cmdResetFirm"){ ctlSmartScroller.Scroll = true; }
      ctlAddress.SetSmartScroller(ctlSmartScroller,PostBackControl); 
    }

    protected void ValidateFirm(object source, ServerValidateEventArgs e)
    {
      if (e == null) return;
      if (RegisteringUser.Firms.Length == 0) e.IsValid = false;
      else e.IsValid = true;
      return;
    }
    
    protected void cmdFindFirm_Click(object sender, System.EventArgs e)
    {
      Epsom.Web.Proxy.UserResponse userResponse;
      

      userResponse = WebServiceHelper.Instance.AddFirmToUser(RegisteringUser, txtFsaNumber.Text);
      if (userResponse.ResponseMessages.Length > 0)
      {
        _FindFirmMessage = userResponse.ResponseMessages[0];
      }

      RegisteringUser = userResponse.User;
      Scatter();
    }

    protected void cmdResetFirm_Click(object sender, System.EventArgs e)
    {
      RegisteringUser.Firms = new Epsom.Web.Proxy.Firm[0];
      Scatter();
    }

    protected void cmdCancel_Click(object sender, System.EventArgs e)
    {
      Response.Redirect(homepage);
    }

    protected void cmdRegister_Click(object sender, System.EventArgs e)
    {
      // Don't process the registration if the page is not valid
      if ( ! Page.IsValid) { return ; }
      
      Epsom.Web.Proxy.UserResponse response = WebServiceHelper.Instance.RegisterNewUser(RegisteringUser);
      HttpSession.CurrentUser = response.User;
      HttpSession.CurrentFirmIndex = 0;

      // This duplicates what is in Logon - may need method
      HttpCookie cookie = FormsAuthentication.GetAuthCookie(HttpSession.CurrentUser.EmailAddress, false);
      cookie.Expires = DateTime.Now.AddMinutes(Session.Timeout);
      Response.Cookies.Add(cookie);
      Response.Redirect("RegistrationConfirmation.aspx");
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
      this.cmdFindFirm.Click += new System.EventHandler(this.cmdFindFirm_Click);
      this.cmdResetFirm.Click += new System.EventHandler(this.cmdResetFirm_Click);
      this.cmdRegister.Click += new System.EventHandler(this.cmdRegister_Click);
      this.cmdCancel.Click += new System.EventHandler(this.cmdCancel_Click);

      cmdCancel.CausesValidation = false;

    }
		#endregion
	}
}
