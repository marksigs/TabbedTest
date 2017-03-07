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
	/// Form to confirm a broker Registration.
	/// </summary>
	public class RegistrationConfirmation : System.Web.UI.Page
	{
    protected System.Web.UI.WebControls.PlaceHolder placeHolder1;
    protected System.Web.UI.WebControls.Label lblFsaNumber;
    protected System.Web.UI.WebControls.Label lblFsaCompanyName;
    protected System.Web.UI.WebControls.Label lblFirstName;
    protected System.Web.UI.WebControls.Label lblLastName;
    //protected System.Web.UI.WebControls.Label lblNINumber;
    protected System.Web.UI.WebControls.Label lblAddress1;
    protected System.Web.UI.WebControls.Label lblAddress2;
    protected System.Web.UI.WebControls.Label lblAddress3;
    protected System.Web.UI.WebControls.Label lblAddress4;
    protected System.Web.UI.WebControls.Label lblAddress5;
    protected System.Web.UI.WebControls.Label lblDateBirth;
    protected System.Web.UI.WebControls.Label lblTelephoneDesk;
    protected System.Web.UI.WebControls.Label lblTelephoneOther;
    protected System.Web.UI.WebControls.Label lblTelephoneOtherTitle;
    protected System.Web.UI.WebControls.Label lblTelephoneMobile;
    protected System.Web.UI.WebControls.Label lblTitle;
    protected System.Web.UI.WebControls.Label lblEmailAddress;
    protected System.Web.UI.WebControls.Button cmdEditProfile;
    protected System.Web.UI.WebControls.Button cmdDone;
    protected Epsom.Web.WebUserControls.AddressDisplay ctlAddress;
    protected System.Web.UI.WebControls.Panel pnlTelephoneOther;

    private void Page_Load(object sender, System.EventArgs e)
		{
      if (! Page.IsPostBack) Scatter();
		}

    private void Page_PreRender(object sender, System.EventArgs e)
    {
      SetVisibility();
    }

    private void Scatter()
    {
      if (HttpSession.CurrentUser == null) return;
      lblEmailAddress.Text = HttpSession.CurrentUser.EmailAddress;
      lblFsaNumber.Text = HttpSession.CurrentUser.Firms[0].FsaNumber;
      lblFsaCompanyName.Text = HttpSession.CurrentUser.Firms[0].CompanyName;
      lblTitle.Text = ComboValues.ComboItemText(ComboValues.ComboGroupType.Title, HttpSession.CurrentUser.Title);
      lblFirstName.Text = HttpSession.CurrentUser.FirstName;
      lblLastName.Text = HttpSession.CurrentUser.LastName;
      //lblNINumber.Text = HttpSession.CurrentUser.NINumber;
      lblDateBirth.Text = HttpSession.CurrentUser.BirthDate.ToShortDateString();
      // might need to put in the company name
//      if (!Epsom.Web.Helpers.StringHelper.IsNullOrEmpty(HttpSession.CurrentUser.Address.Organisation)) 
//      {
//        HttpSession.CurrentUser.Address.Organisation=HttpSession.CurrentUser.CompanyName;
//      }
      ctlAddress.Scatter(HttpSession.CurrentUser.Address);

      // telephone fields
      lblTelephoneDesk.Text = Helpers.PhoneHelper.FormatPhoneNumber(HttpSession.CurrentUser.PhoneNumbers[0]);
      lblTelephoneMobile.Text = Helpers.PhoneHelper.FormatPhoneNumber(HttpSession.CurrentUser.PhoneNumbers[1]);
      lblTelephoneOther.Text = Helpers.PhoneHelper.FormatPhoneNumber(HttpSession.CurrentUser.PhoneNumbers[2]);
      lblTelephoneOtherTitle.Text = Web.Helpers.ComboValues.ComboItemText(Epsom.Web.Helpers.ComboValues.ComboGroupType.ContactTelephoneUsage, HttpSession.CurrentUser.PhoneNumbers[2].Usage); 
    }
  
    private void SetVisibility()
    {

      pnlTelephoneOther.Visible = (HttpSession.CurrentUser.PhoneNumbers[2].Number != "" && HttpSession.CurrentUser.PhoneNumbers[2].Number != null);
   
      ctlAddress.SetVisibility();
    }

    protected void cmdDone_Click(object sender, System.EventArgs e)
    {
      Response.Redirect ("~/");  
    }

    protected void cmdEditProfile_Click(object sender, System.EventArgs e)
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
      this.cmdDone.Click += new System.EventHandler(this.cmdDone_Click);
      this.cmdEditProfile.Click += new System.EventHandler(this.cmdEditProfile_Click);
    }
		#endregion
	}
}
