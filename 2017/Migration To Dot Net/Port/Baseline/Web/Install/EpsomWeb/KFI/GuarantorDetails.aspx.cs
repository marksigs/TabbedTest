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
using Epsom.Web.WebUserControls;
using Epsom.Web.Helpers ;

namespace Epsom.Web.Kfi
{
	/// <summary>
  /// Guarantor Details form.
  /// </summary>
	public class GuarantorDetails : EpsomFormBase
	{

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

    protected Applicant ctlApplicant;
    protected sstchur.web.SmartNav.SmartScroller  ctlSmartScroller;

    private void Page_Load(object sender, System.EventArgs e)
    {
      if (IsPostBack)
      {
        SetSmartScroller();
        ctlApplicant.Gather(HttpSession.CurrentApplication.Guarantor);
      }
      else
      {
        ctlApplicant.Scatter(HttpSession.CurrentApplication.Guarantor);
      }
      ctlApplicant.SetVisibility();
    }

    /// <summary>
    ///   Based on the postback event, determine if the smart scroller is enabled
    /// </summary>
    private void SetSmartScroller()
    {
      System.Web.UI.Control PostBackControl;
      PostBackControl  = WebFormHelper.GetPostBackControl(Page);
      if (PostBackControl.ID == "cmdAddAddress") { ctlSmartScroller.Scroll = true; }
      if (PostBackControl.ID == "cmdValidatePostcode") { ctlSmartScroller.Scroll = true; }
      if (PostBackControl.ID == "cmdResetPostcode") { ctlSmartScroller.Scroll = true; }
      if (PostBackControl.ID == "cmbCountryType") { ctlSmartScroller.Scroll = true; } 
      if (PostBackControl.ID == "chkBritishForcesPostOffice") { ctlSmartScroller.Scroll = true; } 
      if (PostBackControl.ID == "lstValidAddresses") { ctlSmartScroller.Scroll = true; } 
      if (PostBackControl.ID == "cmdResetPostcode") { ctlSmartScroller.Scroll = true; } 
      if (PostBackControl.ID == "cmbResidentFromMonth") { ctlSmartScroller.Scroll = true; } 
      if (PostBackControl.ID == "cmbResidentFromYear") { ctlSmartScroller.Scroll = true; } 
    }


	}
}
