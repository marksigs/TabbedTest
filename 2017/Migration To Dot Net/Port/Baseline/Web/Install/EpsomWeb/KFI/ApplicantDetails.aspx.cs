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
using Epsom.Web.Helpers;

namespace Epsom.Web.Kfi
{
  /// <summary>
  /// Applicant Details form.
  /// </summary>
  public class ApplicantDetails : EpsomFormBase
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
      this.cmdAddApplicant.Click += new System.EventHandler(this.CmdAddApplicant_Click);
    }

    #endregion

    protected System.Web.UI.WebControls.Label lblHeading;
    protected Applicant ctlApplicant;
    protected System.Web.UI.WebControls.ValidationSummary Validationsummary1;
    protected System.Web.UI.WebControls.Button cmdAddApplicant;
    protected sstchur.web.SmartNav.SmartScroller ctlSmartScroller;

    private int CurrentApplicantIndex
    {
      get 
      { 
        if (Epsom.Web.Helpers.NumberHelper.IsNumeric(Request.QueryString["applicant"]))
        {
          return Convert.ToInt32(Request.QueryString["applicant"]) - 1;
        }
        else
        {
          return 0;
        }
      }
    }

    private void Page_Load(object sender, System.EventArgs e)
    {
      if (IsPostBack)
      {
        SetSmartScroller();
        Gather();        
      }
      else
      {
        Scatter();
      } 

      SetVisibility();
    }

    private void Scatter()
    {
      ctlApplicant.Scatter(HttpSession.CurrentApplication.Applicants[CurrentApplicantIndex]);

      lblHeading.Text = "Customer details - applicant " + (CurrentApplicantIndex+1);
    }

    private void Gather()
    {
      ctlApplicant.Gather(HttpSession.CurrentApplication.Applicants[CurrentApplicantIndex]);
    }

    private void SetVisibility()
    {
      if (HttpSession.CurrentApplication.Applicants.Length > 1)
      {
        cmdAddApplicant.Visible = false;
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
      if (PostBackControl.ID == "lstValidAddresses") { ctlSmartScroller.Scroll = true; } 
      if (PostBackControl.ID == "cmdResetPostcode") { ctlSmartScroller.Scroll = true; } 
      if (PostBackControl.ID == "cmbResidentFromMonth") { ctlSmartScroller.Scroll = true; } 
      if (PostBackControl.ID == "cmbResidentFromYear") { ctlSmartScroller.Scroll = true; } 

    }

    protected void CmdAddApplicant_Click(object sender, System.EventArgs e)
    {
      if (Page.IsValid)
      {
        HttpSession.CurrentApplication.Applicants = WebServiceHelper.Instance.AddApplicant(HttpSession.CurrentApplication);
        string nextPage = "~/" + WebFormHelper.CurrentPath() + "?applicant=" + HttpSession.CurrentApplication.Applicants.Length;
        Response.Redirect(nextPage);
      }
    }

  }
}
