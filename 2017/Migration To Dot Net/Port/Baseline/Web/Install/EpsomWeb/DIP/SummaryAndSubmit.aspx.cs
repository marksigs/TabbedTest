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

namespace Epsom.Web.Dip
{
  /// <summary>
  /// Summary And Submit Form
  /// </summary>
  public class SummaryAndSubmit : EpsomFormBase
  {
    protected System.Web.UI.WebControls.Label lblGuarantorDetailTitle;
    protected System.Web.UI.WebControls.Repeater rptCommitments;
    protected WebUserControls.ProductSummary ctlProductSummary;
    protected WebUserControls.ApplicantAndGuarantorSummary ctlApplicantAndGuarantorSummary;
    protected System.Web.UI.WebControls.Label lblIntroducerReference;
    protected System.Web.UI.WebControls.Panel pnlIntroducerReference;
    protected System.Web.UI.WebControls.Button cmdEditReference;
    protected System.Web.UI.WebControls.Panel pnlAdviserDetails;

    private void Page_Load(object sender, System.EventArgs e)
    {
      if (!IsPostBack)
      {
        Scatter();
      }

      // Change next button behaviour to "print", and enable "submit" button
      ctlDIPNavigationButtons1.NextPageText = "Print";
      ctlDIPNavigationButtons2.NextPageText = "Print";

      ctlDIPNavigationButtons1.SubmitPageText = "Submit";
      ctlDIPNavigationButtons2.SubmitPageText = "Submit";
      ctlDIPNavigationButtons1.SubmitPage = String.Empty; // ensures visibility
      ctlDIPNavigationButtons2.SubmitPage = String.Empty;

      // Stop navigation
      ctlDIPNavigationButtons1.NextPageStopNavigation = true;
      ctlDIPNavigationButtons2.NextPageStopNavigation = true;

      // Attach client side event to display please wait page
      ctlDIPNavigationButtons1.SubmitPageClientClick = "ProcessSubmit()";
      ctlDIPNavigationButtons2.SubmitPageClientClick = "ProcessSubmit()";

      ctlDIPNavigationButtons1.NextPageClientClick = "javascript:window.print()";
      ctlDIPNavigationButtons2.NextPageClientClick = "javascript:window.print()";

      // If next page button pressed, get decision and redirect to decision page
      if (IsPostBack)
      {
        System.Web.UI.Control PostBackControl;
        PostBackControl  = WebFormHelper.GetPostBackControl(Page);

        if (PostBackControl.ID == "cmdSubmit")
        {
          Epsom.Web.Proxy.DipDecisionResponse dipDecisionResponse = WebServiceHelper.Instance.GetDiPDecision(HttpSession.CurrentApplication, HttpSession.CurrentUser, HttpSession.CurrentFirmIndex);
          HttpSession.DipDecisionResponse = dipDecisionResponse; 

          if (dipDecisionResponse.ResponseMessages != null && dipDecisionResponse.ResponseMessages.Length > 0)
          {
            throw new Exception(dipDecisionResponse.ResponseMessages[0].Text);
          }

          // Redirect to address targeting if necessary
          if (dipDecisionResponse.DipDecision.TARGETINGDATA != null)
          {
            Response.Redirect("~/DIP/SelectValidAddress.aspx");
          }
          else
          {
            Response.Redirect("~/" + this.NextPage);
          }
        }
      }

    }

    protected void RptComponents_ButtonClicked(object sender, System.Web.UI.WebControls.RepeaterCommandEventArgs e)
    {
      if (e == null) {return;}

      switch (e.CommandName)
      {
        case "cmdEditComponent":
          Response.Redirect("ProductSelection.aspx?component=" + (e.Item.ItemIndex + 1));
          break;
        default:
          break;
      }
    }

    private void Scatter()
    {
      // your reference
      if (!Helpers.StringHelper.IsNullOrEmpty(HttpSession.CurrentApplication.IntroducerReference))
      {
        pnlIntroducerReference.Visible= true;
        lblIntroducerReference.Text = HttpSession.CurrentApplication.IntroducerReference;
      }
      else
      {
        pnlIntroducerReference.Visible = false;
      }

      // adviser details scatters itself, but need to control visiblity here
      pnlAdviserDetails.Visible = (HttpSession.CurrentApplication.Adviser != null);

      ctlApplicantAndGuarantorSummary.Scatter(HttpSession.CurrentApplication);
     
    }

    private void cmdEditReference_Click(object sender, EventArgs e)
    {
      Response.Redirect("AdditionalInfo.aspx");
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
      this.cmdEditReference.Click +=new EventHandler(cmdEditReference_Click);
    }
    #endregion

  }
}