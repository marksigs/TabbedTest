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

namespace Epsom.Web.Kfi
{
  /// <summary>
  /// Submission Route Form.
  /// </summary>
  public class LoanRequirements : EpsomFormBase
  {
    protected Epsom.Web.WebUserControls.LoanRequirements ctlLoanRequirements;
    protected sstchur.web.SmartNav.SmartScroller ctlSmartScroller;
    protected System.Web.UI.WebControls.ValidationSummary Validationsummary1;

    private void Page_Load(object sender, System.EventArgs e)
    {
      if ( IsPostBack ) 
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

    private void Gather()
    {
      HttpSession.CurrentApplication = ctlLoanRequirements.Gather(HttpSession.CurrentApplication, HttpSession.CurrentUser);
    }

    private void Scatter()
    {
      ctlLoanRequirements.Scatter(HttpSession.CurrentApplication, HttpSession.CurrentUser);
    }

    private void SetSmartScroller()
    {
      System.Web.UI.Control PostBackControl;
      PostBackControl  = WebFormHelper.GetPostBackControl(Page);

      ctlLoanRequirements.SetSmartScroller(PostBackControl.ID, ctlSmartScroller);
    }

    private void SetVisibility()
    {
      ctlLoanRequirements.SetVisibility();
    }

    private void ctlLoanRequirements_SchemeChanged(object sender, EventArgs e)
    {
      base.RefreshMenu();
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
      this.ctlLoanRequirements.SchemeChanged +=new EventHandler(ctlLoanRequirements_SchemeChanged);
		}
		#endregion
	}
}
