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
  /// Product Summary Form.
  /// </summary>
  public class ProductSummary : EpsomFormBase
  {
    protected Epsom.Web.WebUserControls.ProductSummary ctlProductSummary;
    protected System.Web.UI.WebControls.Button cmdAddComponent;

    private void Page_Load(object sender, System.EventArgs e)
    {
      Scatter();
      ctlDIPNavigationButtons1.PreviousPage = null;
      ctlDIPNavigationButtons2.PreviousPage = null;
    }

    private void Scatter()
    {
      //      ctlProductSummary.Scatter(HttpSession.CurrentApplication);
    }

    protected void CmdAddComponent_Click(object sender, System.EventArgs e)
    {
      if (Page.IsValid)
      {
        if (HttpSession.CurrentApplication.Components.Length < Epsom.Web.Helpers.ValidationHelper.MaximumComponents) 
        {
          HttpSession.CurrentApplication.Components = WebServiceHelper.Instance.AddComponent(HttpSession.CurrentApplication);

          string nextPage = Request.Path.Substring(0,Request.Path.LastIndexOf("/")+1) + "ProductSelection.aspx?component=" + HttpSession.CurrentApplication.Components.Length;
          Response.Redirect(nextPage);
        }
      }
    }

    /// <summary>
    /// Custom validator for component total
    /// </summary>
    /// <param name="source"></param>
    /// <param name="e"></param>
    protected void IsComponentTotalValid(object source, System.Web.UI.WebControls.ServerValidateEventArgs e)
    {
      if ( e == null ) { return ; }
      System.Web.UI.Control PostBackControl;
      PostBackControl  = WebFormHelper.GetPostBackControl(Page);
      if (PostBackControl.ID == "cmdAddComponent") 
      {
        e.IsValid = true; // bypass this check if adding a new component
      }
      else
      {
        // compare loanAmountRequested with sum of all components
        e.IsValid = (HttpSession.CurrentApplication.CreditLimit == Epsom.Web.Helpers.ValidationHelper.ComponentSum(HttpSession.CurrentApplication, -1)); 
        if (! e.IsValid)
        {
          System.Web.UI.WebControls.CustomValidator s = (System.Web.UI.WebControls.CustomValidator)source;
          s.ErrorMessage = "Credit Limit must be equal to the total amount allocated to components";
        }
      }
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
      this.cmdAddComponent.Click += new System.EventHandler(this.CmdAddComponent_Click);
    }
    #endregion
  }
}
