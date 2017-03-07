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
  /// Minimum criteria for the loan.
  /// </summary>
  public class MinimumCriteria : EpsomFormBase
  {
    protected Epsom.Web.WebUserControls.MinimumCriteria ctlMinimumCriteria;
    protected sstchur.web.SmartNav.SmartScroller ctlSmartScroller;
 
    private void Page_Load(object sender, System.EventArgs e)
    {
      if (!IsPostBack)
      {
        ctlMinimumCriteria.Scatter();
      }
      else
      {
        System.Web.UI.Control PostBackControl;
        PostBackControl  = WebFormHelper.GetPostBackControl(Page);
        ctlMinimumCriteria.SetSmartScroller(ctlSmartScroller,PostBackControl);
        ctlMinimumCriteria.Gather();
      }
      ctlMinimumCriteria.SetVisibility();
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
    }
    #endregion
  }
}
