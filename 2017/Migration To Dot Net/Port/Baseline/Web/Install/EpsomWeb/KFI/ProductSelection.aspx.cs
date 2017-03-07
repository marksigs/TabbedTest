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
  /// Product Selection form.
  /// </summary>
  public class ProductSelection : EpsomFormBase
  {
    protected Epsom.Web.WebUserControls.ProductSelection ctlProductSelection;
    protected System.Web.UI.WebControls.ValidationSummary Validationsummary1;
    protected sstchur.web.SmartNav.SmartScroller ctlSmartScroller;

    private int CurrentComponentIndex
    {
      get 
      { 
        return WebFormHelper.CurrentIndex(Request,"component");
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
    }

    private void Page_PreRender(object sender, System.EventArgs e)
    {
      SetVisibility();
    }

    private void SetVisibility()
    {
      ctlProductSelection.SetVisibility();
    }

    private void Scatter()
    {
      ctlProductSelection.Scatter(HttpSession.CurrentApplication, CurrentComponentIndex);
    }

    private void Gather()
    {
      HttpSession.CurrentApplication = ctlProductSelection.Gather(HttpSession.CurrentApplication, CurrentComponentIndex);
    }

    private void SetSmartScroller()
    {
      System.Web.UI.Control PostBackControl;
      PostBackControl  = WebFormHelper.GetPostBackControl(Page);
      ctlProductSelection.SetSmartScroller(PostBackControl.ID, ctlSmartScroller);
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
    }
    #endregion

  }
}
