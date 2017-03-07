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
  /// AdviserDetails form
  /// </summary>
  public class AdviserDetails : EpsomFormBase
  {
    protected Epsom.Web.WebUserControls.Adviser ctlAdviser;

    private void Page_Load(object sender, System.EventArgs e)
    {
      Control control = null;
      if (IsPostBack) 
      {
        HttpSession.CurrentApplication.Adviser = ctlAdviser.Gather();
       
        string ctrlname = Page.Request.Params.Get("__EVENTTARGET");
        if (ctrlname != null && ctrlname != string.Empty)
        {
          control = Page.FindControl(ctrlname);
        }
        else
        {
          foreach (string ctl in Page.Request.Form)
          {
            Control c = Page.FindControl(ctl);
            if (c is System.Web.UI.WebControls.Button)
              {
                control = c;
                break;
              }
            }
          }
      

      
      }
      else
      {
        ctlAdviser.Scatter(HttpSession.CurrentApplication.Adviser);
      }
      
      
      
      if ( control ==  null )
      {
        return;
      }
      
      ctlAdviser.SetVisibility();  
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
