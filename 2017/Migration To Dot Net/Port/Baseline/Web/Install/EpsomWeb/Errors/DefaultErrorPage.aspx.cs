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
using System.Diagnostics;

namespace Epsom.Web.Errors
{
  /// <summary>
  /// Summary description for DefaultErrorPage.
  /// </summary>
  public class DefaultErrorPage : System.Web.UI.Page
  {
    protected System.Web.UI.WebControls.Panel pnlEnforcedLogOff;
    protected System.Web.UI.WebControls.Panel pnlLogOffOrRetry;
    protected System.Web.UI.WebControls.Button cmdLogOff;
    protected System.Web.UI.WebControls.Button cmdRetry;

    private void Page_Load(object sender, System.EventArgs e)
    {
      if (Epsom.Web.Helpers.StringHelper.IsNullOrEmpty(Page.Request.QueryString["aspxerrorpath"]))
      {
        pnlLogOffOrRetry.Visible=false;
        ForceLogoff();
      } // all we can do
      else
      {
        // offer option to retry or logoff
        pnlEnforcedLogOff.Visible=false;
      }
    }

    private static void ForceLogoff()
    {
      // Force a logOff.
      // Clear our objects from the HTTP session 
      Epsom.Web.Helpers.HttpSession.Clear(); 
      // Remove forms authentication cookie
      System.Web.Security.FormsAuthentication.SignOut();
    }

    private void cmdLogOff_Click(object sender, EventArgs e)
    {
      ForceLogoff();
      pnlEnforcedLogOff.Visible=true;
      pnlLogOffOrRetry.Visible=false;
    }
    

    private void cmdRetry_Click(object sender, EventArgs e)
    {
      Response.Redirect(Page.Request.QueryString["aspxerrorpath"]);
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
      this.cmdLogOff.Click +=new EventHandler(cmdLogOff_Click);
      this.cmdRetry.Click +=new EventHandler(cmdRetry_Click);
    }
    #endregion

  }
}

