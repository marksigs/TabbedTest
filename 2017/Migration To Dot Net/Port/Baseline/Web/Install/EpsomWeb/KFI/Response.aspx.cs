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
	/// Decision Response form.
	/// </summary>
	public class Response : EpsomFormBase
	{
    protected System.Web.UI.WebControls.Panel pnlResponse;
    protected System.Web.UI.WebControls.LinkButton cmdDownloadDocumentImageLink;
    protected System.Web.UI.WebControls.LinkButton cmdDownloadDocumentTextLink;
    protected System.Web.UI.WebControls.Button cmdCancel;
    protected System.Web.UI.WebControls.Button cmdPrePopulateKfi;
    protected System.Web.UI.WebControls.Button cmdPrePopulateDip;

    protected System.Web.UI.WebControls.Label lblMonthlyCost;

    private void Page_Load(object sender, System.EventArgs e)
		{

      Scatter() ;

      //HttpSession.CreateKfiResponse = null;
    }

    private void Scatter()
    {
      lblMonthlyCost.Text = HttpSession.CreateKfiResponse.MonthlyCosts.ToString();
    }

    private void cmdDownloadDocument_Click(object sender, EventArgs e)
    {
      byte[] document = HttpSession.CreateKfiResponse.Document;

      Response.Clear();

      Response.AddHeader("Content-Disposition", "attachment; filename=kfi.pdf");
      Response.AppendHeader("Cache-Control","cache");
      Response.AppendHeader("Cache-Control","must-revalidate");
      Response.AppendHeader("Pragma","public");
      Response.AddHeader("Content-Length", document.Length.ToString());
      Response.ContentType = "application/pdf";

//      Response.AddHeader("Content-Disposition", "inline; filename=kfi.pdf");
//      Response.AddHeader("Content-Length", document.Length.ToString());
//      Response.ContentType = "application/pdf";

      Response.BinaryWrite(document);
      Response.Flush(); 
      Response.End();
    }



    private void cmdPrintKFI_Click(object sender, EventArgs e)
    {

    }


    private void cmdCancel_Click(object sender, EventArgs e)
    {
      Response.Redirect("~/Default.aspx");
    }

    private void cmdPrePopulateDip_Click(object sender, EventArgs e)
    {
      Response.Redirect("~/DIP/Default.aspx");
    }

    private void cmdPrePopulateKfi_Click(object sender, EventArgs e)
    {
      Response.Redirect("~/KFI/Default.aspx");
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
      this.cmdDownloadDocumentTextLink.Click += new EventHandler(cmdDownloadDocument_Click);
      this.cmdDownloadDocumentImageLink.Click += new EventHandler(cmdDownloadDocument_Click);
      this.cmdCancel.Click += new EventHandler(cmdCancel_Click);
      this.cmdPrePopulateDip.Click += new EventHandler(cmdPrePopulateDip_Click);
      this.cmdPrePopulateKfi.Click += new EventHandler(cmdPrintKFI_Click);

    }
    #endregion

  }
}
