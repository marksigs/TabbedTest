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
	/// Regulation Information form.
	/// </summary>
	public class RegulationInformation : EpsomFormBase
	{
    protected System.Web.UI.WebControls.Repeater rptRegulationInformationCollection;
    protected System.Web.UI.WebControls.Repeater rptMinimumCriteria;

    private Epsom.Web.Proxy.MinimumCriterion[] _MinimumCriterias;

		private void Page_Load(object sender, System.EventArgs e)
		{
      if ( ! IsPostBack )
      {
        _MinimumCriterias = WebServiceHelper.Instance.MinimumCriteria();
        Scatter();
      }
		}

    private void Scatter()
    {

      rptRegulationInformationCollection.DataSource = _MinimumCriterias;
      DataBind();

      for (int i=0; i < _MinimumCriterias.Length; i++)
      {
        // Set Address Title
        System.Web.UI.WebControls.Label lblSummary;
        System.Web.UI.WebControls.Label lblDetail;
        System.Web.UI.WebControls.Label lblLastUpdated;

        lblSummary = (System.Web.UI.WebControls.Label)rptRegulationInformationCollection.Controls[i].FindControl("lblSummary");
        lblDetail = (System.Web.UI.WebControls.Label)rptRegulationInformationCollection.Controls[i].FindControl("lblDetail");
        lblLastUpdated = (System.Web.UI.WebControls.Label)rptRegulationInformationCollection.Controls[i].FindControl("lblLastUpdated");


        lblSummary.Text = _MinimumCriterias[i].Summary ;
        lblDetail.Text = _MinimumCriterias[i].Detail ;
        lblLastUpdated.Text = _MinimumCriterias[i].LastUpdate.ToShortDateString();
      }
    }

    protected void rptRegulationInformationCollection_ButtonClicked(object sender, System.Web.UI.WebControls.RepeaterCommandEventArgs  e)
    {

      if ( e == null) {return ;}

      switch (e.CommandName)
      {
        case "cmdExpand" :
          if (((System.Web.UI.WebControls.Panel)e.Item.FindControl("pnlDetail")).Visible)
          {
            ((System.Web.UI.WebControls.Panel)e.Item.FindControl("pnlDetail")).Visible = false;
            ((System.Web.UI.WebControls.Button)e.Item.FindControl("cmdExpand")).Text = "+";
          }
          else
          {
            ((System.Web.UI.WebControls.Panel)e.Item.FindControl("pnlDetail")).Visible = true;
            ((System.Web.UI.WebControls.Button)e.Item.FindControl("cmdExpand")).Text = "-";
          }
                
          break;
        default :
          break;
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

    }
		#endregion
	}
}
