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
	/// Insurance and Investment form.
	/// </summary>
	public class InsuranceAndInvestment : EpsomFormBase
	{
    protected System.Web.UI.WebControls.Panel pnlInsuranceInvestments; 
    protected System.Web.UI.WebControls.RadioButton rdoBuildingYes;
    protected System.Web.UI.WebControls.RadioButton rdoBuildingNo;
    protected System.Web.UI.WebControls.TextBox txtBuildingsCompany;
    protected System.Web.UI.WebControls.TextBox txtBuildingsPayment;
    protected System.Web.UI.WebControls.RadioButton rdoContentsYes;
    protected System.Web.UI.WebControls.RadioButton rdoContentsNo;
    protected System.Web.UI.WebControls.TextBox txtContentsCompany;
    protected System.Web.UI.WebControls.TextBox txtContentsPayment;
    protected System.Web.UI.WebControls.RadioButton rdoPaymentProtectionYes;
    protected System.Web.UI.WebControls.RadioButton rdoPaymentProtectionNo;
    protected System.Web.UI.WebControls.TextBox txtPaymentProtectionCompany;
    protected System.Web.UI.WebControls.TextBox txtPaymentProtectionPayment;
    
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

    private void InsuranceAndInvestment_PreRender(object sender, EventArgs e)
    {
      SetVisibility();
    }

    private void Scatter()
    {
      rdoBuildingYes.Checked = HttpSession.CurrentApplication.Insurances[0].HaveInsurance ;
      rdoBuildingNo.Checked = ! HttpSession.CurrentApplication.Insurances[0].HaveInsurance ;
      txtBuildingsCompany.Text = HttpSession.CurrentApplication.Insurances[0].Company;
      txtBuildingsPayment.Text = HttpSession.CurrentApplication.Insurances[0].MonthlyPayment.ToString() ;

      rdoContentsYes.Checked = HttpSession.CurrentApplication.Insurances[1].HaveInsurance ;
      rdoContentsNo.Checked = ! HttpSession.CurrentApplication.Insurances[1].HaveInsurance ;
      txtContentsCompany.Text = HttpSession.CurrentApplication.Insurances[1].Company;
      txtContentsPayment.Text = HttpSession.CurrentApplication.Insurances[1].MonthlyPayment.ToString() ;

      rdoPaymentProtectionYes.Checked = HttpSession.CurrentApplication.Insurances[2].HaveInsurance ;
      rdoPaymentProtectionNo.Checked = ! HttpSession.CurrentApplication.Insurances[2].HaveInsurance ;
      txtPaymentProtectionCompany.Text = HttpSession.CurrentApplication.Insurances[2].Company;
      txtPaymentProtectionPayment.Text = HttpSession.CurrentApplication.Insurances[2].MonthlyPayment.ToString() ;
    }

    private void Gather()
    {
      HttpSession.CurrentApplication.Insurances[0].HaveInsurance = rdoBuildingYes.Checked ;
      HttpSession.CurrentApplication.Insurances[0].Company = txtBuildingsCompany.Text;
      HttpSession.CurrentApplication.Insurances[0].MonthlyPayment = Helpers.NumberHelper.ToDouble(txtBuildingsPayment.Text);

      HttpSession.CurrentApplication.Insurances[1].HaveInsurance = rdoContentsYes.Checked ;
      HttpSession.CurrentApplication.Insurances[1].Company = txtContentsCompany.Text;
      HttpSession.CurrentApplication.Insurances[1].MonthlyPayment = Helpers.NumberHelper.ToDouble(txtContentsPayment.Text);

      HttpSession.CurrentApplication.Insurances[2].HaveInsurance = rdoPaymentProtectionYes.Checked ;
      HttpSession.CurrentApplication.Insurances[2].Company = txtPaymentProtectionCompany.Text;
      HttpSession.CurrentApplication.Insurances[2].MonthlyPayment = Helpers.NumberHelper.ToDouble(txtPaymentProtectionPayment.Text);
    }

    private void SetSmartScroller()
    {

    }

    private void SetVisibility()
    {

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
      this.PreRender += new EventHandler(InsuranceAndInvestment_PreRender);

    }
    #endregion

   
  }
}
