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
	/// Loan details form.
	/// </summary>
	public class LoanDetails : EpsomFormBase
	{
    protected System.Web.UI.WebControls.ValidationSummary Validationsummary1;
    protected System.Web.UI.WebControls.DropDownList cmbLoanType;
    protected System.Web.UI.WebControls.TextBox txtCreditLimit;
    protected System.Web.UI.WebControls.TextBox txtInitialAmount;
    protected System.Web.UI.WebControls.TextBox txtPropertyPrice;
    protected System.Web.UI.WebControls.Button cmdCalculate;
    protected System.Web.UI.WebControls.Label lblLoanToValueRatio;
    protected System.Web.UI.WebControls.DropDownList cmbTermOfMortgage;
    protected System.Web.UI.WebControls.DropDownList cmbRepaymentMethod;
    protected System.Web.UI.WebControls.TextBox txtInterestOnly;
    protected System.Web.UI.WebControls.TextBox txtCapitalRepayment;
    protected System.Web.UI.WebControls.Panel pnlPartAndPartSplit;
    protected System.Web.UI.WebControls.TextBox txtDiscountedPropertyPrice;
    protected System.Web.UI.WebControls.Panel pnlDiscountedPropertyPrice;
    protected System.Web.UI.WebControls.RadioButtonList rblFinalRepaymentSource;
    protected System.Web.UI.WebControls.Panel pnlFinalRepaymentSource;
    protected System.Web.UI.WebControls.DropDownList cmbDepositSource;
    protected System.Web.UI.WebControls.Panel pnlDepositSource;
    protected System.Web.UI.WebControls.DropDownList cmbRemortgagePurpose;
    protected System.Web.UI.WebControls.TextBox txtOutstandingMortgageBalance;
    protected System.Web.UI.WebControls.Panel pnlRemortgageDetails;
    protected sstchur.web.SmartNav.SmartScroller  ctlSmartScroller;

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
      SetVisibility();
		}

    private void SetSmartScroller()
    {
      System.Web.UI.Control PostBackControl;
      PostBackControl  = WebFormHelper.GetPostBackControl(Page);
      if (PostBackControl.ID == "cmbRepaymentMethod") { ctlSmartScroller.Scroll = true; }
      if (PostBackControl.ID == "cmbLoanType") { ctlSmartScroller.Scroll = true; }
      if (PostBackControl.ID == "cmdCalculate") { ctlSmartScroller.Scroll = true; }
    }

    private void Scatter()
    {
//      PopulateCombos();
//
//      cmbLoanType.SelectedIndex = HttpSession.CurrentApplication.Products[CurrentProductIndex].LoanDetails.LoanType;
//      txtCreditLimit.Text = HttpSession.CurrentApplication.Products[CurrentProductIndex].LoanDetails.CreditLimit.ToString(System.Globalization.CultureInfo.CurrentCulture) ;
//      txtInitialAmount.Text = HttpSession.CurrentApplication.Products[CurrentProductIndex].LoanDetails.LoanAmount.ToString(System.Globalization.CultureInfo.CurrentCulture);
//      txtPropertyPrice.Text = HttpSession.CurrentApplication.Products[CurrentProductIndex].LoanDetails.PropertyPrice.ToString(System.Globalization.CultureInfo.CurrentCulture);
//      lblLoanToValueRatio.Text = HttpSession.CurrentApplication.Products[CurrentProductIndex].LoanDetails.LoanToValue.ToString(System.Globalization.CultureInfo.CurrentCulture);
//      cmbTermOfMortgage.SelectedValue = HttpSession.CurrentApplication.Products[CurrentProductIndex].LoanDetails.TermOfMortgage;
//      cmbRepaymentMethod.SelectedValue = HttpSession.CurrentApplication.Products[CurrentProductIndex].LoanDetails.RepaymentMethod;
//      txtInterestOnly.Text = HttpSession.CurrentApplication.Products[CurrentProductIndex].LoanDetails.InterestOnly.ToString(System.Globalization.CultureInfo.CurrentCulture);
//      txtCapitalRepayment.Text = HttpSession.CurrentApplication.Products[CurrentProductIndex].LoanDetails.CapitalRepayment.ToString(System.Globalization.CultureInfo.CurrentCulture);

    }

    private void Gather()
    {
      // ******************************************************************
      // Consider using Epsom.Web.Helpers.NumberHelper.ToDecimal instead
      // ********************************************************************
//      HttpSession.CurrentApplication.Products[CurrentProductIndex].LoanDetails.LoanType = cmbLoanType.SelectedIndex ;
//      HttpSession.CurrentApplication.Products[CurrentProductIndex].LoanDetails.CreditLimit = Convert.ToDecimal(txtCreditLimit.Text, System.Globalization.CultureInfo.CurrentCulture) ;
//      HttpSession.CurrentApplication.Products[CurrentProductIndex].LoanDetails.LoanAmount = Convert.ToDecimal(txtInitialAmount.Text, System.Globalization.CultureInfo.CurrentCulture) ;
//      HttpSession.CurrentApplication.Products[CurrentProductIndex].LoanDetails.PropertyPrice = Convert.ToDecimal(txtPropertyPrice.Text, System.Globalization.CultureInfo.CurrentCulture) ;
//      HttpSession.CurrentApplication.Products[CurrentProductIndex].LoanDetails.LoanToValue = Convert.ToDecimal(lblLoanToValueRatio.Text, System.Globalization.CultureInfo.CurrentCulture) ;
//      HttpSession.CurrentApplication.Products[CurrentProductIndex].LoanDetails.TermOfMortgage = cmbTermOfMortgage.SelectedItem.Value ;
//      HttpSession.CurrentApplication.Products[CurrentProductIndex].LoanDetails.RepaymentMethod = cmbRepaymentMethod.SelectedItem.Value ;
//      HttpSession.CurrentApplication.Products[CurrentProductIndex].LoanDetails.InterestOnly = Convert.ToDecimal(txtInterestOnly.Text, System.Globalization.CultureInfo.CurrentCulture) ;
//      HttpSession.CurrentApplication.Products[CurrentProductIndex].LoanDetails.CapitalRepayment = Convert.ToDecimal(txtCapitalRepayment.Text, System.Globalization.CultureInfo.CurrentCulture ) ;
    }

    private void SetVisibility()
    {
//      if (HttpSession.CurrentApplication.Products[CurrentProductIndex].RightToBuy)
//      {
//        pnlDiscountedPropertyPrice.Visible = true;
//      }
//      else
//      {
//        pnlDiscountedPropertyPrice.Visible = false;
//      }
//
//      if (HttpSession.CurrentApplication.Products[CurrentProductIndex].LoanDetails.RepaymentMethod == "3")
//      {
//        pnlPartAndPartSplit.Visible = true;
//      }
//      else
//      {
//        pnlPartAndPartSplit.Visible = false;
//      }
//
//      if (HttpSession.CurrentApplication.Products[CurrentProductIndex].LoanDetails.RepaymentMethod == "3" || HttpSession.CurrentApplication.Products[CurrentProductIndex].LoanDetails.RepaymentMethod == "1")
//      {
//        pnlFinalRepaymentSource.Visible = true;
//      }
//      else
//      {
//        pnlFinalRepaymentSource.Visible = false;
//      }
//
//      if (HttpSession.CurrentApplication.Products[CurrentProductIndex].LoanDetails.LoanType != 3)
//      {
//        pnlDepositSource.Visible = true;
//      }
//      else
//      {
//        pnlDepositSource.Visible = false;
//      }
//
//      if (HttpSession.CurrentApplication.Products[CurrentProductIndex].LoanDetails.LoanType == 3)
//      {
//        pnlRemortgageDetails.Visible = true;
//      }
//      else
//      {
//        pnlRemortgageDetails.Visible = false;
//      }
    }

    private void PopulateCombos()
    {
      //WebFormHelper.PopulateListControl(cmbLoanType, WebServiceHelper.Instance.LoanTypes());
      WebFormHelper.PopulateListControl(cmbRepaymentMethod, Web.Helpers.ComboValues.ComboGroup(Web.Helpers.ComboValues.ComboGroupType.RepaymentType ), true);

      cmbTermOfMortgage.Items.Clear();
      for (int i = 5; i <= 25; ++i)
      {
        cmbTermOfMortgage.Items.Add(i.ToString(System.Globalization.CultureInfo.CurrentCulture));
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
