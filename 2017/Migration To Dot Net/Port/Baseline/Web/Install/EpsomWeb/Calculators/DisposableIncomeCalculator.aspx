<%@ Page language="c#" Codebehind="DisposableIncomeCalculator.aspx.cs" AutoEventWireup="false" Inherits="Epsom.Web.DisposableIncomeCalculator" %>
<%@ Register TagPrefix="validators" NameSpace="Epsom.Web.Validators" Assembly="Epsom.Web"%>
<%@ Register TagPrefix="mp" namespace="Microsoft.Web.Samples.MasterPages" assembly="MasterPages" %>
<%@ Register Src="~/WebUserControls/DIPNavigationButtons.ascx" TagName="DIPNavigationButtons" TagPrefix="Epsom" %>
<%@ Register Src="~/WebUserControls/DIPMenu.ascx" TagName="DIPMenu" TagPrefix="Epsom" %>
<%@ Register TagPrefix = "sstchur" Namespace = "sstchur.web.SmartNav" Assembly = "sstchur.web.SmartNav" %>

<mp:contentcontainer id="Contentcontainer1" runat="server" masterpagefile="~/masterpages/masterpage3.ascx">

  <mp:content id="region2" runat="server">

      <h1>Disposable Income Calculator</h1>

      <h2 class="form_heading">Affordability criteria</h2>  
      
      <asp:validationsummary runat="server" displaymode="BulletList" headertext="Input errors:" CssClass="validationsummary" showsummary="true" id="Validationsummary1" />
      
      <table class="formtable">
        <tr>
          <td class="prompt" style="width:26em">
            <span class="mandatoryfieldasterisk">* </span>
            </span><asp:label id="lblAnnualIncome" runat="server" text="Total annual guaranteed gross income" />
          </td>
          <td>
            <asp:textbox id="txtAnnualIncome" runat="server" columns="8" maxlength="7" tooltip="enter your total annual guaranteed gross income from all sources"></asp:textbox>
            <asp:RequiredFieldValidator id=RequiredfieldvalidatorAnnualIncome Runat="Server" ErrorMessage="Please enter your guaranteed gross annual income" ControlToValidate="txtAnnualIncome">***</asp:RequiredFieldValidator>
            <asp:CompareValidator id="CompareValidatorAnnualIncome" Runat="server" Operator="DataTypeCheck" Type="Currency" ControlToValidate="txtAnnualIncome" ErrorMessage="Gross annual income is not numeric">***</asp:CompareValidator>
          </td>                
        </tr>
        
        <tr>
          <td class="prompt" style="width:26em">
            <asp:label id="lblIncomeNonGuaranteed" runat="server" text="Non guaranteed/irregular income" />
          </td>
          <td>
            <asp:textbox id="txtIncomeNonGuaranteed" runat="server" columns="8" maxlength="7" tooltip="enter any income that is from non guaranteed, irregular or investment sources" />
            <asp:CompareValidator id="ComparevalidatorIncomeNonGuaranteed" Runat="server" Operator="DataTypeCheck" Type="Currency" ControlToValidate="txtIncomeNonGuaranteed" ErrorMessage="Non guaranteed income is not numeric">***</asp:CompareValidator>
          </td>
        </tr>
        
        <tr>
          <td class="prompt" style="width:26em">
            <asp:label id="lblMortgageBalance" runat="server" text="Total outstanding balance on all existing mortgages" />
          </td>
          <td>
            <asp:textbox id="txtMortgageBalance" runat="server" columns="8" maxlength="8" tooltip="enter the total current outstanding balance on all existing mortgages from any lender" />
            <asp:CompareValidator id="ComparevalidatorMortgageBalance" Runat="server" Operator="DataTypeCheck" Type="Currency" ControlToValidate="txtMortgageBalance" ErrorMessage="Outstanding mortgage balance is not numeric">***</asp:CompareValidator>
          </td>
        </tr>
        
        <tr>
          <td class="prompt" style="width:26em">
            <asp:label id="lblLoansMonthlyPayment" runat="server" text="Total monthly repayment on all existing other loans" />
          </td>
          <td>
            <asp:textbox id="txtLoansMonthlyPayment" runat="server" columns="8" maxlength="5" tooltip="enter the total monthly repayment on all existing non-mortgage loans" />
            <asp:CompareValidator id="ComparevalidatorLoansMonthlyPayment" Runat="server" Operator="DataTypeCheck" Type="Currency" ControlToValidate="txtLoansMonthlyPayment" ErrorMessage="Monthly repayment of other loans is not numeric">***</asp:CompareValidator>
          </td>      
        </tr>
        
        <tr>
          <td class="prompt" style="width:26em">
            <asp:label id="lblCreditCardBalance" runat="server" text="Total outstanding balance on all credit/store cards" />
          </td>
          <td>
            <asp:textbox id="txtCreditCardBalance" runat="server" columns="8" maxlength="6" tooltip="enter the total current outstanding balance on all credit cards and store cards" />
            <asp:CompareValidator id="ComparevalidatorCreditCardBalance" Runat="server" Operator="DataTypeCheck" Type="Currency" ControlToValidate="txtCreditCardBalance" ErrorMessage="Outstanding balance of credit/store cards is not numeric">***</asp:CompareValidator>
          </td>
        </tr>
        
        <tr>
          <td class="prompt" style="width:26em">
            <asp:label id="lblNoOfDependants" runat="server" text="Number of dependants" />
          </td>
          <td>
            <asp:textbox id="txtNoOfDependants" runat="server" columns="2" maxlength="2" tooltip="enter the total number of dependants" />
            <asp:CompareValidator id="ComparevalidatorNoOfDependants" Runat="server" Operator="DataTypeCheck" Type="Currency" ControlToValidate="txtNoOfDependants" ErrorMessage="Number of dependants is not numeric">***</asp:CompareValidator>
          </td>
        </tr>

		<% /* Peter Edney - 30/01/2007 */ %>
		<% /* Added term of loan */ %>
        <tr>
          <td class="prompt" style="width:26em">
			<span class="mandatoryfieldasterisk">* </span>
            <asp:label id="lblTermofLoan" runat="server" text="Term of Loan (years)" />
          </td>
          <td>
            <asp:textbox id="txtTermofLoan" runat="server" columns="2" maxlength="2" tooltip="enter the term of the required loan in years (minimum=5, maximum=35" />
            <asp:RequiredFieldValidator id="RequiredfieldValidatorTermofLoan" Runat="Server" ErrorMessage="Please enter the term of the required loan in years (minimum=5, maximum=35" ControlToValidate="txtTermofLoan">***</asp:RequiredFieldValidator>
            <asp:CompareValidator id="ComparevalidatorTermofLoan" Runat="server" Operator="DataTypeCheck" Type="Integer" ControlToValidate="txtTermofLoan" ErrorMessage="Term of loan is not numeric">***</asp:CompareValidator>
            <asp:RangeValidator id="RangevalidatorTermofLoan" Runat="server" MinimumValue="5" MaximumValue="35" Type="Integer" ControlToValidate="txtTermofLoan" ErrorMessage="Term of loan must be from 5 to 35 years">***</asp:RangeValidator>
          </td>
        </tr>

		<% /* Peter Edney - 30/01/2007 */ %>
		<% /* Added Monthly Cost of Repayment Vehicle */ %>
        <tr>
          <td class="prompt" style="width:26em">
            <asp:label id="lblMonthlyCostVehicle" runat="server" text="Monthly Cost of Repayment Vehicle (Interest-Only Loans)" />
          </td>
          <td>
            <asp:textbox id="txtMonthlyCostVehicle" runat="server" columns="8" maxlength="5" tooltip="enter the monthly cost of the savings/investment plan that will be used to repay the capital of an interest-only loan" />
            <asp:CompareValidator id="ComparevalidatorMonthlyCostVehicle" Runat="server" Operator="DataTypeCheck" Type="Currency" ControlToValidate="txtMonthlyCostVehicle" ErrorMessage="Monthly Cost of repayment vehicle is not numeric">***</asp:CompareValidator>
          </td>      
        </tr>

        <tr>
          <td class="prompt"></td>
          <td>
			<asp:button id="cmdCalculate" runat="server" text="Calculate"></asp:button>
			<asp:button id="cmdReset" runat="server" text="Reset"></asp:button>
          </td>
        </tr>
      </table>

      <asp:panel id="pnlDisposableIncome" runat="server">

		<% /* Peter Edney - 30/01/2007 */ %>
		<% /* Added table to display results so they are formatted the same as the input fields */ %>
		<table class="formtable">
	
			<tr>
				<td class="prompt" style="width:26em">
					<asp:label id="lblAnnualDisposableIncome" runat="server" text="Annual disposable income" />
				</td>
				<td>
					<asp:textbox id="txtAnnualDisposableIncomeResult" runat="server" columns="8" maxlength="5" ReadOnly="True" />
				</td>                
			</tr>

			<tr>
				<td class="prompt" style="width:26em">
					<asp:label id="lblMaxLoanCapitalInterest" runat="server" text="Maximum Loan (Capital and Interest)" />
				</td>
				<td>
					<asp:textbox id="txtMaxLoanCapitalInterest" runat="server" columns="8" maxlength="5" ReadOnly="True" />
				</td>                
			</tr>

			<tr>
				<td class="prompt" style="width:26em">
					<asp:label id="lblMaxLoanInterest" runat="server" text="Maximum Loan (Interest only)" />
				</td>
				<td>
					<asp:textbox id="txtMaxLoanInterest" runat="server" columns="8" maxlength="5" ReadOnly="True" />
				</td>                
			</tr>

			<tr>
				<td class="prompt" style="width:26em">
					<asp:label id="lblDisclaimer" runat="server" text="" />
				</td>
				<td>
					<asp:label id="lblDisclaimerText" runat="server" text="These figures are indicative based on the information entered and do not represent a commitment by db mortgages to lend the amounts shown" />
				</td>                
			</tr>
	
		</table>
      
      </asp:panel>
     
      <asp:panel id="pnlExclusions" runat="server">

        <h2 class="form_heading">Exclusions</h2>  

        <p>
          Do not include any of the following:
          <ul>
            <li>Existing mortgages and other loans with 12 or fewer months left to run</li>
            <li>Existing buy-to-let mortgages from other lenders (i.e. not dbm)</li>
            <li>Existing dbm buy-to-let (rental-based) mortgages</li>
          </ul>
        </p>

      </asp:panel>

      <sstchur:SmartScroller id="ctlSmartScroller" runat="server" />

  </mp:content>

</mp:contentcontainer>
