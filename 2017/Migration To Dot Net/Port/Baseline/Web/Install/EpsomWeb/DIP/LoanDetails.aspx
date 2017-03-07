<%@ Page language="c#" Codebehind="LoanDetails.aspx.cs" AutoEventWireup="false" Inherits="Epsom.Web.Dip.LoanDetails" %>
<%@ Register TagPrefix="mp" namespace="Microsoft.Web.Samples.MasterPages" assembly="MasterPages" %>
<%@ Register Src="~/WebUserControls/DIPNavigationButtons.ascx" TagName="DIPNavigationButtons" TagPrefix="Epsom" %>
<%@ Register Src="~/WebUserControls/DIPMenu.ascx" TagName="DIPMenu" TagPrefix="Epsom" %>
<%@ Register TagPrefix = "sstchur" Namespace = "sstchur.web.SmartNav" Assembly = "sstchur.web.SmartNav" %>

<mp:contentcontainer id="Contentcontainer1" runat="server" masterpagefile="../masterpages/masterpage2.ascx">

  <mp:content id="region1" runat="server">

    <epsom:dipmenu id="DipMenu" runat="server" />

  </mp:content>

  <mp:content id="region2" runat="server">

    <h1>Loan details</h1>

    <epsom:dipnavigationbuttons id="ctlDIPNavigationButtons1" runat="server" />

    <asp:validationsummary runat="server" displaymode="BulletList" headertext="Input errors:" CssClass="validationsummary" showsummary="true" id="Validationsummary1" />

    <table class="formtable">
	    <tr>
		    <td class="prompt">Loan type</td>
		    <td><asp:dropdownlist id="cmbLoanType" runat="server" AutoPostBack="true"/></td>
	    </tr>
	    <tr>
		    <td class="prompt">Credit limit</td>
		    <td><asp:TextBox id="txtCreditLimit" runat="server" /></td>
	    </tr>
	    <tr>
		    <td class="prompt">Initial amount</td>
		    <td><asp:TextBox id="txtInitialAmount" runat="server" /></td>
	    </tr>
	    
	    <asp:Panel ID="pnlDiscountedPropertyPrice" runat="server">
	    
	      <tr>
		      <td class="prompt">Discounted property price</td>
		      <td><asp:TextBox id="txtDiscountedPropertyPrice" runat="server" /></td>
	      </tr>
	    
	    </asp:Panel>
	    
	    <tr>
		    <td class="prompt">Property price/ Estimated value</td>
		    <td><asp:TextBox id="txtPropertyPrice" runat="server" /></td>
	    </tr>
	    <tr>
		    <td class="prompt"></td>
		    <td><asp:Button id="cmdCalculate" runat="server" text="Calculate" /></td>
	    </tr>
	    <tr>
		    <td class="prompt">Loan to value ratio</td>
		    <td><asp:Label id="lblLoanToValueRatio" runat="server" /></td>
	    </tr>
	    <tr>
		    <td class="prompt">Term of mortgage</td>
		    <td><asp:DropDownList id="cmbTermOfMortgage" runat="server" /></td>
	    </tr>
	    <tr>
		    <td class="prompt">Repayment Method</td>
		    <td><asp:DropDownList id="cmbRepaymentMethod" runat="server" AutoPostBack="true"/></td>
	    </tr>
	    
	    <asp:Panel ID="pnlPartAndPartSplit" runat="server">
	    
	      <tr>
		      <td class="prompt">Interest only</td>
		      <td><asp:TextBox id="txtInterestOnly" runat="server" /></td>
	      </tr>
	      <tr>
		      <td class="prompt">Capital repayment</td>
		      <td><asp:TextBox id="txtCapitalRepayment" runat="server" /></td>
	      </tr>

	    </asp:Panel>
	    
	    <asp:Panel ID="pnlFinalRepaymentSource" runat="server">
	      
	      <tr>
		      <td class="prompt">Final repayment source</td>
		      <td>
		      
		      <asp:RadioButtonList id=rblFinalRepaymentSource runat="server">
            <asp:ListItem Value="1">ISA, Pension, Endowment, Other</asp:ListItem>
            <asp:ListItem Value="2">Sale of Residential Property, Inheritance, Sale of Business</asp:ListItem>
          </asp:RadioButtonList>
		      
		      </td>
	      </tr>
 
	    </asp:Panel>
	    
	    <asp:Panel ID="pnlDepositSource" runat="server" >
	    
	      <tr>
		      <td class="prompt">Deposit Source</td>
		      <td><asp:DropDownList id="cmbDepositSource" runat="server" /></td>
	      </tr>
	    
	    </asp:Panel>
	    
	     <asp:Panel ID="pnlRemortgageDetails" runat="server" >
	    
	      <tr>
		      <td class="prompt">Remortgage Purpose</td>
		      <td><asp:DropDownList id="cmbRemortgagePurpose" runat="server" /></td>
		    </tr>
		    <tr>
		      <td class="prompt">Outstanding mortgage balance</td>
		      <td><asp:TextBox id="txtOutstandingMortgageBalance" runat="server" /></td>
	      </tr>
	    
	    </asp:Panel>
	    
	  </table>

    <epsom:dipnavigationbuttons id="ctlDIPNavigationButtons2" runat="server" />

    <sstchur:SmartScroller id="ctlSmartScroller" runat="server" />

  </mp:content>

</mp:contentcontainer>

