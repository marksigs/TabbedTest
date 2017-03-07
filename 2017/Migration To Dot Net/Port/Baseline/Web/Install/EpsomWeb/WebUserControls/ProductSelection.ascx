<%@ Control Language="c#" AutoEventWireup="false" Codebehind="ProductSelection.ascx.cs" Inherits="Epsom.Web.WebUserControls.ProductSelection" TargetSchema="http://schemas.microsoft.com/intellisense/ie5"%>
<%@ Register Src="~/WebUserControls/LoanDetailsSummary.ascx" TagName="LoanDetailsSummary" TagPrefix="Epsom" %>
<%@ Register Src="~/WebUserControls/SubmissionRouteSummary.ascx" TagName="SubmissionRouteSummary" TagPrefix="Epsom" %>
<%@ Register TagPrefix="validators" NameSpace="Epsom.Web.Validators" Assembly="Epsom.Web"%>
<%@ Register TagPrefix="CMS" namespace="Epsom.Web.Helpers" assembly="Epsom.Web" %>
    
    <cms:cmsBoilerplate cmsref="40" runat="server" />
    
    <h2 class="form_heading">Main loan details</h2>
    
    <epsom:loandetailssummary id="ctlLoanDetailsSummary" runat="server" />

    <epsom:submissionroutesummary id="ctlSubmissionRouteSummary" runat="server" />
    
    <h2 class="form_heading">Product component selection</h2>
   
    <fieldset>
    
      <legend>Component loan details</legend>
    
      <table class="formtable">
        <tr>
          <td class="prompt"><span class="mandatoryfieldasterisk">*</span>Primary purpose of loan</td>
          <td>
            <asp:dropdownlist id="cmbPurposeOfLoan" runat="server" autopostback="true"/>
            <validators:requiredselecteditemvalidator id="RequiredSelectedItemValidator1" runat="server" enableclientscript="false" text="***" 
              errormessage="Purpose of Loan is required" controltovalidate="cmbPurposeOfLoan"></validators:requiredselecteditemvalidator>
          </td>
        </tr>
        <tr>
          <td class="prompt"><span class="mandatoryfieldasterisk">*</span>Secondary purpose of loan</td>
          <td>
            <asp:dropdownlist id="cmbSubPurposeOfLoan" runat="server" />
            <validators:requiredselecteditemvalidator id="Requiredselecteditemvalidator2" runat="server" enableclientscript="false" text="***" 
              errormessage="Sub-purpose of Loan is required" controltovalidate="cmbSubPurposeOfLoan"></validators:requiredselecteditemvalidator>
          </td>
        </tr>
        <tr>
          <td class="prompt"><span class="mandatoryfieldasterisk">*</span>Repayment type</td>
          <td>
            <asp:dropdownlist id="cmbRepaymentType" runat="server" autopostback="true"/>
            <validators:requiredselecteditemvalidator id="Requiredselecteditemvalidator3" runat="server" enableclientscript="false" text="***" 
              errormessage="Repayment Type is required" controltovalidate="cmbRepaymentType"></validators:requiredselecteditemvalidator>
          </td>
        </tr>
        <tr>
          <td class="prompt"><span class="mandatoryfieldasterisk">*</span>Component Amount</td>
          <td>
            <asp:textbox id="txtComponentAmount" runat="server" CssClass="text" />
            <asp:requiredfieldvalidator id="Requiredfieldvalidator1" runat="server" enableclientscript="false" 
              errormessage="Component amount is required" controltovalidate="txtComponentAmount" text="***"></asp:requiredfieldvalidator>
            <asp:comparevalidator id="Comparevalidator2" runat="server" enableclientscript="false" text="***" 
              errormessage="Component amount requested is not numeric" controltovalidate="txtComponentAmount" operator="DataTypeCheck" type="Integer"></asp:comparevalidator>
            <asp:comparevalidator id="Comparevalidator4" runat="server" enableclientscript="false" text="***" 
              errormessage="Component amount must be greater than zero" type="Integer"
              controltovalidate="txtComponentAmount" valuetocompare="0" operator="GreaterThan"></asp:comparevalidator>
            &nbsp;<span class="lowlight">(max <asp:label id="lblCreditRemaining" runat="server" />)</span>
          </td>
        </tr>
        <tr>
          <td class="prompt"><span class="mandatoryfieldasterisk">*</span>Term</td>
          <td>
            <asp:dropdownlist id="cmbTerm" runat="server" />
            <validators:requiredselecteditemvalidator id="Requiredselecteditemvalidator5" runat="server" enableclientscript="false" text="***" 
              errormessage="Term is required" controltovalidate="cmbTerm"></validators:requiredselecteditemvalidator>
          </td>
        </tr>
        <tr>
          <td class="prompt">
          <span class="mandatoryfieldasterisk">
                <asp:label id="lblRepaymentVehicleRequired" runat="server" text="* "></asp:label>
          </span>Repayment vehicle
          </td>
          <td>
            <asp:dropdownlist id="cmbRepaymentVehicle" runat="server" autopostback="true"/>
            <validators:requiredselecteditemvalidator id="RequiredRepaymentVehicle" runat="server" enableclientscript="false" text="***" 
              errormessage="Repayment Vehicle is required" controltovalidate="cmbRepaymentVehicle"></validators:requiredselecteditemvalidator>
          </td>
        </tr>
        <tr>
       
          <td class="prompt">
            <span class="mandatoryfieldasterisk">
              <asp:label id="lblRepaymentAmountRequired" runat="server" text="* "></asp:label>
            </span>
            <asp:label id="lblTitleRepaymentVehicle" runat="server" text="Monthly cost of repayment vehicle"></asp:label>
          </td>
          <td>
            <asp:textbox id="txtRepaymentAmount" runat="server" cssclass="text" />
            <asp:requiredfieldvalidator id="txtRepaymentAmountValidator1" runat="server" enableclientscript="false" 
              errormessage="Repayment amount is required" controltovalidate="txtRepaymentAmount" text="***"></asp:requiredfieldvalidator>
            <asp:comparevalidator id="txtRepaymentAmountValidator2" runat="server" enableclientscript="false" text="***" 
              errormessage="Repayment amount is not numeric" controltovalidate="txtRepaymentAmount" 
              operator="DataTypeCheck" type="Integer"></asp:comparevalidator>
            <asp:comparevalidator id="txtRepaymentAmountValidator3" runat="server" enableclientscript="false" text="***" 
              errormessage="Repayment amount must be greater than zero" type="Integer"
              controltovalidate="txtRepaymentAmount" valuetocompare="0" operator="GreaterThan"></asp:comparevalidator>
          </td>
        </tr>
      </table>

    </fieldset>
    
    <br />
    
    <table>
      <tr>
        <td style="padding: 0 1em 0 0">
          <span class="mandatoryfieldasterisk">*</span>Do you have a product code?
        </td>
        <td>
          <asp:RadioButtonList id="rblProductCodeKnown" runat="Server" autopostback="true" repeatDirection="Horizontal" repeatLayout="Flow">
            <asp:ListItem Value="True">Yes</asp:ListItem>
            <asp:ListItem Value="False">No</asp:ListItem>
          </asp:RadioButtonList>
        </td>
        <td style="padding: 0 0 0 2em">
          <asp:RequiredFieldValidator Runat="server" ID="txtProductCodeKnownRequired" ControlToValidate="rblProductCodeKnown" text="***" errormessage="Specify if you have a product code" EnableClientScript=false ></asp:RequiredFieldValidator> 
          <asp:customvalidator runat="server" id="CustomValidatorProductFound" onservervalidate="ProductSelected" 
             validateemptytext=True  text="***" errormessage="You must either find a known product code, or if its not known, choose no and select one from the list." enableclientscript="false" />
        </td>
      </tr> 
    </table>

    <br />

    <asp:panel id="pnlEnterProductCode" runat="server" visible="false">   

      <fieldset>

        <legend>Component product code</legend>
       
        <table class="formtable">
          <tr>
            <td class="prompt"><span class="mandatoryfieldasterisk">* </span>Product Code</td>
            <td>
              <asp:TextBox ID="txtProductCode" Runat="server" tooltip="Enter a product code" /> 
              &nbsp;
              <asp:button ID="cmdFindProduct" Runat="server" Text="Find" CssClass="button" CausesValidation="False" ></asp:button>
              <asp:RequiredFieldValidator Runat="server" ID="txtProductCodeRequired" ControlToValidate="txtProductCode" text="***" errormessage="Enter a product code" EnableClientScript=false ></asp:RequiredFieldValidator>
            </td>
          </tr>

          <asp:panel id="pnlProductChosen" runat="server" visible="true">
            <tr>
              <td class="prompt">Product details</td>
              <td>
                <asp:label id="lblProductDescription" runat="server" cssclass="text" enabled="false"></asp:label>
              </td>
            </tr>
            <tr>
              <td class="prompt"></td>
              <td> 
                <asp:checkbox id="chkFlexi" runat="server" text="Flexi" enabled="false"></asp:checkbox>
                &nbsp;
                <asp:checkbox id="chkGuarantor" runat="server" text="Guarantor" enabled="false"></asp:checkbox>
              </td>
            </tr>
          </asp:panel>

        </table>

      </fieldset>
        
    </asp:panel>
      
    <asp:panel id="pnlProductList" runat="server" visible="false">
      
      <fieldset>

        <legend>Select product</legend>

        <table class="formtable">
          <tr>
            <td>
              <asp:listbox id="lstProducts" runat="server" autopostback="true" rows="10" style="width: 98%" />
            </td> 
          </tr>
        </table>

      </fieldset>
         
    </asp:panel>

    <asp:panel id="pnlFindProductMessage" runat="server" Visible="False">

      <br />

      <p class="error">
        <asp:label id="lblFindProductMessage" runat="server" />
      </p>

    </asp:panel>
