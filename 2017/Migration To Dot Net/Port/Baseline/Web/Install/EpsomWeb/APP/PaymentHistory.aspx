<%@ Page language="c#" Codebehind="PaymentHistory.aspx.cs" AutoEventWireup="false" Inherits="Epsom.Web.App.PaymentHistory" %>
<%@ Register TagPrefix="mp" namespace="Microsoft.Web.Samples.MasterPages" assembly="MasterPages" %>
<%@ Register Src="~/WebUserControls/DIPNavigationButtons.ascx" TagName="DIPNavigationButtons" TagPrefix="Epsom" %>
<%@ Register Src="~/WebUserControls/DIPMenu.ascx" TagName="DIPMenu" TagPrefix="Epsom" %>
<%@ Register Src="~/WebUserControls/ApplicantType.ascx" TagName="applicantType" TagPrefix="Epsom" %>
<%@ Register TagPrefix="validators" NameSpace="Epsom.Web.Validators" Assembly="Epsom.Web"%>
<%@ Register TagPrefix = "sstchur" Namespace = "sstchur.web.SmartNav" Assembly = "sstchur.web.SmartNav" %>
<%@ Register Src="~/WebUserControls/NullableYesNo.ascx" TagName="NullableYesNo" TagPrefix="Epsom" %>
<%@ Register TagPrefix="CMS" namespace="Epsom.Web.Helpers" assembly="Epsom.Web" %>

<mp:contentcontainer id="Contentcontainer1" runat="server" masterpagefile="../masterpages/masterpage2.ascx">

  <mp:content id="region1" runat="server">

    <epsom:dipmenu id="DipMenu" runat="server"></epsom:dipmenu>

  </mp:content>

  <mp:content id="region2" runat="server">
	<cms:helplink id="helplink" class="helplink" runat="server" helpref="2290" />
    <h1><cms:cmsBoilerplate cmsref="629" runat="server" /></h1>

    <epsom:dipnavigationbuttons id="ctlDIPNavigationButtons1" runat="server"></epsom:dipnavigationbuttons>

    <asp:validationsummary runat="server" displaymode="BulletList" headertext="Input errors:" CssClass="validationsummary" showsummary="true" id="Validationsummary1" />
   
   <cms:cmsBoilerplate cmsref="290" runat="server" /> 
   
    <h2 class="form_heading">Payment History</h2>

    <table class="formtable">
      <tr>
      <td class="prompt"><span class="mandatoryfieldasterisk">*</span> Have you been bankrupt in the last 6 years?</td> 
        <td>
          <epsom:nullableyesno id="nynBankruptInLastSixYears" runat="server" required="true" autopostback="true" Enabled="true"
            errormessage="Specify if you have been bankrupt in the last six years" >
            </epsom:nullableyesno>
        </td>
      </tr> 
      </table>
      
      <asp:panel id="pnlBankruptInLastSixYearsDate" runat="server">
      <asp:repeater id="rptBankruptcies" runat="server">
      <itemtemplate>
        <fieldset>
      <epsom:applicanttype id="ctlBankruptcyApplicantType" runat="server" />
      <table class="formtable">
        <tr>
          <td class="prompt"><span class="mandatoryfieldasterisk">*</span> Discharge date</td>
          <td><asp:textbox id="txtDischargeDate" runat="server"></asp:textbox>
            <validators:datevalidator id="DateValidator1" runat="server" controltovalidate="txtDischargeDate" text="***" errormessage="Invalid Discharge Date"
              enableclientscript="false" /> dd/mm/yyyy
          <asp:requiredfieldvalidator runat="server" id="Requiredfieldvalidator4" controltovalidate="txtDischargeDate" text="***" 
          errormessage="Specify the discharge date" enableclientscript=false ></asp:requiredfieldvalidator> 
          </td>
        </tr>
        <tr>
          <td class="prompt"><span class="mandatoryfieldasterisk">*</span> Amount of debt</td>
          <td>
          <asp:textbox id="txtAmountOfDebt" runat="server"></asp:textbox>
          <asp:requiredfieldvalidator runat="server" controltovalidate="txtAmountOfDebt" text="***" 
              errormessage="Specify amount of debt"
              id="txtAmountOfDebt_validator" enableclientscript="false" />
          <asp:comparevalidator operator=DataTypeCheck type="Currency" id="CurrencyValidator1" runat="server" 
              controltovalidate="txtAmountOfDebt"
              text="***" errormessage="Amount Of Debt is invalid" enableclientscript="false" />
          </td>
        </tr>
        <tr>
          <td class="prompt"><span class="mandatoryfieldasterisk">*</span> Monthly repayment</td>
          <td>
          <asp:textbox id="txtMonthlyRepayment" runat="server"></asp:textbox>
          <asp:requiredfieldvalidator runat="server" controltovalidate="txtMonthlyRepayment" text="***" 
              errormessage="Specify monthly repayment"
              id="Requiredfieldvalidator1" enableclientscript="false" />
          <asp:comparevalidator operator=DataTypeCheck type="Currency" id="Comparevalidator1" runat="server" 
              controltovalidate="txtMonthlyRepayment"
              text="***" errormessage="Monthly repayment is invalid" enableclientscript="false" />
          </td>
        </tr>
        </table> 
      <table class="formtable">
      <tr>
        <td class="prompt"><span class="mandatoryfieldasterisk">*</span> Have you entered into an IVA in the last 6 years?
        </td>
        <td>
          <asp:radiobuttonlist id="rblIVAInLastSixYears" runat="Server" autopostback="true" repeatdirection="Horizontal" repeatlayout="Flow">
            <asp:listitem value="True">Yes</asp:listitem>
            <asp:listitem value="False">No</asp:listitem>
          </asp:radiobuttonlist>
          <asp:requiredfieldvalidator runat="server" id="RequiredfieldvalidatorIVA" controltovalidate="rblIVAInLastSixYears" text="***" errormessage="Specify if you have entered into an IVA in the last six years" enableclientscript=false ></asp:requiredfieldvalidator> 
        </td>
      </tr> 
      </table>
      <asp:panel id="pnlIVAInLastSixYearsDate" runat="server">
      <table class="formtable">
        <tr>
          <td class="prompt"><span class="mandatoryfieldasterisk">*</span> Arrangement / completion date</td>
          <td><asp:textbox id="txtIVAInLastSixYearsDate" runat="server" ></asp:textbox>
          <asp:requiredfieldvalidator runat="server" id="Requiredfieldvalidator2" controltovalidate="txtIVAInLastSixYearsDate" text="***" 
          errormessage="Specify the IVA arrangement/completion date" enableclientscript=false ></asp:requiredfieldvalidator> 
            <validators:datevalidator id="Datevalidator2" runat="server" controltovalidate="txtIVAInLastSixYearsDate" text="***" 
            errormessage="Invalid IVA arrangement/completion date"
              enableclientscript="false" />
              dd/mm/yyyy
          </td>
        </tr>
      </table> 
      </asp:panel>
      </fieldset>
      </itemtemplate>
      </asp:repeater>
     <p class="button_orphan">
      <asp:button id="cmdRemoveBankruptcy" runat="server" causesvalidation="False" text="Remove" cssclass="button" />
      <asp:button id="cmdAddBankruptcy" runat="server" causesvalidation="False" text="Add another bankruptcy" cssclass="button" />
     </p>
     </asp:panel>

      <table class="formtable">
      <tr>
        <td class="prompt"><span class="mandatoryfieldasterisk">*</span> Have you had a County or High Court Judgement for debt registered against you in the last 6 years?</td> 
              <td>
                <epsom:nullableyesno id="nynCCJOrHighCourtJudgementInLastSixYears" runat="server" required="true" autopostback="true" Enabled="true"
                           errormessage="Specify if you have had a CCJ in the last six years" >
                </epsom:nullableyesno> 
        </td>
      </tr>
      </table>
      <table class="formtable">
      <tr>
        <td class="prompt"><span class="mandatoryfieldasterisk">*</span> Have you been refused a mortgage in the last 12 months?</td> 
        <td>
          <epsom:nullableyesno id="nynMortgageRefusedInLast12Months" runat="server" required="true" autopostback="true" Enabled="true"
                    errormessage="Specify if you have been refused a mortgage in the last six years" >
                </epsom:nullableyesno> 
        </td>
      </tr> 
      </table>

      <asp:panel id="pnlMortgageRefusedInLastSixYearsExplanation" runat="server">
      <asp:repeater id="rptMortgageRefused" runat="server">
      <itemtemplate>
        <fieldset>
      <epsom:applicanttype id="ctlMortgageRefusalApplicantType" runat="server" />
      <table class="formtable">
        <tr>
          <td class="prompt" style="vertical-align:top"><span class="mandatoryfieldasterisk">*</span> Details</td>
          <td>
            <asp:textbox id="txtMortgageRefusedInLastSixYearsExplanation" runat="server" textmode="multiline" rows="5" width="300px"></asp:textbox>
            <asp:requiredfieldvalidator runat="server" id="RequiredfieldvalidatorRefused" controltovalidate="txtMortgageRefusedInLastSixYearsExplanation" text="***" 
            errormessage="Specify the details of mortgate refusal" enableclientscript=false ></asp:requiredfieldvalidator> 
            <asp:customvalidator id="Customvalidator1" runat="server"  
              controltovalidate="txtMortgageRefusedInLastSixYearsExplanation"
              errormessage="Maximum number of characters exceeded in refusal explanation - maximum is 250"
              enableclientscript="false" onservervalidate="TextLengthCheck" text="***" >
            </asp:customvalidator> 
          </td>
        </tr>
        </table> 
        </fieldset>
        </itemtemplate>
     </asp:repeater>
     <p class="button_orphan">
      <asp:button id="cmdRemoveMortgageRefusal" runat="server" causesvalidation="False" text="Remove" cssclass="button" />
      <asp:button id="cmdAddMortgageRefusal" runat="server" causesvalidation="False" text="Add another refusal" cssclass="button" />
     </p>
     </asp:panel>

    <epsom:dipnavigationbuttons buttonbar="true" savepage="true" saveandclosepage="true" id="ctlDIPNavigationButtons2" runat="server"></epsom:dipnavigationbuttons>

    <sstchur:smartscroller id="ctlSmartScroller" runat = "server" />
   
  </mp:content>

</mp:contentcontainer>
