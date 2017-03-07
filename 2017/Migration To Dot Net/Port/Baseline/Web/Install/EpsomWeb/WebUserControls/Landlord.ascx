<%@ Control Language="c#" AutoEventWireup="false" Codebehind="Landlord.ascx.cs" Inherits="Epsom.Web.WebUserControls.Landlord" TargetSchema="http://schemas.microsoft.com/intellisense/ie5"%>
<%@ Register Src="Address.ascx" TagName="Address" TagPrefix="Epsom" %>
<%@ Register TagPrefix="validators" NameSpace="Epsom.Web.Validators" Assembly="Epsom.Web"%>

  <table class="formtable">
  <tr>
    <td class="prompt"><span class="mandatoryfieldasterisk">*</span>Landlord's Name</td>
    <td><asp:textbox id="txtLandlordName" runat="server" tooltip="enter the name of your landlord"></asp:textbox>
          <asp:requiredfieldvalidator runat="server" controltovalidate="txtLandlordName" text="***" errormessage="Landlord's name cannot be blank"
        id="txtLandlordName_validator" enableclientscript="false" />
   </td>
  </tr>
  </table> 

  <epsom:address id="ctlAddress" runat="server"  />
  
  <table class=formtable>
    <tr>
    <td class="prompt"> </td>
    <td>Area</td> 
    <td>Number</td>
    <td>ext</td>
    </tr> 
    <tr>
    <td class="prompt"><span class="mandatoryfieldasterisk">*</span>Telephone number</td>
    <td><asp:textbox id="txtTelephoneAreaCode" runat="server" columns ="10"></asp:textbox></td> 
    <td><asp:textbox id="txtTelephoneNumber" runat="server" columns ="10"></asp:textbox>
    </td>
    <td><asp:textbox id="txtTelephoneExtension" runat="server" columns ="10"></asp:textbox>
              <asp:requiredfieldvalidator runat="server" controltovalidate="txtTelephoneAreaCode" text="***" errormessage="Telephone Area Code cannot be blank"
        id="Requiredfieldvalidator1" enableclientscript="false" />
          <asp:requiredfieldvalidator runat="server" controltovalidate="txtTelephoneNumber" text="***" errormessage="Telephone Number cannot be blank"
        id="Requiredfieldvalidator2" enableclientscript="false" />
    </td> 
    </tr> 
    <tr>
    <td class="prompt">Fax number </td>
    <td><asp:textbox id="txtFaxAreaCode" runat="server" columns ="10"></asp:textbox></td>
    <td><asp:textbox id="txtFaxNumber" runat="server" columns ="10"></asp:textbox>
              <asp:requiredfieldvalidator runat="server" controltovalidate="txtFaxAreaCode" text="***" errormessage="Fax Area Code cannot be blank"
        id="Requiredfieldvalidator7" enableclientscript="false" />
          <asp:requiredfieldvalidator runat="server" controltovalidate="txtFaxNumber" text="***" errormessage="Fax Number cannot be blank"
        id="Requiredfieldvalidator8" enableclientscript="false" />
    </td> 
    </tr> 
  </table> 
  <table class="formtable">
  <tr>
    <td class="prompt"><span class="mandatoryfieldasterisk">*</span>Account number</td>
    <td><asp:textbox id="txtAccountNumber" runat="server" ></asp:textbox>
          <asp:requiredfieldvalidator runat="server" controltovalidate="txtAccountNumber" text="***" errormessage="Account number cannot be blank"
        id="Requiredfieldvalidator3" enableclientscript="false" />
   </td>
  </tr>
  <tr>
    <td class="prompt"><span class="mandatoryfieldasterisk">*</span>Monthly payment</td>
    <td><asp:textbox id="txtMonthlyPayment" runat="server" ></asp:textbox>
        <asp:requiredfieldvalidator runat="server" controltovalidate="txtMonthlyPayment" text="***" errormessage="Monthly Payment cannot be blank"
        id="Requiredfieldvalidator4" enableclientscript="false" />
        <asp:comparevalidator id="CompareValidator1" controltovalidate="txtMonthlyPayment" errormessage="Monthly payment is not numeric"
              text="***" enableclientscript="false" runat="server" type="Currency" operator="DataTypeCheck" />

   </td>
  </tr>
  <tr>
    <td class="prompt"><span class="mandatoryfieldasterisk">*</span>Date tenancy started</td>
    <td><asp:textbox id="txtDateTenancyStarted" runat="server" ></asp:textbox>
        <asp:requiredfieldvalidator runat="server" controltovalidate="txtDateTenancyStarted" text="***" errormessage="Date tenancy started cannot be blank"
        id="Requiredfieldvalidator5" enableclientscript="false" />
        <validators:datevalidator id="DateValidator1" runat="server" controltovalidate="txtDateTenancyStarted" text="***" errormessage="Date tenancy started must be a valid date"
          enableclientscript="false" />
   </td>
  </tr>
  <tr>
    <td class="prompt"><span class="mandatoryfieldasterisk">*</span>Number of months currently in arrears?</td>
    <td><asp:textbox id="txtMonthsInArrears" runat="server" ></asp:textbox>
        <asp:requiredfieldvalidator runat="server" controltovalidate="txtMonthsInArrears" text="***" errormessage="Months In Arrears cannot be blank"
        id="Requiredfieldvalidator6" enableclientscript="false" />
        <asp:comparevalidator id="Comparevalidator2" controltovalidate="txtMonthsInArrears" errormessage="Months In Arrears is not numeric"
              text="***" enableclientscript="false" runat="server" type="Integer" operator="DataTypeCheck" />
   </td>
  </tr>
    <tr>
    <td class="prompt"><span class="mandatoryfieldasterisk">*</span> Have you been in receipt of Housing Benefit in the last 12 months?
          <td>
            <asp:radiobuttonlist id="rblReceivedHousingBenefit" runat="Server" autopostback="true" repeatdirection="Horizontal" repeatlayout="Flow">
              <asp:listitem value="True">Yes</asp:listitem>
              <asp:listitem value="False">No</asp:listitem>
            </asp:radiobuttonlist>
            <asp:requiredfieldvalidator runat="server" id="rblReceivedHousingBenefitValidator" controltovalidate="rblReceivedHousingBenefit" text="***" errormessage="Specify if you have received housing benefit" enableclientscript=false ></asp:requiredfieldvalidator> 
          </td>
        </tr> 
    </td>
  </tr>
  <asp:panel id="pnlHousingBenefitExplanation" runat="server">
  <tr>
    <td class="prompt">Please give details</td>
    <td><asp:textbox id="txtHousingBenefitExplanation" runat="server" rows="5" textmode="multiline"></asp:textbox>
   </td>
  </tr>
</asp:panel>
  </table> 
