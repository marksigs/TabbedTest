<%@ Control Language="c#" AutoEventWireup="false" Codebehind="LandlordDetail.ascx.cs" Inherits="Epsom.Web.WebUserControls.LandlordDetail" TargetSchema="http://schemas.microsoft.com/intellisense/ie5"%>
<%@ Register Src="TelephoneNumbers.ascx" TagName="TelephoneNumbers" TagPrefix="Epsom" %>
<%@ Register Src="Address.ascx" TagName="Address" TagPrefix="Epsom" %>
<%@ Register TagPrefix="validators" NameSpace="Epsom.Web.Validators" Assembly="Epsom.Web"%>

<table class="formtable">
  <tr>
    <td  class="prompt">Please enter the landlord details for this address</td>
    <td>Address details<br><asp:Label ID="lblRentalAddress" Runat="server" /></td>
  </tr>
  <tr>
    <td  class="prompt"><span class="mandatoryfieldasterisk">*</span> Landlord name</td>
    <td>
    <asp:TextBox ID="txtLandlordName" Runat="server" />
    <asp:requiredfieldvalidator controltovalidate="txtLandlordName" runat="server" text="***" 
      errormessage="Please enter landlord name" id="Requiredfieldvalidator1" enableclientscript="false"/>
    </td>
  </tr>
</table>

  <epsom:address id="ctlLandlordAddress" runat="server"  isresidential="false" />

  <epsom:TelephoneNumbers id="ctlLandlordTelephoneNumbers" runat="server" FirstItemMandatory="true" usageComboGroup="ContactTelephoneUsage"/>

<table class="formtable">
  <tr>
    <td  class="prompt">Email address</td>
    <td><asp:TextBox ID="txtLandlordEmail" Runat="server" /></td>
  </tr>
  <tr>
    <td  class="prompt"><span class="mandatoryfieldasterisk">*</span> Applicant account number</td>
    <td>
      <asp:TextBox ID="txtAccountNumber" Runat="server" />
      <asp:requiredfieldvalidator controltovalidate="txtAccountNumber" runat="server" text="***" 
      errormessage="Please enter account number" id="Requiredfieldvalidator2" enableclientscript="false"/>
    </td>
  </tr>
  <tr>
    <td  class="prompt"><span class="mandatoryfieldasterisk">*</span> Monthly payment</td>
    <td>
      <asp:TextBox ID="txtMonthlyPayment" Runat="server" />
      <asp:requiredfieldvalidator controltovalidate="txtMonthlyPayment" runat="server" text="***" 
        errormessage="Please enter monthly payment" id="Requiredfieldvalidator3" enableclientscript="false"/>
      <asp:rangevalidator id="Rangevalidator1" runat="server" enableclientscript="false" 
                  errormessage="Monthly payment must be numeric" controltovalidate="txtMonthlyPayment" 
                  minimumvalue="0" maximumvalue="9999999999" type="Currency" text="***"></asp:rangevalidator>

    </td>
  </tr>
  <tr>
    <td  class="prompt"><span class="mandatoryfieldasterisk">*</span> Date tenancy started</td>
    <td>
    <asp:TextBox ID="txtDateTenancyStarted" Runat="server" />
    <asp:requiredfieldvalidator controltovalidate="txtDateTenancyStarted" runat="server" text="***" 
      errormessage="Please enter date tenancy started" id="Requiredfieldvalidator4" enableclientscript="false"/>
    <validators:datevalidator id="DateValidator1" runat="server" controltovalidate="txtDateTenancyStarted" 
            text="***" errormessage="Invalid date tenancy started"
          enableclientscript="false" /> dd/mm/yyyy
    </td>
  </tr>
</table>

<asp:Panel ID="pnlCurrentAndFirstInstanceAddress" Runat="server">
  <table class="formtable">
    <tr>
      <td  class="prompt"><span class="mandatoryfieldasterisk">*</span> Is this tenancy in arrears?</td>
      <td>
        <asp:RadioButton ID="rdoArrearsYes" GroupName="grpArrears" Runat="server" AutoPostBack="True" Text="Yes" />
        <asp:RadioButton ID="rdoArrearsNo"  GroupName="grpArrears" Runat="server" AutoPostBack="True" Text="No" />
        <asp:customvalidator id="Customvalidator1" runat="server" enableclientscript="false"
        onservervalidate="ValidatorArrearsDeclarationRequired"
        errormessage="You must say if there are any arrears" text="***" />
       </td>
    </tr>
  </table>
  
  <asp:Panel ID="pnlArrearsDetails" Runat="server">
    <table class="formtable">
      <tr>
        <td  class="prompt"><span class="mandatoryfieldasterisk">*</span> Who is in arears?</td>
        <td>
          <asp:CheckBox ID="chkApplicant1Arrears" Runat="server" Text="Applicant 1"/>
          <asp:CheckBox ID="chkApplicant2Arrears" Runat="server" Text="Applicant 2"/>
          <asp:CheckBox ID="chkGuarantorArrears"  Runat="server" Text="Guarantor"/>
          <asp:customvalidator id="Customvalidator2" runat="server" enableclientscript="false"
        onservervalidate="ValidatorArrearsApplicantRequired"
        errormessage="You must say who is in arrears" text="***" />

        </td>
      </tr>
      <tr>
        <td  class="prompt"><span class="mandatoryfieldasterisk">*</span> Number of months currently in arrears</td>
        <td>
          <asp:dropdownlist id="cmbMonthsInArrears" runat="server" height="13">
            <asp:ListItem runat="server" Value=""  id="nullMonthInArrears">Please Select...</asp:ListItem>
            <asp:listitem runat="server" value="1" id="Listitem23">1</asp:listitem>
            <asp:listitem runat="server" value="2" id="Listitem24">2</asp:listitem>
            <asp:listitem runat="server" value="3" id="Listitem25">3</asp:listitem>
            <asp:listitem runat="server" value="4" id="Listitem26">4</asp:listitem>
            <asp:listitem runat="server" value="5" id="Listitem27">5</asp:listitem>
            <asp:listitem runat="server" value="6" id="Listitem28">6</asp:listitem>
            <asp:listitem runat="server" value="7" id="Listitem29">7</asp:listitem>
            <asp:listitem runat="server" value="8" id="Listitem30">8</asp:listitem>
            <asp:listitem runat="server" value="9" id="Listitem31">9</asp:listitem>
            <asp:listitem runat="server" value="10" id="Listitem32">10</asp:listitem>
            <asp:listitem runat="server" value="11" id="Listitem33">11</asp:listitem>
            <asp:listitem runat="server" value="12" id="Listitem34">12</asp:listitem>
          </asp:dropdownlist>
      <validators:requiredselecteditemvalidator runat="server" enableclientscript="false"
          text="***" errormessage="Number of months in arrears is required" controltovalidate="cmbMonthsInArrears" 
          id="Requiredselecteditemvalidator2"/> 
        </td>
      </tr>
    </table>
  </asp:Panel>

</asp:Panel>