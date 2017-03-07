<%@ Control Language="c#" AutoEventWireup="false" Codebehind="EmploymentSelfEmployedDetails.ascx.cs" Inherits="Epsom.Web.WebUserControls.EmploymentSelfEmployedDetails" TargetSchema="http://schemas.microsoft.com/intellisense/ie5"%>
<%@ Register Src="~/WebUserControls/Address.ascx" TagName="Address" TagPrefix="Epsom" %>
<%@ Register Src="~/WebUserControls/TelephoneNumbers.ascx" TagName="TelephoneNumbers" TagPrefix="Epsom" %>
  <H2 class="form_heading">Self Employed/Controlling Directors</H2>
  <TABLE class="formtable">
    <TR>
      <TD class="prompt"><asp:label id="lblBusinessName" runat="server" text="Name of business"></asp:label></TD>
      <TD><asp:textbox id="txtBusinessName" runat="server"  cssclass="text"   tooltip="enter the name of your business"></asp:textbox></TD>
   <!--   <TD><asp:requiredfieldvalidator id="txtBusinessName_validator" runat="server" text="***" controltovalidate="txtBusinessName"
        errormessage="Business name cannot be blank" EnableClientScript="false"></asp:requiredfieldvalidator></TD>-->
   </TR>
      
   <TR> 
     <TD></TD>
 
  </TR>
 </TABLE>

 <epsom:Address id="ctlBusinessAddress" runat="server" isresidential="True" unitedkingdomonly="True" shownatureofoccupancy="False" showresidencydates="False" ></epsom:Address>
 <epsom:telephonenumbers id="ctlTelephoneNumbers" runat="server" />
  
   
   <TABLE class="formtable"> 
    <TR>
      <TD class="prompt">
        <asp:label id="lblBusinessEmail" runat="server" text="Business email address"></asp:label></TD>
      <TD>
        <asp:textbox  cssclass="text" id="txtBusinessEmail" runat="server" text="" tooltip="enter your business email address"></asp:textbox></TD>
        
    </TR>
    <TR>
      <TD class="prompt">
        <asp:label id="lblJobTitleNatureOfBusiness" runat="server" text="JobTitle/Nature of business "></asp:label><span class="mandatoryfieldasterisk">*</span></TD>
      <TD>
        <asp:textbox  cssclass="text" id="txtJobTitleNatureOfBusiness" runat="server" text="" tooltip="enter your job title or the nature of your business"></asp:textbox></TD>
        
    </TR>
    <TR>
      <TD class="prompt">
        <asp:label id="lblSelfEmpNINumber" runat="server" text="National Insurance Number"></asp:label></TD>
      <TD>
        <asp:textbox  cssclass="text" id="txtSelfEmpNINumber" runat="server" text="" tooltip="enter your national insurance number"></asp:textbox></TD>
        
    </TR>
    <TR>
      <TD class="prompt">
        <asp:label id="lblCompanyRegNo" runat="server" text="Company Registration Number (if applicable)"></asp:label></TD>
      <TD>
        <asp:textbox  cssclass="text" id="txtCompanyRegNo" runat="server" text="" tooltip="enter your company registration number"></asp:textbox></TD>
        
    </TR>
    <TR>
      <TD class="prompt">
        <asp:label id="lblVatNo" runat="server" text="VAT number (if applicable)"></asp:label></TD>
      <TD>
        <asp:textbox  cssclass="text" id="txtVatNo" runat="server" text="" tooltip="enter your VAT number"></asp:textbox></TD>
        
    </TR>
    <TR>
      <TD class="prompt">
        <asp:label id="lblTradingLength" runat="server" text="How long trading?"></asp:label></TD>
      <TD>
        <asp:textbox  cssclass="text" id="txtTradingLength" runat="server" tooltip="enter the amount of time you have been trading for"></asp:textbox></TD>
        
    </TR>
    <TR>
      <TD class="prompt">
        <asp:label id="lblBusinessAcquireDate" runat="server" text="What date did you acquire an interest in the business?"></asp:label></TD>
      <TD>
        <asp:textbox  cssclass="text" id="txtBusinessAcquireDate" runat="server" text="" tooltip="enter the date you acquired an interest in the business (dd/mm/yyyy)"
          columns="10"></asp:textbox>
          
    </TR>
    <TR>
      <TD class="prompt">
        <asp:label id="lblEstablishedLength" runat="server" text="How long has the business been established?"></asp:label></TD>
      <TD>
        <asp:textbox  cssclass="text" id="txtEstablishedLength" runat="server" tooltip="enter the amount of time that the business has been established"></asp:textbox></TD>
       
    </TR>
    <TR>
      <TD class="prompt">
        <asp:label id="lblPercentShare" runat="server" text="What is your % shareholding?"></asp:label></TD>
      <TD>
        <asp:textbox  cssclass="text" id="txtPercentShare" runat="server" tooltip="enter your percentage shareholding"></asp:textbox></TD>
       
    </TR>
    <TR>
      <TD class="prompt">
        <asp:label id="lblBusinessType" runat="server" text="Is the business"></asp:label></TD>
      <TD>
        <asp:dropdownlist id="ddlBusinessType" runat="server">
          <asp:ListItem value="0" runat="server" ID="Listitem1" NAME="Listitem1">Limited Company</asp:ListItem>
          <asp:ListItem value="1" runat="server" ID="Listitem2" NAME="Listitem2">Partnership</asp:ListItem>
          <asp:ListItem value="2" runat="server" ID="Listitem3" NAME="Listitem3">Sole Trader</asp:ListItem>
        </asp:dropdownlist></TD>
        
    </TR>
    <TR>
      <TD class="prompt">
        <asp:label id="lblSelfAssess" runat="server" text="Do you self assess to the Inland Revenue?"></asp:label></TD>
      <TD>
        <asp:RadioButtonList id="rblSelfAssess" runat="Server" tooltip="state whether you self assess to the Inland Revenue or not"
          RepeatLayout="flow" RepeatDirection="horizontal" autopostback="true">
          <asp:ListItem value="1" runat="server" ID="Listitem4" NAME="Listitem4">Yes</asp:ListItem>
          <asp:ListItem value="2" runat="server" ID="Listitem5" NAME="Listitem5">No</asp:ListItem>
        </asp:RadioButtonList></TD>
    </TR>
    <tr><td class="prompt"><asp:label id="lblAccountantYrs" runat="server" text="How long has this Accountant acted for you? "></asp:label></td>
    <td><asp:textbox  cssclass="text" id="txtAccountantYrs" runat="server" text=""></asp:textbox></td></tr>
    <tr><td class="prompt"><asp:label id="lblAccountantName" runat="server" text="Name of accountant"></asp:label></td>
    <td><asp:textbox  cssclass="text" id="txtAccountantName" runat="server" text=""></asp:textbox></td></tr>
    <tr><td class="prompt"><asp:label id="lblAccountantFirm" runat="server" text="Name of company / firm"></asp:label></td>
    <td><asp:textbox  cssclass="text" id="txtAccountantFirm" runat="server" text=""></asp:textbox></td></tr>
  
   
 
  
    <TR>
      <TD class="prompt" TD>
        <asp:label id="Label1" runat="server" text="Building number"></asp:label></TD>
      <TD>
        <asp:textbox  cssclass="text" id="Textbox1" runat="server" text=""></asp:textbox></TD>
        
    <TR>
      <TD class="prompt" TD>
        <asp:label id="Label2" runat="server" text="Building number/name"></asp:label></TD>
      <TD class="prompt" TD>
        <asp:textbox  cssclass="text" id="Label12" runat="server" text="" width="50px"></asp:textbox>
        <asp:textbox  cssclass="text" id="Textbox8" runat="server" text=""></asp:textbox></TD>
        
    <TR>
      <TD class="prompt" TD>
        <asp:label id="Label3" runat="server" text="Postcode*"></asp:label></TD>
      <TD>
        <asp:textbox  cssclass="text" id="txtPostCode" runat="server"></asp:textbox></td>
       <TD> <asp:button id="Validate" runat="server" Text="Validate Address"></asp:button></TD>
        
    </TR>

   </TABLE>
     <epsom:address  cssclass="text" id="ctlAccountantAddress" runat="server" Width="200px" Height="100px" isresidential="False" unitedkingdomonly="True" shownatureofoccupancy="False" showresidencydates="False" ></epsom:address>
     <epsom:telephonenumbers id="ctlAccountantTelephoneNumbers" runat="server" />

      <TABLE class="formtable" >
    <TR>
      <TD class="prompt">
        <asp:label id="lblAccountantPhoneNo" runat="server" text="Accountant's telephone number"></asp:label></TD>
      <TD>
        <asp:textbox  Columns="5" cssclass="text" id="txtAccountantAreaCode" runat="server" text="" tooltip="enter your accountant's phone number"></asp:textbox>
         <asp:textbox Columns="14"  cssclass="text" id="txtAccountantTelephone" runat="server" text="" tooltip=""></asp:textbox>
          <asp:textbox Columns="5"  cssclass="text" id="txtAccountantExt" runat="server" text="" tooltip=""></asp:textbox>
        </TD>
        
    </TR>
    <TR>
      <TD class="prompt">
        <asp:label id="lblAccountantFaxNo" runat="server" text="Accountant's fax number"></asp:label></TD>
      <TD>
    <asp:textbox Columns="5"  cssclass="text" id="txtAccountantFaxCode" runat="server" text="" tooltip="enter your accountant's fax number"></asp:textbox></TD>
    <asp:textbox  Columns="14" cssclass="text" id="txtAccountantFax" runat="server" text="" tooltip="enter your accountant's fax number"></asp:textbox></TD>
 
    </TR>
    <TR>
      <TD class="prompt">
        <asp:label id="lblAccountantEmail" runat="server" text="Accountant's email address"></asp:label></TD>
      <TD>
        <asp:textbox  cssclass="text" id="txtAccountantEmail" runat="server" text="" tooltip="enter your accountant's email address"></asp:textbox></TD>
        
    </TR>
    <TR>
      <TD class="prompt">
        <asp:label id="lblAccountantQuals" runat="server" text="Accountant's qualifications"></asp:label></TD>
      <TD>
        <asp:textbox  cssclass="text" id="txtAccountantQuals" runat="server" text="" tooltip="enter your accountant's qualifications"></asp:textbox></TD>
        
    </TR>
  </TABLE>



