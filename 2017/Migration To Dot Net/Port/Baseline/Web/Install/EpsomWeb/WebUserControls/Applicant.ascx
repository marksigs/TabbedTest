<%@ Control Language="c#" AutoEventWireup="false" Codebehind="Applicant.ascx.cs" Inherits="Epsom.Web.WebUserControls.Applicant" TargetSchema="http://schemas.microsoft.com/intellisense/ie5"%>
<%@ Register Src="Address.ascx" TagName="Address" TagPrefix="Epsom" %>
<%@ Register Src="Commitment.ascx" TagName="Commitment" TagPrefix="Epsom" %>
<%@ Register Src="Income.ascx" TagName="Income" TagPrefix="Epsom" %>
<%@ Register Src="ApplicantOtherName.ascx" TagName="ApplicantOtherName" TagPrefix="Epsom" %>
<%@ Register Src="TelephoneNumbers.ascx" TagName="TelephoneNumbers" TagPrefix="Epsom" %>
<%@ Register TagPrefix="validators" NameSpace="Epsom.Web.Validators" Assembly="Epsom.Web"%>
<%@ Register Src="~/WebUserControls/NullableYesNo.ascx" TagName="NullableYesNo" TagPrefix="Epsom" %>
<%@ Register TagPrefix="CMS" namespace="Epsom.Web.Helpers" assembly="Epsom.Web" %>

<h2 class="form_heading">Personal details</h2>

<table class="formtable">
  <tr>
    <td class="prompt"><span class="mandatoryfieldasterisk">*</span>Title</td>
    <td><asp:dropdownlist id="cmbTitle" runat="server" />
        <validators:RequiredSelectedItemValidator Runat="server" EnableClientScript="false"
          Text="***" ErrorMessage="Title is required" ControlToValidate="cmbTitle" />
    </td>
  </tr>
  <tr>
    <td class="prompt"><span class="mandatoryfieldasterisk">*</span>First forename</td>
    <td>
      <asp:textbox id="txtApplicantFirstName" runat="server" columns="40" CssClass="text" />
      <asp:requiredfieldvalidator runat="server" controltovalidate="txtApplicantFirstName" text="***" errormessage="First name cannot be blank"
        id="txtApplicantFirstName_validator" EnableClientScript="false" />
    </td>
  </tr>
  <tr>
    <td class="prompt">Second forename</td>
    <td><asp:textbox id="txtApplicantMiddleNames" runat="server" columns="40" CssClass="text" /></td>
  </tr>
  <tr>
    <td class="prompt"><span class="mandatoryfieldasterisk">*</span>Surname</td>
    <td>
      <asp:textbox id="txtApplicantLastName" runat="server" columns="40" CssClass="text" />
      <asp:requiredfieldvalidator runat="server" controltovalidate="txtApplicantLastName" text="***" errormessage="Last name cannot be blank"
        id="txtApplicantLastName_validator" EnableClientScript="false" />
    </td>
  </tr>
  <tr>
    <td class="prompt"><span class="mandatoryfieldasterisk">*</span>Date of birth</td>
    <td>
      <asp:textbox id="txtDateOfBirth" runat="server" columns="10" CssClass="text" />
      <validators:DateOfBirthValidator ID="DateOfBirthValidator1" Runat="server" ControlToValidate="txtDateOfBirth" text="***" ErrorMessage="Invalid date of birth"
        EnableClientScript="false" />
        dd/mm/yyyy
    </td>
  </tr>
  <tr>
    <td class="prompt"><span class="mandatoryfieldasterisk">*</span>Gender</td>
    <td><asp:dropdownlist id="cmbGender" runat="server" />
        <validators:RequiredSelectedItemValidator Runat="server" EnableClientScript="false"
          Text="***" ErrorMessage="Gender is required" ControlToValidate="cmbGender"/>
    </td>
  </tr>
  <tr>
    <td class="prompt"><span class="mandatoryfieldasterisk">*</span>Marital status</td>
    <td><asp:dropdownlist id="cmbMaritalStatus" runat="server" />
        <validators:RequiredSelectedItemValidator Runat="server" EnableClientScript="false"
          Text="***" ErrorMessage="Marital status is required" ControlToValidate="cmbMaritalStatus"/>
    </td>
  </tr>
  <asp:panel runat ="server" id="pnlNationality1">
  <tr>
    <td class="prompt"><span class="mandatoryfieldasterisk">*</span>Nationality</td>
    <td><asp:dropdownlist id="cmbNationality1" runat="server" />
        <validators:RequiredSelectedItemValidator Runat="server" EnableClientScript="false"
          Text="***" ErrorMessage="Nationality is required" ControlToValidate="cmbNationality1"/>
    </td>
  </tr>
  </asp:panel>
</table>

<asp:Panel ID="pnlAllowOtherNames" Runat="server">

  <h2 class="form_heading">Other Names</h2>
  <table class=formtable>
    <tr>
      <td style="padding: 0 0.5em 0 0.5em">
        Has your name changed in the last three years? 
        <asp:radiobuttonlist id="rblOtherNames" runat="Server" autopostback="true" repeatdirection="Horizontal" repeatlayout="Flow">
          <asp:listitem value="True">Yes</asp:listitem>
          <asp:listitem value="False">No</asp:listitem>
        </asp:radiobuttonlist>
      </td>
    </tr> 
  </table>

  <asp:panel runat="server" id="pnlOtherNames">

    <table class=formtable>
      <tr>
        <td style="padding: 0 0.5em 0 0.5em">
        Please list all names you have been known by during the last three years:
        </td>
      </tr>
    </table> 

    <asp:repeater id="rptOtherNames" runat="server">
      <itemtemplate>
        <fieldset>
          <legend><asp:label id="lblOtherNameTitle" runat="server" /></legend>
          <epsom:applicantOtherName id="ctlApplicantOtherName" runat="server" />
        </fieldset> 
      </itemtemplate>
    </asp:repeater>

    <br />

    <p class="button_orphan">
      <asp:button id="cmdRemoveOtherName" runat="server" causesvalidation="False" text="Remove this name" cssclass="button" />
      <asp:button id="cmdAddOtherName" runat="server" causesvalidation="False" text="Add another name" cssclass="button" />
    </p>

  </asp:panel>
  
</asp:Panel>

<asp:panel runat ="server" id="pnlNationality2" visible="false">
  <h2 class="form_heading">Nationality and Residency</h2>

  <table class="formtable">
    <tr>
      <td class="prompt"><span class="mandatoryfieldasterisk">*</span>Nationality</td>
      <!-- cmbNationality2 is a duplicate of cmbNationality1. 
      However, when it is visible, it is readonly, so no need to validate -->
      <td>
        <asp:dropdownlist id="cmbNationality2" runat="server" enabled="false" />
      </td>
    </tr>
    <tr>
      <td class="prompt"><span class="mandatoryfieldasterisk">*</span>Resident in the UK?
      <td>
        <epsom:nullableyesno  id="nynUkResident" runat="server" autopostback="true" required="true" errormessage="Specify if you are a UK resident" enableclientscript=false >
        </epsom:nullableyesno> 
      </td>
    </tr> 
  <asp:panel runat ="server" id="pnlNotResidentUk" visible="false">
    <tr>
      <td class="prompt">If no, where do you reside?</td>
      <td><asp:textbox id="txtUKNonResidentCountry" runat="server" columns="40" cssclass="text" /></td>
    </tr>
       <tr>
      <td class="prompt">Please give details why?</td>
      <td><asp:textbox id="txtUKNonResidentExplanation" runat="server" columns="40" cssclass="text" /></td>
    </tr>
   </asp:Panel>
    <tr>
      <td class="prompt"><span class="mandatoryfieldasterisk">*</span>Do you have the right to work and reside permanently in the UK?</td>
      <td>
        <epsom:nullableyesno  id="nynUKWorkAndResidePermitted" runat="server" autopostback="true" required="true"
         errormessage="Specify if you have the right to reside and work in the UK" >
         </epsom:nullableyesno> 
      </td>
    </tr> 
  <asp:panel runat ="server" id="pnlNotWorkUk" visible="false">
    <tr>
      <td class="prompt">Please give details</td>
      <td><asp:textbox id="txtUKWorkAndResidePermittedExplanation" runat="server" columns="40" cssclass="text" /></td>
    </tr>
   </asp:Panel>
    <tr>
      <td class="prompt"><span class="mandatoryfieldasterisk">*</span>Do you have Diplomatic Immunity?
      <td>
        <epsom:nullableyesno  id="nynDiplomaticImmunity" runat="server" autopostback="true" required="true"
         errormessage="Specify if you have diplomatic immunity"  >
         </epsom:nullableyesno> 
      </td>
    </tr> 
  <asp:panel runat ="server" id="pnlNotDiplomaticImmunity" visible="false">
    <tr>
      <td class="prompt">Please give details</td>
      <td><asp:textbox id="txtDiplomaticImmunityExplanation" runat="server" columns="40" cssclass="text" /></td>
    </tr>
   </asp:Panel>
  </table>
</asp:panel>

<h2 class="form_heading">Address details</h2>

<asp:repeater id="rptAddresses" runat="server">
  <itemtemplate>
    <fieldset>
      <legend><asp:Label ID="lblAddressTitle" Runat="server" /></legend>
      <epsom:Address id="ctlAddress" runat="server"  isresidential="true" />
    </fieldset> 
  </itemtemplate>
</asp:repeater>

<asp:Panel ID="pnlAllowOtherAddresses" Runat="server">   
 
  <br />

  <p class="button_orphan">
    <asp:button id="cmdRemoveAddress" runat="server" causesvalidation="False" text="Remove address" cssclass="button" />
    <asp:button id="cmdAddAddress" runat="server" CausesValidation="False" Text="Add another address" CssClass="button" />
    <asp:CustomValidator Runat="server" ID="rptEnoughAddresses" OnServerValidate="EnoughAddresses" 
      ValidateEmptyText=True  text="***" errormessage="Not Enough Addresses" EnableClientScript="false" />
    <asp:Label ID="lblAddAddressError" Runat="server" cssclass="mandatoryfieldasterisk"/>
  </p>

</asp:Panel>

<asp:panel id="pnlTelephoneDetails" runat="server" visible ="false"> 
<h2 class="form_heading">Telephone details</h2>
    <epsom:TelephoneNumbers id="ctlTelephoneNumbers" runat="server" ShowPreference="true"/>
</asp:panel> 

<asp:panel id="pnlCorrespondenceAddressHeader" runat="server" visible="false">
  <h2 class="form_heading">Correspondence Address</h2>
</asp:panel>

<asp:panel id="pnlCorrespondenceAddressAdd" runat="server" visible="false">
  <br />
  <p class="button_orphan">
    <asp:button id="cmdAddCorrespondenceAddress" runat="server" causesvalidation="False" text="Add Correspondence Address" cssclass="button" />
  </p>
</asp:panel>

<asp:panel id="pnlCorrespondenceAddress" runat="server" visible="false">   
  <epsom:address id="ctlCorrespondenceAddress" runat="server"  isresidential="false" />
  <br />
  <p class="button_orphan">
    <asp:button id="cmdRemoveCorrespondenceAddress" runat="server" causesvalidation="False" text="Remove" cssclass="button" />
  </p>
</asp:panel>

<asp:panel id="pnlHousingBenefit" runat="server" visible="false">   
<h2 class="form_heading">Housing Benefit</h2>
  <table class="formtable">
    <tr>
      <td class="prompt"><span class="mandatoryfieldasterisk">* </span>Have you been in receipt of housing benefit in the last 12 months?
      <td>
        <epsom:nullableyesno  id="nynHousingBenefit" runat="server" autopostback="true" required="true"
         errormessage="Specify if you been in receipt of housing benefit in the last 12 months" enableclientscript=false >
        </epsom:nullableyesno> 
      </td>
    </tr> 
  <asp:panel runat ="server" id="pnlNoHousingBenefit" visible="false">
    <tr>
      <td class="prompt">Please give details</td>
      <td><asp:textbox id="txtNoHousingBenefit" runat="server" columns="40" cssclass="text" /></td>
    </tr>
   </asp:Panel>
  </table>
</asp:panel>


<asp:Panel ID="pnlAllowIncomeDetails" runat="server"> 

  <h2 class="form_heading">Income details</h2>

  <epsom:income id="ctlIncome" runat="server" />

</asp:Panel> 


</asp:Panel>

<br />

<p class="button_orphan">
  <asp:button id="cmdRemoveApplicant" runat="server" causesvalidation="False" text="Remove Applicant" cssclass="button" />
</p>
