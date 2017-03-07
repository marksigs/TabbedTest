<%@ Control Language="c#" AutoEventWireup="false" Codebehind="Address.ascx.cs" Inherits="Epsom.Web.WebUserControls.Address" TargetSchema="http://schemas.microsoft.com/intellisense/ie5"%>
<%@ Register TagPrefix="validators" NameSpace="Epsom.Web.Validators" Assembly="Epsom.Web"%>
<asp:panel runat="server" id="pnlBFPO">
 <table class="formtable">
   <tr>
    <td class="prompt"><asp:label id="lblBritishForcesPostOffice" runat="server" text="BFPO" /></td>
    <td>
    <asp:radiobutton 
      groupname="grpBFPOYN" 
      id="rdoBFPOYes"
      text="Yes" 
      runat="server" autopostback="True">
    </asp:radiobutton>
    <asp:radiobutton 
      groupname="grpBFPOYN" 
      id="rdoBFPONo"
      text="No" 
      runat="server" autopostback="True" Checked="True">
    </asp:radiobutton>
    </td>
   </tr>
 </table>
</asp:Panel> 

<asp:panel runat="server" id="pnlCountryType">
  <table class="formtable">
    <tr>
      <td class="prompt"><asp:label id="lblCountryType" runat="server" text="Country"  /></td>
      <td>
        <asp:dropdownlist id="cmbCountryType" runat="server" autopostback="True" />
       </td>
    </tr>
  </table> 
</asp:panel> 

<asp:label id="lblSystemException" runat="server" text="" ForeColor="red" />
  
<asp:panel runat="server" id="pnlPafSearch">
  <table class="formtable" >
   <tr>
      <td class="prompt"><span class="mandatoryfieldgroupasterisk">*</span> <asp:label id="lblFlatNameOrNumber" runat="server" text="Flat" /></td>
      <td>
        <asp:textbox id="txtSearchFlatNameOrNumber" runat="server" columns="23" cssclass="text" />
        </td>
      <td rowspan="3" width="20%">
        <asp:label id="lblPrevalidationExplanation" runat="server">
        Enter the flat, building number or name and the postcode to validate</asp:label>
     </td>
   </tr> 
   <tr>
      <td class="prompt"><span class="mandatoryfieldgroupasterisk">*</span> <asp:label id="lblHouseNameOrNumber" runat="server" text="Building Number/Name" /></td>
      <td>
        <asp:textbox id="txtSearchHouseNumber" runat="server" columns="4" cssclass="text" />
        <asp:textbox id="txtSearchHouseName" runat="server" columns="39" CssClass="text" />
        <asp:button id="cmdCopyAddress" runat="server" text="Copy Address" CssClass="button" CausesValidation="False" />
      </td>
      <td></td>
    </tr> 
    <tr>
      <td class="prompt"><span class="mandatoryfieldasterisk">*</span> <asp:label id="lblSearchPostcode" runat="server" text="Postcode" /></td>
      <td>
        <asp:textbox id="txtSearchPostcode" runat="server" columns="10"  CssClass="text" />
        <asp:button id="cmdValidatePostcode" runat="server" text="Validate" CssClass="button" CausesValidation="False" />
        <asp:button id="cmdResetPostcode" runat="server" text="Reset" CssClass="button" CausesValidation="False" />
        <asp:customvalidator runat="server" id="Customvalidator1" onservervalidate="MandatoryAddressValidator" 
            validateemptytext=True  text="***" errormessage="You must validate the postcode and supply an address" enableclientscript="false" />
        <asp:label class="mandatoryfieldasterisk" id="lblValidationFails" runat="server"></asp:label>
      </td>
      <td></td>
    </tr>
  </table>
</asp:panel> 


    
<asp:panel ID="pnlPafResults" runat="server">
  <table class="formtable">
    <tr>
      <td class="prompt"><span class="mandatoryfieldasterisk">*</span> <asp:label id="lblValidAddresses" runat="server" text="Valid addresses - please click to choose" /></td>
      <td>
        <asp:listbox id="lstValidAddresses" runat="server" rows="5" style="width: 95%" cssclass="validaddresseslist" AutoPostBack="True"/>
      </td>
    </tr>
  </table>
</asp:Panel>
 
<asp:panel id="pnlNoValidAddresses" runat="server">
  <table class="formtable">
    <tr>
      <td class="prompt"></td>
      <td>
          <asp:label id="lblNoValidAddresses" runat="server" text="No valid address found. Either reset and try another building and postcode, or fill in an address below" />
      </td>
    </tr>
  </table>
</asp:panel> 

<asp:panel runat="server" id="pnlAddressLines">
  <table class="formtable">
    <asp:panel runat="server" visible="False" id="pnlAddressFlatNumber" >
    <tr>
      <td class="prompt"><span class="mandatoryfieldgroupasterisk">*</span> <asp:label  id="lblAddressFlatNumber" runat="server" text="Flat" /></td>
      <td>
       <asp:textbox id="txtAddressFlatNumber" runat="server" columns="23" CssClass="text" maxlength="40" />
      </td>
      <td rowspan="2" width="20%">
       <asp:label id="lblMustEnterFlatEtcetera" runat="server">You must enter the flat, building number or building name.</asp:label>
      </td>
    </tr>
    </asp:panel>
    
    <asp:panel runat="server"  visible="False" id="pnlAddressBuildingName" >  
    <tr>
      <td class="prompt"><span class="mandatoryfieldgroupasterisk">*</span> <asp:label  id="lblAddressBuildingName" runat="server" text="Address" /></td>
      <td>
      <asp:textbox id="txtAddressBuildingNumber" runat="server" columns="4" cssclass="text" maxlength="40" />
      <asp:textbox id="txtAddressBuildingName" runat="server" columns="39" cssclass="text" maxlength="40" />
      <asp:customvalidator runat="server" id="validateNullAddress" onservervalidate="NotNullAddressValidator" 
            validateemptytext=True  text="***" errormessage="Please enter a valid and complete address" enableclientscript="false" />
      </td>
      <td></td>
    </tr>
   </asp:panel>
  
    
    <asp:panel runat="server"  visible="False" id="pnlAddressStreet" >
    <tr>
      <td class="prompt"><asp:label class="prompt" id="lblAddressStreet" runat="server" text="Address" /></td>
      <td><asp:textbox id="txtAddressStreet" runat="server" columns="50" CssClass="text" maxlength="40" /></td>
    </tr>
    </asp:panel>
   
   <asp:panel runat="server"  visible="False" id="pnlAddressDistrict" >
    <tr>
      <td class="prompt"><asp:label class="prompt" id="lblAddressDistrict" runat="server" text="Address" /></td>
      <td><asp:textbox id="txtAddressDistrict" runat="server" columns="50" CssClass="text" maxlength="40" /></td>
    </tr>
   </asp:panel>
   
   <asp:panel runat="server"  visible="False"  id="pnlTown" > 
    <tr>
      <td  class="prompt"><span class="mandatoryfieldasterisk">*</span> <asp:label class="prompt" id="lblTown" runat="server" text="Town" /></td>
      <td>
        <asp:textbox id="txtTown" runat="server" columns="50" CssClass="text" MaxLength="40" />
        <asp:requiredfieldvalidator runat="server" controltovalidate="txtTown" text="***" errormessage="Please enter a Town"
        id="Requiredfieldvalidator1" enableclientscript="false" />
      </td>
    </tr>
    </asp:panel>
    
    <asp:panel id="pnlCounty" visible="False"  runat="server" >
    <tr>
      <td class="prompt"><asp:label class="prompt" id="lblCounty" runat="server" text="County" /></td>
      <td><asp:textbox id="txtCounty" runat="server" columns="50" CssClass="text" MaxLength="40" /></td>
    </tr>
    </asp:panel>
    
    <asp:panel id="pnlPostcode" visible="False" runat="server" >
    <tr>
      <td class="prompt"><span class="mandatoryfieldasterisk">*</span> <asp:label id="lblPostcode" runat="server" text="Post code" /></td>
      <td>
        <asp:textbox id="txtPostcode" runat="server" columns="10" CssClass="text" />
        <asp:requiredfieldvalidator runat="server" controltovalidate="txtPostcode" text="***" errormessage="Please enter a Postcode"
        id="txtPostcode_validator" EnableClientScript="false" />
      </td>
    </tr>
    </asp:panel>
    <asp:panel runat="server" visible="False" id="pnlCountry"  >
    
      <tr>
        <td class="prompt"><span class="mandatoryfieldasterisk">*</span> <asp:label id="lblCountry" runat="server" text="Country" /></td>
        <td>
          <asp:textbox id="txtCountry" runat="server" columns="50" CssClass="text" />
          <asp:requiredfieldvalidator runat="server" controltovalidate="txtCountry" text="***" errormessage="Please enter a Country"
          id="txtCountry_validator" EnableClientScript="false" />
        </td>
      </tr>
    
    </asp:Panel>
    
  </table>
</asp:Panel>
<asp:panel runat="server" id="pnlResidentFrom">
  <table class="formtable">
    <tr>
      <td class="prompt"><span class="mandatoryfieldasterisk">*</span> Resident From</td>
      <td>
        <asp:textbox id="txtResidentFrom" runat="server" columns="10" cssclass="text" autopostback="True" />
        <validators:datevalidator id="DateValidator1" runat="server" controltovalidate="txtResidentFrom" text="***" errormessage="Invalid resident from date"
          enableclientscript="false" /> dd/mm/yyyy
      </td>
    </tr>  
    <asp:panel runat="server" id="pnlResidentTo">
    <tr>
      <td class="prompt"><span class="mandatoryfieldasterisk">*</span> to</td>
      <td>
        <asp:textbox id="txtResidentTo" runat="server" columns="10" cssclass="text" />
        <validators:datevalidator id="Datevalidator2" runat="server" controltovalidate="txtResidentTo" text="***" errormessage="Invalid resident to date"
          enableclientscript="false" />
      </td>
    </tr> 
    </asp:panel> 
    <tr>
      <td></td>
      <td>If you cannot remember the exact date, please use the first of the month</td>
    </tr>
  </table>
</asp:panel>

<asp:panel runat="server" id="pnlNatureOfOccupancy">
    <table class="formtable">
    <tr>
      <td class="prompt"><span class="mandatoryfieldasterisk">*</span> Nature of Occupancy</td>
      <td>
        <asp:dropdownlist id="cmbNatureOfOccupancy" runat="server" />
        <validators:requiredselecteditemvalidator runat="server" enableclientscript="false"
          text="***" errormessage="Nature of occupancy is required" controltovalidate="cmbNatureOfOccupancy" id="Requiredselecteditemvalidator1"/>
       </td>
    </tr>
  </table>
</asp:Panel> 
  