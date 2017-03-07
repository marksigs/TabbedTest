<%@ Control Language="c#" AutoEventWireup="false" Codebehind="SelectValidAddress.ascx.cs" Inherits="Epsom.Web.WebUserControls.SelectValidAddress" TargetSchema="http://schemas.microsoft.com/intellisense/ie5"%>
<%@ Register Src="~/WebUserControls/AddressDisplay.ascx" TagName="AddressDisplay" TagPrefix="Epsom" %>
<%@ Register TagPrefix="validators" NameSpace="Epsom.Web.Validators" Assembly="Epsom.Web"%>

<asp:repeater id="rptInvalidAddresses" runat="server">

  <itemtemplate>

    <fieldset>

      <legend><asp:label id="lblAddressType" runat="server" /></legend>

      <epsom:AddressDisplay id="ctlInvalidAddress" runat="server" />


      <table class="formtable">
        <tr>
          <td class="prompt">Your reference</td>
          <td><asp:label id="lblYourReference" text="" runat="server"></asp:label>
          </td>
        </tr>
        <tr>
          <td class="prompt">Credit agency reference</td>
          <td><asp:label id="lblCreditAgencyReference" text="" runat="server"></asp:label>
          </td>
        </tr>
      </table> 
      <table class="formtable">
        <tr>
          <td><b>Select alternative address</b></td>
        </tr>
        <tr>
          <td>
            <asp:listbox id="lstTargetAddresses" runat="server" autopostback="false" rows="10" style="width: 98%" />
            <validators:requiredselecteditemvalidator id="RequiredSelectedItemValidator1" runat="server" enableclientscript="false" text="***" 
              errormessage="Please select address for each item" controltovalidate="lstTargetAddresses"></validators:requiredselecteditemvalidator>
          </td> 
        </tr>
      </table>

    </fieldset> 

  </itemtemplate> 

</asp:repeater> 
