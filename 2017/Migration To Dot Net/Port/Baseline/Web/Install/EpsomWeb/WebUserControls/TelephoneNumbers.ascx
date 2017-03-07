<%@ Control Language="c#" AutoEventWireup="false" Codebehind="TelephoneNumbers.ascx.cs" Inherits="Epsom.Web.WebUserControls.TelephoneNumbers" TargetSchema="http://schemas.microsoft.com/intellisense/ie5"%>

<asp:Panel ID="pnlTelephoneNumbers" Runat="server">

  <table class="formtable">

    <tr style="text-align: left">
      <th>Telephone</th>
      <th>Area</th> 
      <th>Number</th>
      <th>Ext</th>
      <th><asp:label Runat="server" ID="lblPreferredContact">Preferred?</asp:label></th>
    </tr> 

    <asp:Repeater ID="rptTelephoneNumbers" Runat="server">
      <ItemTemplate>
        <tr style="text-align: left">
          <td class="prompt">
            <span class="mandatoryfieldasterisk"><asp:Literal runat="server" id="lblMandatory">* </asp:Literal></span>
            <asp:Label ID="lblTelephoneDescription" Runat="server"></asp:Label>
            
          </td>
          <td>
            <asp:textbox id="txtTelephoneAreaCode" runat="server" columns="8"  maxlength="6" CssClass="text"></asp:textbox>
            <asp:requiredfieldvalidator runat="server" controltovalidate="txtTelephoneAreaCode" text="***" errormessage="Please enter the telephone area code" id="requiredAreaCode" enableclientscript="false" />
          </td> 
          <td>
            <asp:textbox id="txtTelephoneNumber" runat="server" columns="25" maxlength="30" CssClass="text"></asp:textbox>
            <asp:requiredfieldvalidator runat="server" controltovalidate="txtTelephoneNumber" text="***" errormessage="Please enter the telephone number" id="requiredNumber" enableclientscript="false" />
          </td>
          <td><asp:textbox id="txtTelephoneExtension" runat="server" columns="8" maxlength="10" CssClass="text"></asp:textbox></td>
          <td><asp:checkbox id="rdoTelephonePreferred" text="" runat="server" /></td>
        </tr> 
      </ItemTemplate>
    </asp:Repeater>

  </table>

</asp:Panel>