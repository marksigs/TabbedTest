<%@ Control Language="c#" AutoEventWireup="false" Codebehind="TelephoneNumberDisplay.ascx.cs" Inherits="Epsom.Web.WebUserControls.TelephoneNumberDisplay" TargetSchema="http://schemas.microsoft.com/intellisense/ie5"%>
<table class=displaytable>
    <tr>
      <td class="prompt"> </td>
      <td>Area</td> 
      <td>Number</td>
      <td>Ext</td>
      <td><asp:label id="lblPreferredHeading" runat="server" visible="False"></asp:label></td>
    </tr> 
    
    <asp:repeater id="rptTelephoneNumbers" runat="server">
      <itemtemplate>
      <asp:panel id="pnlTelephoneNumber" runat="server" >
      <tr>
        <td class="prompt"><asp:label id="lblTelephoneDescription" runat="server"></asp:label></td>
        <td><asp:label id="lblTelephoneAreaCode" runat="server" columns ="10"></asp:label></td> 
        <td><asp:label id="lblTelephoneNumber" runat="server" columns ="10"></asp:label></td>
        <td><asp:label id="lblTelephoneExtension" runat="server" columns ="10"></asp:label></td>
        <td>
          <asp:radiobutton visible="False"
            groupname="grpTelephonePreferredNumber" id="rdoTelephoneHomePreferred" text="" runat="server" enabled="False" >
          </asp:radiobutton>
        </td>
      </tr> 
      </asp:panel> 
      </itemtemplate>
    </asp:repeater>
</table>