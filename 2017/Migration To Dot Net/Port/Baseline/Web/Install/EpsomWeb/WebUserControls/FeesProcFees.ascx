<%@ Control Language="c#" AutoEventWireup="false" Codebehind="FeesProcFees.ascx.cs" Inherits="Epsom.Web.WebUserControls.FeesProcFees" TargetSchema="http://schemas.microsoft.com/intellisense/ie5"%>
<%@ Register TagPrefix="CMS" namespace="Epsom.Web.Helpers" assembly="Epsom.Web" %>



<table>
  <tr>
    <td></TD>
    <td><b>Amount to be paid</b></TD>
    <td><b>When payable</b></TD>
        <td><b>Refundable</b></TD>
        </TR>
        <asp:repeater id=rptFees runat="server">
    <itemtemplate>
      <tr>
        <td class="prompt"">
          <asp:label id="lblIdentifier" runat="server" /></td>
        <td width="100px">
          <asp:textbox id="txtAmount" runat="server" CssClass="text" /></td>
        <td>
          <asp:dropdownlist id="ddlWhenPayable" runat="server" /></td>
         <td>
         <asp:checkbox id="chbRefundable" runat="server"/> 
         </td>
      </tr>
      
    </itemtemplate>
  </asp:repeater>
  
  </table>