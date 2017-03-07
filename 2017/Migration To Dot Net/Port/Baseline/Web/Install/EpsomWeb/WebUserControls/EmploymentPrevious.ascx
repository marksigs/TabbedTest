<%@ Control Language="c#" AutoEventWireup="false" Codebehind="EmploymentPrevious.ascx.cs" Inherits="Epsom.Web.WebUserControls.EmploymentPrevious" TargetSchema="http://schemas.microsoft.com/intellisense/ie5"%>
<%@ Register Src="~/WebUserControls/Address.ascx" TagName="Address" TagPrefix="Epsom" %>
<%@ Register Src="~/WebUserControls/TelephoneNumbers.ascx" TagName="TelephoneNumbers" TagPrefix="Epsom" %>
<asp:panel id="pnlEmploymentHistory" runat="server">
  <H2 class="form_heading">Previous employment details</H2>
  <TABLE class="formtable">
    <TR>
      <TD class="prompt">
        <asp:label id="lblPreviousEmployerName" runat="server" text="Name of previous employer "></asp:label>
      </TD>
      <td><asp:textbox columns="40" cssclass="text" id="txtPreviousEmployerName" runat="server" tooltip="enter the name of your previous employer"></asp:textbox>
      </TD>
    </TR>
</table>
<epsom:telephonenumbers id="ctlTelephoneNumbers" runat="server" />
<epsom:address id="ctlAddress" runat="server"></epsom:address>
<table class="formtable">
    <TR>
      <TD class="prompt">
        <asp:label id="lblPreviousEmployerBusinessNature" runat="server" text="Nature of business"></asp:label></TD>
      <TD>
        <asp:textbox  cssclass="text" columns="40" id="txtPreviousEmployerBusinessNature" runat="server" tooltip="enter the nature of your previous employer's business"></asp:textbox></TD>
    </TR>
    <TR>
      <TD class="prompt">
        <asp:label id="lblPreviousPosition" runat="server" text="Position held"></asp:label></TD>
      <TD>
        <asp:textbox  cssclass="text" columns="40"  id="txtPreviousPosition" runat="server" tooltip="enter the position you held in your last employment"></asp:textbox></TD>
    </TR>
    <TR>
      <TD class="prompt">
        <asp:label id="lblPreviousEmploymentStartDate" runat="server" text="Start date of employment"></asp:label></TD>
      <TD>
        <asp:textbox cssclass="text" id="txtPreviousEmploymentStartDate" runat="server" tooltip="enter the strat date of your previous employment"
          columns="10"></asp:textbox>
     </TD>
    <TR>
      <TD class="prompt">
        <asp:label id="lblPreviousEmploymentEndDate" runat="server" text="End date of employment"></asp:label></TD>
      <TD>
        <asp:textbox  cssclass="text" id="txtPreviousEmploymentEndDate" runat="server" tooltip="enter the end date of your previous employment"
          columns="10"></asp:textbox>
        </TD>
    </TR>
  </TABLE>
</asp:panel>
