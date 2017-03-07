<%@ Control Language="c#" AutoEventWireup="false" Codebehind="EmploymentContractDetails.ascx.cs" Inherits="Epsom.Web.WebUserControls.EmploymentContractDetails" TargetSchema="http://schemas.microsoft.com/intellisense/ie5"%>

  <H2 class="form_heading">Previous employment</H2>
  <TABLE class="formtable">
    <TR>
      <TD class="prompt" TD>
        <asp:label id="Label22" runat="server" text="Name of previous employer "></asp:label><span class="mandatoryfieldasterisk">*</span></TD>
      <TD>
        <asp:textbox cssclass="text" id="Textbox21" runat="server"></asp:textbox>
        <asp:button id="Button2" runat="server" Text="Validate Address"></asp:button></TD>
    </TR>

    <TR>
      <TD class="prompt"></TD>
      <TD>
        <asp:TextBox cssclass="text" id="Textbox13" runat="server" Width="400px" Height="100px"></asp:TextBox></TD>
    <TR>

    <TR>
      <TD class="prompt"></TD>
      <TD>
        <asp:TextBox cssclass="text" id="txtAddress" runat="server" Width="400px" Height="100px"></asp:TextBox></TD>
    <TR>

    <TR>
      <TD class="prompt">
        <asp:label id="lblContractLength" runat="server" text="Length of contract*"></asp:label></TD>
      <TD>
        <asp:dropdownlist id="ddlContractLengthYears" runat="server"></asp:dropdownlist>years
        <asp:dropdownlist id="Dropdownlist1" runat="server"></asp:dropdownlist>months</TD>
    </TR>
    <TR>
      <TD class="prompt">
        <asp:label id="lblContractEndDate" runat="server" text="Date the contract is due to finish *"></asp:label></TD>
      <TD>
        <asp:textbox  cssclass="text" id="txtContractEndDate" runat="server" text="" columns="10" tooltip="date the contract is due to finish (dd/mm/yyyy)"></asp:textbox>dd/mm/yyyy
    </TD>
    </TR>
    <TR>
      <TD class="prompt">
        <asp:label id="lblContractStartDate" runat="server" text="Date the contract started *"></asp:label></TD>
      <TD>
        <asp:textbox  cssclass="text" id="txtContractStartDate" runat="server" Text="" columns="10" tooltip="date the contract started (dd/mm/yyyy)"></asp:textbox>dd/mm/yyyy
        </TD>
    </TR>
    <TR>
      <TD class="prompt">
        <asp:label id="lblRenewableContract" runat="server" text="Is the contract renewable?"></asp:label></TD>
      <TD>
        <asp:RadioButtonList id="rblRenewableContract" runat="Server" tooltip="state whether your contract is renewable or not">
          <asp:ListItem value="1" runat="server" ID="Listitem1" NAME="Listitem1">Yes</asp:ListItem>
          <asp:ListItem value="2" runat="server" ID="Listitem2" NAME="Listitem2">No</asp:ListItem>
        </asp:RadioButtonList></TD>
    </TR>
    <TR>
      <TD class="prompt">
        <asp:label id="lblNoTimesRenewed" runat="server" text="How many times has the contract been renewed?"></asp:label></TD>
      <TD>
        <asp:textbox  cssclass="text" id="txtNoTimesRenewed" runat="server" text="" tooltip="enter the number of times your contract has been renewed"></asp:textbox></TD>
    </TR>
  </TABLE>

