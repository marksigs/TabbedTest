<%@ Control Language="c#" AutoEventWireup="false" Codebehind="DipNavigationButtons.ascx.cs" Inherits="Epsom.Web.WebUserControls.DipNavigationButtons" TargetSchema="http://schemas.microsoft.com/intellisense/ie5"%>

<asp:Panel id="navigationButtons" runat="server">

  <span style="left:1.5em; position:absolute">
  	<asp:button id="cmdSave" runat="server" text="Save" CssClass="button" />
	  &nbsp;
	  <asp:button id="cmdSaveAndClose" runat="server" text="Stop &amp; Save" CssClass="button" />
  </span>

  <asp:button id="cmdCancelPage" runat="server" text="Cancel" cssclass="button" CausesValidation="False" />
  &nbsp;
	<asp:button id="cmdPreviousPage" runat="server" text="Previous" CssClass="button" />
  &nbsp;
	<asp:button id="cmdNextPage" runat="server" text="Next" CssClass="button" />
  &nbsp;
  <asp:button id="cmdSubmit" runat="server" text="Submit" CssClass="button" />
   
</asp:Panel>
