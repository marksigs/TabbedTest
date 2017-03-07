<%@ Control Language="c#" AutoEventWireup="false" Codebehind="Documents.ascx.cs" Inherits="Epsom.Web.WebUserControls.Documents" TargetSchema="http://schemas.microsoft.com/intellisense/ie5"%>

<asp:repeater id="rptDocuments" runat="server">
  <itemtemplate>
    <p>
      <asp:LinkButton id="btnDownloadDocument" CssClass="fancylink" runat="server" />
      <input id="hiddenIndex" type="hidden" runat="server" name="hiddenIndex"/>
    </p>
  </itemtemplate>
</asp:repeater>
