<%@ Page language="c#" Codebehind="ViewDocument.aspx.cs" AutoEventWireup="false" Inherits="Epsom.Web.DocumentViewer.ViewDocument" %>
<%@ Register TagPrefix="mp" namespace="Microsoft.Web.Samples.MasterPages" assembly="MasterPages" %>
<%@ Register TagPrefix = "sstchur" Namespace = "sstchur.web.SmartNav" Assembly = "sstchur.web.SmartNav" %>
<%@ Register Src="~/WebUserControls/DIPMenu.ascx" TagName="DIPMenu" TagPrefix="Epsom" %>

<epsom:dipmenu id="DipMenu" runat="server" />

<mp:contentcontainer id="Contentcontainer1" runat="server" masterpagefile="../masterpages/masterpage1.ascx">

  <mp:content id="region1" runat="server">
  
    <h1>View document</h1>

    <asp:Panel id="pnlDownloadDocument" runat="server" visible="False">

      <p>
        To view the document you will need <b>Adobe Acrobat Reader</b> installed.
        Please click the link on this page to download the reader.
      </p>

      <br />

      <p>
        Click on the button below to view the document.
      </p>

      <br />

      <asp:Button id="cmdDownloadDocument" runat="server" text="View document" CssClass="button" />

    </asp:Panel>

    <asp:Panel id="pnlError" runat="server" visible="False">

      <p>
        The document is not available for download. Please see the error below.
      </p>

      <br />

      <p class="error">
  		  <asp:Literal id="ltlError" runat="server" />
  		</p>

    </asp:Panel>

    <p>
		  <asp:Label id="DebugLabel" runat="server" visible="False" />
    </p>

  </mp:content>
 
</mp:contentcontainer>
  
