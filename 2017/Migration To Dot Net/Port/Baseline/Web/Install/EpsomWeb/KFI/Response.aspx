<%@ Page language="c#" Codebehind="Response.aspx.cs" AutoEventWireup="false" 
   SmartNavigation="false" Inherits="Epsom.Web.Kfi.Response" %>
<%@ Register Src="~/WebUserControls/DIPMenu.ascx" TagName="DIPMenu" TagPrefix="Epsom" %>
<%@ Register Src="~/WebUserControls/DIPNavigationButtons.ascx" TagName="DIPNavigationButtons" TagPrefix="Epsom" %>
<%@ Register TagPrefix="mp" namespace="Microsoft.Web.Samples.MasterPages" assembly="MasterPages" %>
<%@ Register TagPrefix="CMS" namespace="Epsom.Web.Helpers" assembly="Epsom.Web" %>

<mp:contentcontainer id="Contentcontainer1" masterpagefile="../masterpages/masterpage2.ascx" runat="server">

  <mp:content id="region1" runat="server">
    <epsom:dipmenu id=DipMenu runat="server"></epsom:dipmenu>
  </mp:content>
  
  <mp:content id="region2" runat="server">
	<cms:helplink id="helplink" class="helplink" runat="server" helpref="40900" />
    <h1>Response: Key facts illustration</h1>

      <table class="formtable">
        <tr>
          <td class="promptblack">Case reference number</td>
          <td><asp:label id="lblCaseReferenceNumber" runat="server" /></td>
          <td class="promptblack">Initial monthly payment</td>
        </tr> 
        <tr>
        
          <td width="50px">
          <asp:linkbutton id="cmdDownloadDocumentImageLink" runat="server">
          <img runat="server" id="pdfImage" src="../_gfx/pdf_logo.gif" ></img></asp:linkbutton> 
          </td>
          <td>
            <cms:cmsBoilerplate cmsref="140" runat="server" /></br>
            <asp:linkbutton id="cmdDownloadDocumentTextLink" runat="server">[Download-32Kb]</asp:linkbutton></a>
            <asp:linkbutton id="cmdDownloadDocument2" runat="server" visible="false" >[Download 2nd method-32Kb]</asp:linkbutton></a>
            <asp:linkbutton id="cmdDownloadDocument3" runat="server" visible="false" >[Download 3rd method-32Kb]</asp:linkbutton></a>
          </td>
          <td class="promptblack"><asp:Label id="lblMonthlyCost" runat="server" /></td>
        </tr> 
        <tr>
          <td colspan="2">
            <cms:cmsBoilerplate cmsref="141" runat="server" />
          </td>
        </tr> 
      </table>
  
    <p class="dipnavigationbuttonswithbar">
      <asp:button id="cmdCancel" runat="server" text="Done" />
	    &nbsp;
	    <asp:button id="cmdPrePopulateKfi" runat="server" text="Prepop Kfi" />
	    &nbsp;
	    <asp:button id="cmdPrePopulateDip" runat="server" text="Get DIP" />
    </p>
     
  </mp:content>

</mp:contentcontainer>
