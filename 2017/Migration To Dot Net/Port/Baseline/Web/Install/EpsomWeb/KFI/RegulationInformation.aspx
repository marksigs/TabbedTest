<%@ Register Src="~/WebUserControls/DIPMenu.ascx" TagName="DIPMenu" TagPrefix="Epsom" %>
<%@ Register Src="~/WebUserControls/DIPNavigationButtons.ascx" TagName="DIPNavigationButtons" TagPrefix="Epsom" %>
<%@ Register TagPrefix="mp" namespace="Microsoft.Web.Samples.MasterPages" assembly="MasterPages" %>
<%@ Page language="c#" Codebehind="RegulationInformation.aspx.cs" AutoEventWireup="false" 
  SmartNavigation="false" Inherits="Epsom.Web.Kfi.RegulationInformation" %>
<%@ Register TagPrefix="CMS" namespace="Epsom.Web.Helpers" assembly="Epsom.Web" %>

<mp:contentcontainer id="Contentcontainer1" runat="server" masterpagefile="../masterpages/masterpage2.ascx">

  <mp:content id=region1 runat="server">
    <epsom:dipmenu id=DipMenu runat="server"></epsom:dipmenu>
  </mp:content>

  <mp:content id=region2 runat="server">

    <cms:helplink id="helplink" class="helplink" runat="server" helpref="2200" />

    <h1>Regulation Information</h1>

    <epsom:dipnavigationbuttons id=ctlDIPNavigationButtons1 runat="server"></epsom:dipnavigationbuttons>

    <asp:Repeater id=rptRegulationInformationCollection Runat="server">
	      <itemtemplate>
	        <div>
	        <asp:Button Runat="server" ID="cmdExpand" commandName="cmdExpand" Text="+" BorderStyle="None" BackColor="White"></asp:Button>
	        <asp:Label ID="lblSummary" Runat="server"></asp:Label>
	        <asp:Panel ID="pnlDetail" Runat="server" Visible="False">
	        <blockquote>
	          <asp:Label ID="lblDetail" Runat="server"></asp:Label><br>
	          Last Updated : 
            <asp:Label ID="lblLastUpdated" Runat="server"></asp:Label>
	        </blockquote>
	        </asp:Panel>
	        </div>
	      </itemtemplate>
    </asp:Repeater>

    <epsom:dipnavigationbuttons buttonbar="true" id=ctlDIPNavigationButtons2 runat="server"></epsom:dipnavigationbuttons>
  
  </mp:content>

</mp:contentcontainer>
