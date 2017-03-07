<%@ Page language="c#" Codebehind="PersonalDetails.aspx.cs" AutoEventWireup="false" Inherits="Epsom.Web.App.ApplicantDetails" %>
<%@ Register Src="~/WebUserControls/Applicant.ascx" TagName="Applicant" TagPrefix="Epsom" %>
<%@ Register TagPrefix="mp" namespace="Microsoft.Web.Samples.MasterPages" assembly="MasterPages" %>
<%@ Register Src="~/WebUserControls/DIPNavigationButtons.ascx" TagName="DIPNavigationButtons" TagPrefix="Epsom" %>
<%@ Register Src="~/WebUserControls/DIPMenu.ascx" TagName="DIPMenu" TagPrefix="Epsom" %>
<%@ Register TagPrefix = "sstchur" Namespace = "sstchur.web.SmartNav" Assembly = "sstchur.web.SmartNav" %>
<%@ Register TagPrefix="CMS" namespace="Epsom.Web.Helpers" assembly="Epsom.Web" %>

<mp:contentcontainer id="Contentcontainer1" runat="server" masterpagefile="../masterpages/masterpage2.ascx">

  <mp:content id="region1" runat="server">

    <epsom:dipmenu id="DipMenu" runat="server" />

  </mp:content>

  <mp:content id="region2" runat="server">
	<cms:helplink id="helplink" class="helplink" runat="server" helpref="2060" />
    <h1><asp:label id="lblHeading" runat="server" /></h1>

    
    
    <epsom:dipnavigationbuttons id="ctlDIPNavigationButtons1" runat="server" />
	
	<cms:cmsBoilerplate cmsref="60" runat="server" />
    
    <asp:validationsummary runat="server" displaymode="BulletList" headertext="Input errors:" cssclass="validationsummary" showsummary="true" id="Validationsummary1" />

		<epsom:applicant id="ctlApplicant" runat="server" formmode="APP" />
			
		<br />	
				
    <p class="button_orphan">
		  <asp:button id="cmdAddApplicant" runat="server" text="Add another applicant" enabled="false"/>
    </p>

    <epsom:dipnavigationbuttons buttonbar="true" savepage="true" saveandclosepage="true" id="ctlDIPNavigationButtons2" runat="server" />

    <sstchur:smartscroller id="ctlSmartScroller" runat = "server" />

  </mp:content>

</mp:contentcontainer>

