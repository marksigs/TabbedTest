<%@ Page language="c#" Codebehind="ApplicantDetails.aspx.cs" AutoEventWireup="false" 
    SmartNavigation="false" Inherits="Epsom.Web.Kfi.ApplicantDetails" %>
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

    <h1><asp:Label id="lblHeading" runat="server" /></h1>
    
    
    <epsom:dipnavigationbuttons id="ctlDIPNavigationButtons1" runat="server" />
	
	<cms:cmsBoilerplate cmsref="60" runat="server" />

    <asp:validationsummary runat="server" displaymode="BulletList" headertext="Input errors:" CssClass="validationsummary" showsummary="true" id="Validationsummary1" />

		<epsom:applicant id="ctlApplicant" runat="server" FormMode="KFI" />

    <br />
				
    <p class="button_orphan">
  		<asp:Button ID="cmdAddApplicant" runat="server" text="Add another applicant" />
    </p>

    <epsom:dipnavigationbuttons buttonbar="true" id="ctlDIPNavigationButtons2" runat="server" />

    <sstchur:SmartScroller id="ctlSmartScroller" runat = "server" />

  </mp:content>

</mp:contentcontainer>

